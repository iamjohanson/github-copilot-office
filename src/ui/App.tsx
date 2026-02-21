import { useState, useEffect, useRef, useCallback } from "react";
import {
  FluentProvider,
  webLightTheme,
  webDarkTheme,
  makeStyles,
} from "@fluentui/react-components";
import { ChatInput, ImageAttachment } from "./components/ChatInput";
import { Message, MessageList, DebugEvent } from "./components/MessageList";
import { HeaderBar, ModelType } from "./components/HeaderBar";
import { SessionHistory } from "./components/SessionHistory";
import { PermissionDialog, PermissionDecision } from "./components/PermissionDialog";
import { PermissionManager } from "./components/PermissionManager";
import { useIsDarkMode } from "./useIsDarkMode";
import { useLocalStorage } from "./useLocalStorage";
import { createWebSocketClient, PermissionRequest, PermissionResult, ModelInfo } from "./lib/websocket-client";
import { PermissionService } from "./lib/permissionService";
import { getToolsForHost } from "./tools";
import { remoteLog } from "./lib/remoteLog";
import { trafficStats } from "./lib/websocket-transport";
import { 
  SavedSession, 
  OfficeHost, 
  saveSession, 
  generateSessionTitle, 
  getHostFromOfficeHost 
} from "./sessionStorage";
import React from "react";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    backgroundColor: "var(--colorNeutralBackground3)",
  },
});

const FALLBACK_MODELS = [
  { key: "claude-sonnet-4.5", label: "Claude Sonnet 4.5" },
];

function modelIdToLabel(id: string): string {
  return id
    .split("-")
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1))
    .join(" ");
}

function pickDefaultModel(models: { key: string }[]): ModelType {
  const preferred = ["claude-sonnet-4.6", "claude-sonnet-4.5"];
  for (const id of preferred) {
    if (models.some((m) => m.key === id)) return id;
  }
  return models[0]?.key || "claude-sonnet-4.5";
}

const permissionService = new PermissionService();

export const App: React.FC = () => {
  const styles = useStyles();
  const [availableModels, setAvailableModels] = useState(FALLBACK_MODELS);
  const [messages, setMessages] = useState<Message[]>([]);
  const [inputValue, setInputValue] = useState("");
  const [images, setImages] = useState<ImageAttachment[]>([]);
  const [isTyping, setIsTyping] = useState(false);
  const [currentActivity, setCurrentActivity] = useState<string>("");
  const [streamingText, setStreamingText] = useState<string>("");
  const [debugEvents, setDebugEvents] = useState<DebugEvent[]>([]);
  const [session, setSession] = useState<any>(null);
  const [client, setClient] = useState<any>(null);
  const [error, setError] = useState("");
  const [selectedModel, setSelectedModel] = useLocalStorage<ModelType>("word-addin-selected-model", "");
  const [showHistory, setShowHistory] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [currentSessionId, setCurrentSessionId] = useState<string>("");
  const [officeHost, setOfficeHost] = useState<OfficeHost>("word");
  const [cwd, setCwd] = useLocalStorage<string>("copilot-cwd", "");
  const [allowAll, setAllowAll] = useState(permissionService.allowAll);
  const [permRules, setPermRules] = useState(permissionService.getRules());
  const isDarkMode = useIsDarkMode();

  // Permission prompt state
  const [pendingPermission, setPendingPermission] = useState<{
    request: PermissionRequest;
    resolve: (result: PermissionResult) => void;
  } | null>(null);
  
  // Track session creation time
  const sessionCreatedAt = useRef<string>("");

  // Keep permissionService.cwd in sync
  useEffect(() => {
    permissionService.cwd = cwd || null;
  }, [cwd]);

  // Permission handler called by the WebSocket client
  const handlePermissionRequest = useCallback(
    (request: PermissionRequest): Promise<PermissionResult> => {
      // Try auto-evaluation first
      const autoResult = permissionService.evaluate(request);
      if (autoResult) return Promise.resolve(autoResult);

      // Prompt the user
      return new Promise<PermissionResult>((resolve) => {
        setPendingPermission({ request, resolve });
      });
    },
    [],
  );

  const handlePermissionDecision = useCallback(
    (decision: PermissionDecision) => {
      if (!pendingPermission) return;
      const { request, resolve } = pendingPermission;

      if (decision === "always") {
        // Save a rule for this kind + path
        const pathPrefix = request.path || request.fileName || cwd || "/";
        permissionService.addRule({ kind: request.kind, pathPrefix });
        setPermRules(permissionService.getRules());
        resolve({ kind: "approved" });
      } else if (decision === "allow") {
        resolve({ kind: "approved" });
      } else {
        resolve({ kind: "denied-interactively-by-user" });
      }
      setPendingPermission(null);
    },
    [pendingPermission, cwd],
  );

  // Fetch available models from CLI via models.list RPC (or fallback to /api/models)
  const fetchModels = useCallback(async (wsClient: any) => {
    try {
      const models: ModelInfo[] = await wsClient.listModels();
      if (models?.length) {
        const mapped = models.map((m: ModelInfo) => ({ key: m.id, label: m.name || modelIdToLabel(m.id) }));
        setAvailableModels(mapped);
        if (!selectedModel) {
          setSelectedModel(pickDefaultModel(mapped));
        }
        return;
      }
    } catch {
      // listModels not supported by this CLI version, fall back
    }
    // Fallback: server-side /api/models
    try {
      const r = await fetch("/api/models");
      const data = await r.json();
      if (data.models?.length) {
        const mapped = data.models.map((id: string) => ({ key: id, label: modelIdToLabel(id) }));
        setAvailableModels(mapped);
        if (!selectedModel) {
          setSelectedModel(pickDefaultModel(mapped));
        }
      }
    } catch {
      if (!selectedModel) {
        setSelectedModel(pickDefaultModel(FALLBACK_MODELS));
      }
    }
  }, [selectedModel, setSelectedModel]);

  // Save session whenever messages change (debounced effect)
  useEffect(() => {
    if (messages.length === 0 || !currentSessionId) return;
    
    // Only save if there's at least one user message
    const hasUserMessage = messages.some(m => m.sender === "user");
    if (!hasUserMessage) return;

    const savedSession: SavedSession = {
      id: currentSessionId,
      title: generateSessionTitle(messages),
      model: selectedModel,
      messages: messages,
      createdAt: sessionCreatedAt.current,
      updatedAt: new Date().toISOString(),
    };
    
    saveSession(officeHost, savedSession);
  }, [messages, currentSessionId, selectedModel, officeHost]);

  const startNewSession = async (model: ModelType, restoredMessages?: Message[]) => {
    // Generate new session ID
    const newSessionId = crypto.randomUUID();
    setCurrentSessionId(newSessionId);
    sessionCreatedAt.current = new Date().toISOString();
    
    setMessages(restoredMessages || []);
    setInputValue("");
    setImages([]);
    setIsTyping(false);
    setCurrentActivity("");
    setStreamingText("");
    setError("");
    setShowHistory(false);
    setShowSettings(false);
    
    try {
      if (client) {
        await client.stop();
      }
      const host = Office.context.host;
      setOfficeHost(getHostFromOfficeHost(host));
      const tools = getToolsForHost(host);
      const newClient = await createWebSocketClient(`wss://${location.host}/api/copilot`);
      setClient(newClient);

      // Fetch models via RPC
      fetchModels(newClient);
      
      // Build host-specific system message
      const hostName = host === Office.HostType.PowerPoint ? "PowerPoint" 
        : host === Office.HostType.Word ? "Word" 
        : host === Office.HostType.Excel ? "Excel" 
        : "Office";
      
      const systemMessage = {
        mode: "replace" as const,
        content: `You are a helpful AI assistant embedded inside Microsoft ${hostName} as an Office Add-in. You have direct access to the open ${hostName} document through the tools provided.

IMPORTANT: You are NOT a file system assistant. The user's document is already open in ${hostName}. Use your ${hostName} tools (like get_presentation_content, get_presentation_overview, get_slide_image, etc.) to read and modify the document directly. Do NOT search for files on disk or ask the user to provide file paths.

${host === Office.HostType.PowerPoint ? `For PowerPoint:
- Use get_presentation_overview first to see all slides and understand the deck structure
- Use get_presentation_content to read slide text (supports ranges like startIndex/endIndex for large decks)
- Use get_slide_image to capture a slide's visual design, colors, and layout
- The presentation is already open - just call the tools directly` : ''}

${host === Office.HostType.Word ? `For Word:
- Use get_document_content to read the document
- Use set_document_content to modify it
- The document is already open - just call the tools directly` : ''}

${host === Office.HostType.Excel ? `For Excel:
- Use get_workbook_info to understand the workbook structure
- Use get_workbook_content to read cell data
- The workbook is already open - just call the tools directly` : ''}

Always use your tools to interact with the document. Never ask users to save, export, or provide file paths.`
      };

      const newSession = await newClient.createSession({
        model,
        tools,
        systemMessage,
        requestPermission: false,
        workingDirectory: cwd || undefined,
      });

      // Register permission handler on the session
      newSession.registerPermissionHandler(handlePermissionRequest);
      
      setSession(newSession);
    } catch (e: any) {
      setError(`Failed to create session: ${e.message}`);
    }
  };

  const handleRestoreSession = (savedSession: SavedSession) => {
    // Restore the session with its messages and model
    setCurrentSessionId(savedSession.id);
    sessionCreatedAt.current = savedSession.createdAt;
    setSelectedModel(savedSession.model);
    startNewSession(savedSession.model, savedSession.messages);
  };

  useEffect(() => {
    if (selectedModel) {
      startNewSession(selectedModel);
    }
  }, [selectedModel === "" ? "" : "ready"]);

  const handleModelChange = (newModel: ModelType) => {
    setSelectedModel(newModel);
    startNewSession(newModel);
  };

  const handleCwdChange = (newCwd: string) => {
    setCwd(newCwd);
    permissionService.cwd = newCwd;
  };

  const handleAllowAllChange = (v: boolean) => {
    permissionService.allowAll = v;
    setAllowAll(v);
  };

  const handleRemoveRule = (index: number) => {
    permissionService.removeRule(index);
    setPermRules(permissionService.getRules());
  };

  const handleClearRules = () => {
    permissionService.clearRules();
    setPermRules([]);
  };

  const handleSend = async () => {
    if ((!inputValue.trim() && images.length === 0) || !session) return;

    // Add user message with images
    setMessages((prev) => [...prev, {
      id: crypto.randomUUID(),
      text: inputValue || (images.length > 0 ? `Sent ${images.length} image${images.length > 1 ? 's' : ''}` : ''),
      sender: "user",
      timestamp: new Date(),
      images: images.length > 0 ? images.map(img => ({ dataUrl: img.dataUrl, name: img.name })) : undefined,
    }]);
    const userInput = inputValue;
    const userImages = [...images];
    setInputValue("");
    setImages([]);
    setIsTyping(true);
    setCurrentActivity("Processing...");
    setStreamingText("");
    setDebugEvents([]);
    setError("");

    try {
      // Upload images to server and get file paths
      const attachments: Array<{ type: "file", path: string, displayName?: string }> = [];
      
      for (const image of userImages) {
        try {
          const response = await fetch('/api/upload-image', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ 
              dataUrl: image.dataUrl,
              name: image.name 
            }),
          });
          
          if (!response.ok) {
            throw new Error(`Failed to upload image: ${response.statusText}`);
          }
          
          const result = await response.json();
          attachments.push({
            type: "file",
            path: result.path,
            displayName: image.name,
          });
        } catch (uploadError: any) {
          console.error('Image upload error:', uploadError);
          setError(`Failed to upload image: ${uploadError.message}`);
        }
      }

      const addDebugMessage = (text: string) => {
        setMessages((prev) => [...prev, {
          id: `debug-${Date.now()}`,
          text,
          sender: "assistant" as const,
          timestamp: new Date(),
        }]);
      };

      let eventCount = 0;
      trafficStats.reset();
      for await (const event of session.query({ 
        prompt: userInput || "Here are some images for you to analyze.",
        attachments: attachments.length > 0 ? attachments : undefined
      })) {
        eventCount++;
        console.log('[event]', event.type, event);

        // Build debug preview
        const data = event.data as any;
        let preview = '';
        if (event.type === 'assistant.message_delta') {
          preview = (data.deltaContent || '').slice(0, 80);
        } else if (event.type === 'assistant.message') {
          preview = (data.content || '').slice(0, 80);
        } else if (event.type === 'assistant.reasoning_delta') {
          preview = (data.deltaContent || '').slice(0, 80);
        } else if (event.type === 'tool.execution_start') {
          preview = data.toolName || '';
        } else if (event.type === 'session.error') {
          preview = data.message || data.error || '';
        }
        setDebugEvents(prev => [...prev, { type: event.type, preview, timestamp: Date.now() }]);
        
        if (event.type === 'assistant.message_delta') {
          const delta = (event.data as any).deltaContent || '';
          setStreamingText(prev => prev + delta);
          setCurrentActivity("");
        } else if (event.type === 'assistant.message' && (event.data as any).content) {
          setStreamingText("");
          setCurrentActivity("");
          setMessages((prev) => [...prev, {
            id: event.id,
            text: (event.data as any).content,
            sender: "assistant",
            timestamp: new Date(event.timestamp),
          }]);
        } else if (event.type === 'tool.execution_start') {
          const toolName = (event.data as any).toolName;
          const toolArgs = (event.data as any).arguments || {};
          setCurrentActivity(`Calling ${toolName}...`);
          setMessages((prev) => [...prev, {
            id: event.id,
            text: JSON.stringify(toolArgs, null, 2),
            sender: "tool",
            toolName: toolName,
            toolArgs: toolArgs,
            timestamp: new Date(event.timestamp),
          }]);
        } else if (event.type === 'tool.execution_complete') {
          setCurrentActivity("Processing result...");
        } else if (event.type === 'assistant.reasoning' || event.type === 'assistant.reasoning_delta') {
          setCurrentActivity("Thinking...");
        } else if (event.type === 'assistant.turn_start') {
          setCurrentActivity("Starting response...");
        } else if (event.type === 'assistant.turn_end') {
          setCurrentActivity("");
          setStreamingText("");
        } else if (event.type === 'session.error') {
          const msg = (event.data as any).message || (event.data as any).error || JSON.stringify(event.data);
          addDebugMessage(`⚠️ Session error: ${msg}`);
        }
      }
      if (eventCount === 0) {
        addDebugMessage("⚠️ No events received from server. The query may have failed silently.");
      }
    } catch (e: any) {
      setMessages((prev) => [...prev, {
        id: `error-${Date.now()}`,
        text: `❌ Error: ${e.message || 'Unknown error'}`,
        sender: "assistant",
        timestamp: new Date(),
      }]);
    } finally {
      setIsTyping(false);
    }
  };

  // Show history panel
  if (showHistory) {
    return (
      <FluentProvider theme={isDarkMode ? webDarkTheme : webLightTheme}>
        <SessionHistory
          host={officeHost}
          onSelectSession={handleRestoreSession}
          onClose={() => setShowHistory(false)}
        />
      </FluentProvider>
    );
  }

  // Show settings panel
  if (showSettings) {
    return (
      <FluentProvider theme={isDarkMode ? webDarkTheme : webLightTheme}>
        <PermissionManager
          cwd={cwd || null}
          onCwdChange={handleCwdChange}
          rules={permRules}
          onRemoveRule={handleRemoveRule}
          onClearRules={handleClearRules}
          allowAll={allowAll}
          onAllowAllChange={handleAllowAllChange}
          onClose={() => setShowSettings(false)}
        />
      </FluentProvider>
    );
  }

  return (
    <FluentProvider theme={isDarkMode ? webDarkTheme : webLightTheme}>
      <div className={styles.container}>
        <HeaderBar 
          onNewChat={() => startNewSession(selectedModel)} 
          onShowHistory={() => setShowHistory(true)}
          onShowSettings={() => setShowSettings(true)}
          selectedModel={selectedModel}
          onModelChange={handleModelChange}
          models={availableModels}
          cwd={cwd || null}
          allowAll={allowAll}
        />

        <MessageList
          messages={messages}
          isTyping={isTyping}
          isConnecting={!session && !error}
          currentActivity={currentActivity}
          streamingText={streamingText}
          debugEvents={debugEvents}
        />

        {error && <div style={{ color: 'red', padding: '8px' }}>{error}</div>}

        <ChatInput
          value={inputValue}
          onChange={setInputValue}
          onSend={handleSend}
          images={images}
          onImagesChange={setImages}
        />

        {/* Permission prompt overlay */}
        {pendingPermission && (
          <PermissionDialog
            request={pendingPermission.request}
            cwd={cwd || null}
            onDecision={handlePermissionDecision}
          />
        )}
      </div>
    </FluentProvider>
  );
};
