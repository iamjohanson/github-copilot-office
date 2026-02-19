import { useState, useEffect, useRef } from "react";
import {
  FluentProvider,
  webLightTheme,
  webDarkTheme,
  makeStyles,
} from "@fluentui/react-components";
import { ChatInput, ImageAttachment } from "./components/ChatInput";
import { Message, MessageList } from "./components/MessageList";
import { HeaderBar, ModelType } from "./components/HeaderBar";
import { SessionHistory } from "./components/SessionHistory";
import { useIsDarkMode } from "./useIsDarkMode";
import { useLocalStorage } from "./useLocalStorage";
import { createWebSocketClient } from "./lib/websocket-client";
import { getToolsForHost } from "./tools";
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

export const App: React.FC = () => {
  const styles = useStyles();
  const [availableModels, setAvailableModels] = useState(FALLBACK_MODELS);
  const [messages, setMessages] = useState<Message[]>([]);
  const [inputValue, setInputValue] = useState("");
  const [images, setImages] = useState<ImageAttachment[]>([]);
  const [isTyping, setIsTyping] = useState(false);
  const [currentActivity, setCurrentActivity] = useState<string>("");
  const [streamingText, setStreamingText] = useState<string>("");
  const [session, setSession] = useState<any>(null);
  const [client, setClient] = useState<any>(null);
  const [error, setError] = useState("");
  const [selectedModel, setSelectedModel] = useLocalStorage<ModelType>("word-addin-selected-model", "");
  const [showHistory, setShowHistory] = useState(false);
  const [currentSessionId, setCurrentSessionId] = useState<string>("");
  const [officeHost, setOfficeHost] = useState<OfficeHost>("word");
  const isDarkMode = useIsDarkMode();
  
  // Track session creation time
  const sessionCreatedAt = useRef<string>("");

  // Fetch available models from server
  useEffect(() => {
    fetch("/api/models")
      .then((r) => r.json())
      .then((data) => {
        if (data.models?.length) {
          const models = data.models.map((id: string) => ({ key: id, label: modelIdToLabel(id) }));
          setAvailableModels(models);
          // Set default model if none stored yet
          if (!selectedModel) {
            setSelectedModel(pickDefaultModel(models));
          }
        }
      })
      .catch(() => {
        // Use fallback models
        if (!selectedModel) {
          setSelectedModel(pickDefaultModel(FALLBACK_MODELS));
        }
      });
  }, []);

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
    
    try {
      if (client) {
        await client.stop();
      }
      const host = Office.context.host;
      setOfficeHost(getHostFromOfficeHost(host));
      const tools = getToolsForHost(host);
      const newClient = await createWebSocketClient(`wss://${location.host}/api/copilot`);
      setClient(newClient);
      
      // Build host-specific system message
      const hostName = host === Office.HostType.PowerPoint ? "PowerPoint" 
        : host === Office.HostType.Word ? "Word" 
        : host === Office.HostType.Excel ? "Excel" 
        : "Office";
      
      const systemMessage = {
        mode: "append" as const,
        content: `You are an AI assistant embedded inside Microsoft ${hostName} as an Office Add-in. You have direct access to the open ${hostName} document through the tools provided.

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
      
      setSession(await newClient.createSession({ model, tools, systemMessage }));
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

      for await (const event of session.query({ 
        prompt: userInput || "Here are some images for you to analyze.",
        attachments: attachments.length > 0 ? attachments : undefined
      })) {
        console.log('[event]', event.type, event);
        
        if (event.type === 'assistant.message.delta') {
          // Streaming text chunk
          const delta = (event.data as any).delta || (event.data as any).content || '';
          setStreamingText(prev => prev + delta);
          setCurrentActivity("");
        } else if (event.type === 'assistant.message' && (event.data as any).content) {
          // Complete message - add to messages and clear streaming
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
        } else if (event.type === 'tool.execution_end') {
          setCurrentActivity("Processing result...");
        } else if (event.type === 'assistant.thinking') {
          setCurrentActivity("Thinking...");
        } else if (event.type === 'assistant.turn_start') {
          setCurrentActivity("Starting response...");
        } else if (event.type === 'assistant.turn_end') {
          setCurrentActivity("");
          setStreamingText("");
          console.log('[turn_end]', (event.data as any).stopReason);
        }
      }
      console.log('[query complete]');
    } catch (e: any) {
      setError(e.message || 'Unknown error');
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

  return (
    <FluentProvider theme={isDarkMode ? webDarkTheme : webLightTheme}>
      <div className={styles.container}>
        <HeaderBar 
          onNewChat={() => startNewSession(selectedModel)} 
          onShowHistory={() => setShowHistory(true)}
          selectedModel={selectedModel}
          onModelChange={handleModelChange}
          models={availableModels}
        />

        <MessageList
          messages={messages}
          isTyping={isTyping}
          isConnecting={!session && !error}
          currentActivity={currentActivity}
          streamingText={streamingText}
        />

        {error && <div style={{ color: 'red', padding: '8px' }}>{error}</div>}

        <ChatInput
          value={inputValue}
          onChange={setInputValue}
          onSend={handleSend}
          images={images}
          onImagesChange={setImages}
        />
      </div>
    </FluentProvider>
  );
};
