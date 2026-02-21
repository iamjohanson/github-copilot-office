import * as React from "react";
import { Button, Tooltip, makeStyles, Dropdown, Option, tokens } from "@fluentui/react-components";
import { Compose24Regular, History24Regular, Settings24Regular } from "@fluentui/react-icons";

export type ModelType = string;

interface HeaderBarProps {
  onNewChat: () => void;
  onShowHistory: () => void;
  onShowSettings: () => void;
  selectedModel: ModelType;
  onModelChange: (model: ModelType) => void;
  models: { key: string; label: string }[];
  cwd: string | null;
  allowAll: boolean;
}

const useStyles = makeStyles({
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "8px 12px",
    paddingRight: "40px",
    gap: "8px",
    minHeight: "40px",
  },
  leftSection: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
    minWidth: 0,
    flex: 1,
  },
  cwdRow: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
    cursor: "pointer",
    ":hover": {
      color: tokens.colorNeutralForeground1,
    },
  },
  cwdText: {
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    fontFamily: "monospace",
  },
  allowAllBadge: {
    fontSize: "9px",
    fontWeight: 600,
    padding: "0 4px",
    borderRadius: "3px",
    backgroundColor: "#107c10",
    color: "white",
    whiteSpace: "nowrap",
  },
  dropdown: {
    minWidth: "120px",
    opacity: 0.6,
    fontSize: "12px",
    borderBottom: "none",
    ":hover": {
      opacity: 1,
    },
  },
  buttonGroup: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    flexShrink: 0,
  },
  iconButton: {
    minWidth: "28px",
    width: "28px",
    height: "28px",
    padding: "0",
  },
  clearButton: {
    backgroundColor: "#0078d4",
    color: "white",
    borderRadius: "4px",
    padding: "4px",
    width: "28px",
    height: "28px",
    minWidth: "28px",
    border: "none",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    ":hover": {
      backgroundColor: "#106ebe",
    },
  },
});

export const HeaderBar: React.FC<HeaderBarProps> = ({
  onNewChat,
  onShowHistory,
  onShowSettings,
  selectedModel,
  onModelChange,
  models,
  cwd,
  allowAll,
}) => {
  const styles = useStyles();
  const selectedLabel = models.find(m => m.key === selectedModel)?.label || selectedModel;
  const cwdDisplay = cwd ? cwd.split("/").pop() || cwd : null;

  return (
    <div className={styles.header}>
      <div className={styles.leftSection}>
        <Dropdown
          className={styles.dropdown}
          appearance="underline"
          value={selectedLabel}
          selectedOptions={[selectedModel]}
          onOptionSelect={(_, data) => {
            if (data.optionValue && data.optionValue !== selectedModel) {
              onModelChange(data.optionValue as ModelType);
            }
          }}
        >
          {models.map((model) => (
            <Option key={model.key} value={model.key}>
              {model.label}
            </Option>
          ))}
        </Dropdown>
        <Tooltip content={cwd || "No working directory set — click Settings to configure"} relationship="label">
          <div className={styles.cwdRow} onClick={onShowSettings}>
            <span>{cwd ? "📂" : "⚠️"}</span>
            <span className={styles.cwdText}>{cwdDisplay || "No cwd set"}</span>
            {allowAll && <span className={styles.allowAllBadge}>ALLOW ALL</span>}
          </div>
        </Tooltip>
      </div>
      <div className={styles.buttonGroup}>
        <Tooltip content="Settings" relationship="label">
          <Button
            icon={<Settings24Regular />}
            appearance="subtle"
            onClick={onShowSettings}
            aria-label="Settings"
            className={styles.iconButton}
          />
        </Tooltip>
        <Tooltip content="History" relationship="label">
          <Button
            icon={<History24Regular />}
            appearance="subtle"
            onClick={onShowHistory}
            aria-label="History"
            className={styles.iconButton}
          />
        </Tooltip>
        <Tooltip content="New chat" relationship="label">
          <Button
            icon={<Compose24Regular />}
            onClick={onNewChat}
            aria-label="New chat"
            className={styles.clearButton}
          />
        </Tooltip>
      </div>
    </div>
  );
};
