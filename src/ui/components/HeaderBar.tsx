import * as React from "react";
import { Button, Tooltip, makeStyles, Dropdown, Option } from "@fluentui/react-components";
import { Compose24Regular, History24Regular } from "@fluentui/react-icons";

export type ModelType = string;

interface HeaderBarProps {
  onNewChat: () => void;
  onShowHistory: () => void;
  selectedModel: ModelType;
  onModelChange: (model: ModelType) => void;
  models: { key: string; label: string }[];
}

const useStyles = makeStyles({
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "8px 12px",
    paddingRight: "40px", // Avoid Office taskpane buttons (close/info)
    gap: "8px",
    minHeight: "40px",
  },
  titleSection: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  logo: {
    width: "20px",
    height: "20px",
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
  historyButton: {
    minWidth: "28px",
    width: "28px",
    height: "28px",
    padding: "0",
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
  },
});

export const HeaderBar: React.FC<HeaderBarProps> = ({ onNewChat, onShowHistory, selectedModel, onModelChange, models }) => {
  const styles = useStyles();
  const selectedLabel = models.find(m => m.key === selectedModel)?.label || selectedModel;

  return (
    <div className={styles.header}>
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
      <div className={styles.buttonGroup}>
        <Tooltip content="History" relationship="label">
          <Button
            icon={<History24Regular />}
            appearance="subtle"
            onClick={onShowHistory}
            aria-label="History"
            className={styles.historyButton}
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
