import { getDocumentContent } from "./getDocumentContent";
import { setDocumentContent } from "./setDocumentContent";
import { insertOoxml } from "./insertOoxml";
import { replaceText } from "./replaceText";
import { getSelection } from "./getSelection";
import { webFetch } from "./webFetch";

export const wordTools = [
  getDocumentContent,
  setDocumentContent,
  insertOoxml,
  replaceText,
  getSelection,
  webFetch,
];
