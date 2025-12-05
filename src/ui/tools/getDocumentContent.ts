import { Tool } from "../../../copilot-sdk-nodejs/types";

export const getDocumentContent: Tool = {
  name: "get_document_content",
  description: "Get the OOXML content of the Word document (w:document element only).",
  parameters: {
    type: "object",
    properties: {},
  },
  handler: async () => {
    try {
      return await Word.run(async (context) => {
        const ooxml = context.document.body.getOoxml();
        await context.sync();
        
        // Extract just the w:document element
        const match = ooxml.value.match(/<w:document[^>]*>[\s\S]*<\/w:document>/);
        return match ? match[0] : "(empty document)";
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
