import { Tool } from "../../../copilot-sdk-nodejs/types";

export const setDocumentContent: Tool = {
  name: "set_document_content",
  description: "Replace the entire document body with new OOXML content. Pass the w:document element from get_document_content (modified as needed).",
  parameters: {
    type: "object",
    properties: {
      ooxml: {
        type: "string",
        description: "The w:document element with modifications.",
      },
    },
    required: ["ooxml"],
  },
  handler: async ({ arguments: args }) => {
    const { ooxml } = args as { ooxml: string };
    
    if (!ooxml.trimStart().startsWith('<w:document')) {
      return { 
        textResultForLlm: "Invalid input: must start with <w:document>", 
        resultType: "failure", 
        error: "Invalid input: must start with <w:document>",
        toolTelemetry: {} 
      };
    }
    
    try {
      return await Word.run(async (context) => {
        // Get full package to preserve styles
        const fullOoxml = context.document.body.getOoxml();
        await context.sync();
        
        // Replace w:document element in the full package
        const newOoxml = fullOoxml.value.replace(
          /<w:document[^>]*>[\s\S]*<\/w:document>/,
          ooxml
        );
        
        const body = context.document.body;
        body.clear();
        body.insertOoxml(newOoxml, Word.InsertLocation.start);
        await context.sync();
        return "Document content replaced successfully.";
      });
    } catch (e: any) {
      return { textResultForLlm: e.message, resultType: "failure", error: e.message, toolTelemetry: {} };
    }
  },
};
