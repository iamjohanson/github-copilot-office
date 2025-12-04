/*---------------------------------------------------------------------------------------------
 *  FORK: WebSocket transport for browser environments
 *  This file is an addition for Word add-in support, not part of the original SDK.
 *--------------------------------------------------------------------------------------------*/

import {
    AbstractMessageReader,
    AbstractMessageWriter,
    DataCallback,
    Disposable,
    Message,
    MessageReader,
    MessageWriter,
} from "vscode-jsonrpc";

/**
 * Parses LSP-style messages from a buffer.
 * Messages are framed as: Content-Length: <length>\r\n\r\n<json>
 */
function parseMessages(buffer: string): { messages: Message[]; remainder: string } {
    const messages: Message[] = [];
    let pos = 0;

    while (pos < buffer.length) {
        const headerEnd = buffer.indexOf("\r\n\r\n", pos);
        if (headerEnd === -1) break;

        const header = buffer.slice(pos, headerEnd);
        const match = header.match(/Content-Length:\s*(\d+)/i);
        if (!match) break;

        const contentLength = parseInt(match[1], 10);
        const contentStart = headerEnd + 4;
        const contentEnd = contentStart + contentLength;

        if (contentEnd > buffer.length) break;

        const content = buffer.slice(contentStart, contentEnd);
        try {
            messages.push(JSON.parse(content));
        } catch {
            // Skip malformed JSON
        }
        pos = contentEnd;
    }

    return { messages, remainder: buffer.slice(pos) };
}

export class WebSocketMessageReader extends AbstractMessageReader implements MessageReader {
    private buffer = "";
    private callback: DataCallback | null = null;

    constructor(private socket: WebSocket) {
        super();
        socket.addEventListener("message", async (event) => {
            let text: string;
            if (event.data instanceof Blob) {
                text = await event.data.text();
            } else {
                text = event.data;
            }

            this.buffer += text;
            const { messages, remainder } = parseMessages(this.buffer);
            this.buffer = remainder;

            for (const msg of messages) {
                this.callback?.(msg);
            }
        });

        socket.addEventListener("error", (event) => {
            this.fireError(new Error("WebSocket error"));
        });

        socket.addEventListener("close", () => {
            this.fireClose();
        });
    }

    listen(callback: DataCallback): Disposable {
        this.callback = callback;
        return {
            dispose: () => {
                this.callback = null;
            },
        };
    }
}

export class WebSocketMessageWriter extends AbstractMessageWriter implements MessageWriter {
    private errorCount = 0;

    constructor(private socket: WebSocket) {
        super();
    }

    async write(msg: Message): Promise<void> {
        try {
            const content = JSON.stringify(msg);
            const header = `Content-Length: ${new TextEncoder().encode(content).length}\r\n\r\n`;
            this.socket.send(header + content);
        } catch (error) {
            this.errorCount++;
            this.fireError(error, msg, this.errorCount);
        }
    }

    end(): void {
        // WebSocket close is handled externally
    }
}
