const { WebSocketServer } = require('ws');
const { spawn } = require('child_process');
const path = require('path');

// Resolve the @github/copilot bin entry point
const COPILOT_MODULE = path.resolve(__dirname, '../node_modules/@github/copilot/index.js');

function setupCopilotProxy(httpsServer) {
  const wss = new WebSocketServer({ noServer: true });

  httpsServer.on('upgrade', (request, socket, head) => {
    const url = new URL(request.url, `https://${request.headers.host}`);
    
    if (url.pathname === '/api/copilot') {
      wss.handleUpgrade(request, socket, head, (ws) => {
        wss.emit('connection', ws, request);
      });
    }
    // Let other WebSocket connections (e.g., Vite HMR) pass through
  });

  wss.on('connection', (ws) => {
    const child = spawn(process.execPath, [COPILOT_MODULE, '--server', '--stdio'], {
      stdio: ['pipe', 'pipe', 'pipe'],
    });

    child.on('error', (err) => {
      ws.close(1011, 'Child process error');
    });

    child.on('exit', (code, signal) => {
      ws.close(1000, 'Child process exited');
    });

    // Proxy child stdout -> WebSocket
    child.stdout.on('data', (data) => {
      if (ws.readyState === ws.OPEN) {
        ws.send(data);
      }
    });

    // Log child stderr
    child.stderr.on('data', (data) => {
    });

    // Proxy WebSocket -> child stdin
    ws.on('message', (data) => {
      if (!child.killed) {
        child.stdin.write(data);
      }
    });

    ws.on('close', () => {
      if (!child.killed) {
        child.kill();
      }
    });

    ws.on('error', (err) => {
      if (!child.killed) {
        child.kill();
      }
    });
  });
}

module.exports = { setupCopilotProxy };
