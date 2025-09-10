import express from 'express';
import cors from 'cors';
import { WebSocket, WebSocketServer } from 'ws';
import { spawn, ChildProcess } from 'child_process';
import { createServer } from 'http';
import path from 'path';

interface MCPMessage {
  jsonrpc: string;
  id?: string | number;
  method?: string;
  params?: any;
  result?: any;
  error?: any;
}

interface BridgeConfig {
  port: number;
  mcpServerPath: string;
  corsOrigins: string[];
  maxConnections: number;
}

class MCPBridge {
  private app: express.Application;
  private server: any;
  private wss: WebSocketServer;
  private config: BridgeConfig;
  private mcpProcess: ChildProcess | null = null;
  private connections: Set<WebSocket> = new Set();

  constructor(config: BridgeConfig) {
    this.config = config;
    this.app = express();
    this.server = createServer(this.app);
    this.wss = new WebSocketServer({ server: this.server });
    
    this.setupExpress();
    this.setupWebSocket();
    this.startMCPServer();
  }

  private setupExpress(): void {
    // Enable CORS for specified origins
    this.app.use(cors({
      origin: this.config.corsOrigins,
      methods: ['GET', 'POST', 'OPTIONS'],
      allowedHeaders: ['Content-Type', 'Authorization'],
      credentials: true
    }));

    this.app.use(express.json({ limit: '50mb' }));
    this.app.use(express.urlencoded({ extended: true, limit: '50mb' }));

    // Health check endpoint
    this.app.get('/health', (req, res) => {
      res.json({
        status: 'healthy',
        mcpServer: this.mcpProcess ? 'running' : 'stopped',
        connections: this.connections.size,
        uptime: process.uptime(),
        timestamp: new Date().toISOString()
      });
    });

    // MCP server status
    this.app.get('/mcp/status', (req, res) => {
      res.json({
        running: this.mcpProcess !== null && !this.mcpProcess.killed,
        pid: this.mcpProcess?.pid,
        connections: this.connections.size,
        maxConnections: this.config.maxConnections
      });
    });

    // Restart MCP server
    this.app.post('/mcp/restart', async (req, res) => {
      try {
        await this.restartMCPServer();
        res.json({ success: true, message: 'MCP server restarted successfully' });
      } catch (error) {
        res.status(500).json({ 
          success: false, 
          error: error instanceof Error ? error.message : 'Failed to restart MCP server' 
        });
      }
    });

    // HTTP endpoint for MCP calls (alternative to WebSocket)
    this.app.post('/mcp/call', async (req, res) => {
      try {
        const message: MCPMessage = req.body;
        
        if (!this.mcpProcess) {
          throw new Error('MCP server not running');
        }

        // Send message to MCP server and wait for response
        const response = await this.sendToMCPServer(message);
        res.json(response);
      } catch (error) {
        res.status(500).json({
          jsonrpc: '2.0',
          id: req.body.id || null,
          error: {
            code: -32603,
            message: error instanceof Error ? error.message : 'Internal error'
          }
        });
      }
    });

    // Serve static files for a simple web interface
    this.app.get('/', (req, res) => {
      res.send(this.generateWebInterface());
    });
  }

  private setupWebSocket(): void {
    this.wss.on('connection', (ws: WebSocket, req) => {
      console.log(`New WebSocket connection from ${req.socket.remoteAddress}`);
      
      // Check connection limit
      if (this.connections.size >= this.config.maxConnections) {
        ws.close(1008, 'Maximum connections exceeded');
        return;
      }

      this.connections.add(ws);

      // Handle incoming messages
      ws.on('message', async (data: Buffer) => {
        try {
          const message: MCPMessage = JSON.parse(data.toString());
          console.log('Received WebSocket message:', message.method || 'response');
          
          if (!this.mcpProcess) {
            ws.send(JSON.stringify({
              jsonrpc: '2.0',
              id: message.id || null,
              error: { code: -32603, message: 'MCP server not running' }
            }));
            return;
          }

          // Forward message to MCP server
          const response = await this.sendToMCPServer(message);
          ws.send(JSON.stringify(response));
          
        } catch (error) {
          console.error('WebSocket message error:', error);
          ws.send(JSON.stringify({
            jsonrpc: '2.0',
            id: null,
            error: {
              code: -32700,
              message: 'Parse error: ' + (error instanceof Error ? error.message : 'Unknown error')
            }
          }));
        }
      });

      // Handle connection close
      ws.on('close', (code, reason) => {
        console.log(`WebSocket connection closed: ${code} ${reason.toString()}`);
        this.connections.delete(ws);
      });

      // Handle errors
      ws.on('error', (error) => {
        console.error('WebSocket error:', error);
        this.connections.delete(ws);
      });

      // Send welcome message
      ws.send(JSON.stringify({
        jsonrpc: '2.0',
        method: 'bridge/connected',
        params: {
          message: 'Connected to Excel Finance MCP Bridge',
          serverStatus: this.mcpProcess ? 'running' : 'stopped',
          timestamp: new Date().toISOString()
        }
      }));
    });
  }

  private async startMCPServer(): Promise<void> {
    try {
      console.log('Starting MCP server...');
      
      // Spawn the MCP server process
      this.mcpProcess = spawn('node', [this.config.mcpServerPath], {
        stdio: ['pipe', 'pipe', 'pipe'],
        cwd: path.dirname(this.config.mcpServerPath)
      });

      if (!this.mcpProcess.stdin || !this.mcpProcess.stdout || !this.mcpProcess.stderr) {
        throw new Error('Failed to create MCP server stdio streams');
      }

      // Handle MCP server output
      this.mcpProcess.stdout.on('data', (data: Buffer) => {
        const output = data.toString().trim();
        if (output && !output.includes('Excel Finance MCP Server running on stdio')) {
          console.log('MCP Server output:', output);
        }
      });

      this.mcpProcess.stderr.on('data', (data: Buffer) => {
        const error = data.toString().trim();
        if (error && !error.includes('Excel Finance MCP Server running on stdio')) {
          console.error('MCP Server error:', error);
        }
      });

      this.mcpProcess.on('exit', (code, signal) => {
        console.log(`MCP server exited with code ${code}, signal ${signal}`);
        this.mcpProcess = null;
        
        // Notify all connected clients
        const message = JSON.stringify({
          jsonrpc: '2.0',
          method: 'bridge/server_disconnected',
          params: { code, signal, timestamp: new Date().toISOString() }
        });
        
        this.connections.forEach(ws => {
          if (ws.readyState === WebSocket.OPEN) {
            ws.send(message);
          }
        });
      });

      this.mcpProcess.on('error', (error) => {
        console.error('MCP server process error:', error);
        this.mcpProcess = null;
      });

      console.log(`MCP server started with PID: ${this.mcpProcess.pid}`);
    } catch (error) {
      console.error('Failed to start MCP server:', error);
      throw error;
    }
  }

  private async restartMCPServer(): Promise<void> {
    if (this.mcpProcess) {
      this.mcpProcess.kill('SIGTERM');
      
      // Wait for process to exit
      await new Promise<void>((resolve) => {
        if (!this.mcpProcess) {
          resolve();
          return;
        }
        
        this.mcpProcess.on('exit', () => {
          resolve();
        });
        
        // Force kill after 5 seconds
        setTimeout(() => {
          if (this.mcpProcess && !this.mcpProcess.killed) {
            this.mcpProcess.kill('SIGKILL');
          }
          resolve();
        }, 5000);
      });
    }

    await this.startMCPServer();
  }

  private sendToMCPServer(message: MCPMessage): Promise<MCPMessage> {
    return new Promise((resolve, reject) => {
      if (!this.mcpProcess?.stdin || !this.mcpProcess?.stdout) {
        reject(new Error('MCP server not available'));
        return;
      }

      const messageId = message.id || Date.now().toString();
      const messageWithId = { ...message, id: messageId };
      
      // Set up response handler
      const responseHandler = (data: Buffer) => {
        try {
          const response: MCPMessage = JSON.parse(data.toString());
          if (response.id === messageId) {
            this.mcpProcess?.stdout?.removeListener('data', responseHandler);
            resolve(response);
          }
        } catch (error) {
          // Ignore parsing errors for other messages
        }
      };

      // Set up timeout
      const timeout = setTimeout(() => {
        this.mcpProcess?.stdout?.removeListener('data', responseHandler);
        reject(new Error('MCP server response timeout'));
      }, 30000); // 30 second timeout

      this.mcpProcess.stdout.on('data', responseHandler);

      // Send message to MCP server
      try {
        this.mcpProcess.stdin.write(JSON.stringify(messageWithId) + '\n');
      } catch (error) {
        clearTimeout(timeout);
        this.mcpProcess?.stdout?.removeListener('data', responseHandler);
        reject(error);
      }
    });
  }

  private generateWebInterface(): string {
    return `
<!DOCTYPE html>
<html>
<head>
    <title>Excel Finance MCP Bridge</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; background-color: #f5f5f5; }
        .container { max-width: 800px; margin: 0 auto; background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
        .status { padding: 15px; margin: 20px 0; border-radius: 5px; }
        .status.healthy { background-color: #d4edda; border: 1px solid #c3e6cb; color: #155724; }
        .status.warning { background-color: #fff3cd; border: 1px solid #ffeaa7; color: #856404; }
        .endpoint { background: #f8f9fa; padding: 15px; margin: 10px 0; border-left: 4px solid #007bff; }
        .code { background: #2d3748; color: #e2e8f0; padding: 15px; border-radius: 5px; font-family: 'Courier New', monospace; margin: 10px 0; }
        button { background: #007bff; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; margin: 5px; }
        button:hover { background: #0056b3; }
        .connections { display: inline-block; margin-left: 20px; color: #666; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Excel Finance MCP Bridge</h1>
        
        <div id="status" class="status">
            <strong>Status:</strong> Loading...
        </div>
        
        <div>
            <button onclick="checkStatus()">🔄 Refresh Status</button>
            <button onclick="restartServer()">🔂 Restart MCP Server</button>
            <span id="connections" class="connections"></span>
        </div>
        
        <h2>🔗 Connection Endpoints</h2>
        
        <div class="endpoint">
            <strong>WebSocket:</strong><br>
            <code>ws://localhost:${this.config.port}/</code><br>
            <small>For real-time bidirectional communication with the MCP server</small>
        </div>
        
        <div class="endpoint">
            <strong>HTTP API:</strong><br>
            <code>POST http://localhost:${this.config.port}/mcp/call</code><br>
            <small>For single request/response MCP calls</small>
        </div>
        
        <div class="endpoint">
            <strong>Health Check:</strong><br>
            <code>GET http://localhost:${this.config.port}/health</code><br>
            <small>Check bridge and MCP server status</small>
        </div>
        
        <h2>📝 Usage Examples</h2>
        
        <h3>WebSocket Connection (JavaScript)</h3>
        <div class="code">const ws = new WebSocket('ws://localhost:${this.config.port}');
ws.onopen = () => {
    ws.send(JSON.stringify({
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/list'
    }));
};
ws.onmessage = (event) => {
    console.log('Response:', JSON.parse(event.data));
};</div>

        <h3>HTTP API Call (curl)</h3>
        <div class="code">curl -X POST http://localhost:${this.config.port}/mcp/call \\
  -H "Content-Type: application/json" \\
  -d '{
    "jsonrpc": "2.0",
    "id": 1,
    "method": "tools/list"
  }'</div>

        <h3>Claude Desktop Configuration</h3>
        <p>Add this to your Claude Desktop MCP configuration:</p>
        <div class="code">{
  "mcpServers": {
    "excel-finance-remote": {
      "command": "npx",
      "args": [
        "@modelcontextprotocol/server-fetch",
        "http://localhost:${this.config.port}/mcp/call"
      ]
    }
  }
}</div>
        
        <h2>🛠 Available Tools</h2>
        <ul>
            <li><strong>Excel Operations:</strong> Create, read, write workbooks with formulas</li>
            <li><strong>Financial Analysis:</strong> NPV, IRR, DCF valuation models</li>
            <li><strong>Cash Flow Forecasting:</strong> 13-week and 52-week projections</li>
            <li><strong>Risk Analytics:</strong> Monte Carlo simulations, sensitivity analysis</li>
            <li><strong>Executive Dashboards:</strong> KPI tracking and automated insights</li>
            <li><strong>Professional Templates:</strong> Financial ratios, loan analysis, rent rolls</li>
        </ul>
    </div>
    
    <script>
        async function checkStatus() {
            try {
                const response = await fetch('/health');
                const data = await response.json();
                
                const statusDiv = document.getElementById('status');
                const connectionsSpan = document.getElementById('connections');
                
                statusDiv.className = 'status ' + (data.status === 'healthy' ? 'healthy' : 'warning');
                statusDiv.innerHTML = \`
                    <strong>Bridge Status:</strong> \${data.status}<br>
                    <strong>MCP Server:</strong> \${data.mcpServer}<br>
                    <strong>Uptime:</strong> \${Math.floor(data.uptime)} seconds
                \`;
                
                connectionsSpan.textContent = \`Active Connections: \${data.connections}\`;
            } catch (error) {
                document.getElementById('status').innerHTML = '<strong>Error:</strong> ' + error.message;
            }
        }
        
        async function restartServer() {
            try {
                const response = await fetch('/mcp/restart', { method: 'POST' });
                const data = await response.json();
                
                if (data.success) {
                    alert('MCP server restarted successfully');
                    checkStatus();
                } else {
                    alert('Error: ' + data.error);
                }
            } catch (error) {
                alert('Error restarting server: ' + error.message);
            }
        }
        
        // Check status on page load and every 30 seconds
        checkStatus();
        setInterval(checkStatus, 30000);
    </script>
</body>
</html>`;
  }

  public start(): Promise<void> {
    return new Promise((resolve, reject) => {
      this.server.listen(this.config.port, () => {
        console.log(`🌉 Excel Finance MCP Bridge running on port ${this.config.port}`);
        console.log(`📊 Web interface: http://localhost:${this.config.port}`);
        console.log(`🔗 WebSocket endpoint: ws://localhost:${this.config.port}`);
        console.log(`🚀 HTTP API endpoint: http://localhost:${this.config.port}/mcp/call`);
        resolve();
      }).on('error', reject);
    });
  }

  public async stop(): Promise<void> {
    console.log('Stopping MCP Bridge...');
    
    // Close all WebSocket connections
    this.connections.forEach(ws => {
      ws.close(1001, 'Server shutting down');
    });
    
    // Stop MCP server
    if (this.mcpProcess) {
      this.mcpProcess.kill('SIGTERM');
      this.mcpProcess = null;
    }
    
    // Close HTTP server
    return new Promise((resolve) => {
      this.server.close(() => {
        console.log('MCP Bridge stopped');
        resolve();
      });
    });
  }
}

// Configuration
const config: BridgeConfig = {
  port: parseInt(process.env.BRIDGE_PORT || '3001'),
  mcpServerPath: process.env.MCP_SERVER_PATH || path.join('..', 'dist', 'index.js'),
  corsOrigins: process.env.CORS_ORIGINS?.split(',') || ['*'],
  maxConnections: parseInt(process.env.MAX_CONNECTIONS || '10')
};

// Start the bridge
async function main() {
  const bridge = new MCPBridge(config);
  
  // Handle graceful shutdown
  process.on('SIGINT', async () => {
    console.log('\\nReceived SIGINT, shutting down gracefully...');
    await bridge.stop();
    process.exit(0);
  });
  
  process.on('SIGTERM', async () => {
    console.log('\\nReceived SIGTERM, shutting down gracefully...');
    await bridge.stop();
    process.exit(0);
  });
  
  try {
    await bridge.start();
  } catch (error) {
    console.error('Failed to start MCP Bridge:', error);
    process.exit(1);
  }
}

main().catch(console.error);