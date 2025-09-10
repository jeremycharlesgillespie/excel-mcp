# Excel Finance MCP Bridge 🌉

A HTTP/WebSocket bridge that enables **remote connections** to the Excel Finance MCP server. This allows Claude Desktop to connect to the MCP server running on a different machine over the network.

## 🚀 Quick Start

### 1. Install Dependencies
```bash
cd bridge
npm install
```

### 2. Build the Bridge
```bash
npm run build
```

### 3. Start the Bridge
```bash
npm start
```

The bridge will:
- Start the Excel Finance MCP server automatically
- Create HTTP and WebSocket endpoints
- Provide a web interface at `http://localhost:3001`

## 🔧 Configuration

Environment variables:
- `BRIDGE_PORT` - Bridge server port (default: 3001)
- `MCP_SERVER_PATH` - Path to MCP server (default: ../dist/index.js)
- `CORS_ORIGINS` - Allowed CORS origins (default: *)
- `MAX_CONNECTIONS` - Maximum WebSocket connections (default: 10)

Example with custom configuration:
```bash
BRIDGE_PORT=8080 MCP_SERVER_PATH=/path/to/server.js npm start
```

## 📡 Connection Methods

### WebSocket (Recommended)
Real-time bidirectional communication:
```javascript
const ws = new WebSocket('ws://your-server:3001');
ws.onopen = () => {
    ws.send(JSON.stringify({
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/list'
    }));
};
```

### HTTP API
Single request/response calls:
```bash
curl -X POST http://your-server:3001/mcp/call \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": 1,
    "method": "analytics_executive_dashboard",
    "params": { ... }
  }'
```

## 🔗 Claude Desktop Integration

### Option 1: Using MCP Fetch Server
Add to your Claude Desktop configuration:
```json
{
  "mcpServers": {
    "excel-finance-remote": {
      "command": "npx",
      "args": [
        "@modelcontextprotocol/server-fetch",
        "http://your-server-ip:3001/mcp/call"
      ]
    }
  }
}
```

### Option 2: Using SSH Tunnel (Secure)
For secure connections over untrusted networks:

1. Create SSH tunnel:
```bash
ssh -L 3001:localhost:3001 user@your-server-ip
```

2. Use localhost in Claude Desktop:
```json
{
  "mcpServers": {
    "excel-finance-remote": {
      "command": "npx",
      "args": [
        "@modelcontextprotocol/server-fetch",
        "http://localhost:3001/mcp/call"
      ]
    }
  }
}
```

## 🖥 Web Interface

Access the web interface at `http://localhost:3001` to:
- Monitor server status
- View active connections
- Restart the MCP server
- See usage examples
- Test API endpoints

## 📊 Available Endpoints

- `GET /` - Web interface
- `GET /health` - Health check
- `GET /mcp/status` - MCP server status
- `POST /mcp/restart` - Restart MCP server
- `POST /mcp/call` - MCP API calls
- `WebSocket /` - Real-time MCP communication

## 🛡 Security Considerations

### For Local Network Use:
- Default configuration allows all origins (`*`)
- No authentication required
- Suitable for trusted local networks

### For Internet Use:
- **Use SSH tunneling** instead of direct exposure
- Consider adding authentication middleware
- Restrict CORS origins to specific domains
- Use HTTPS/WSS in production

Example with restricted CORS:
```bash
CORS_ORIGINS="https://yourdomain.com,https://anotherdomain.com" npm start
```

## 🔧 Development

### Start in Development Mode
```bash
npm run dev
```

### Build Only
```bash
npm run build
```

### Quick Bridge Start
```bash
npm run bridge
```

## 📝 Example Usage

### List All Available Tools
```javascript
// WebSocket
ws.send(JSON.stringify({
    jsonrpc: '2.0',
    id: 1,
    method: 'tools/list'
}));

// HTTP
fetch('http://localhost:3001/mcp/call', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/list'
    })
});
```

### Generate Executive Dashboard
```javascript
const dashboardRequest = {
    jsonrpc: '2.0',
    id: 2,
    method: 'tools/call',
    params: {
        name: 'analytics_executive_dashboard',
        arguments: {
            kpis: [
                {
                    name: 'Revenue',
                    current: 1000000,
                    target: 1200000,
                    previous: 950000,
                    unit: 'USD',
                    category: 'financial',
                    threshold: {
                        excellent: 1200000,
                        good: 1100000,
                        warning: 1000000,
                        critical: 900000
                    },
                    trend: 'increasing',
                    importance: 'high'
                }
            ],
            cashFlow: {
                current: 500000,
                projected13Week: 100000,
                projected52Week: 800000,
                burnRate: 50000,
                runwayMonths: 10
            },
            financial: {
                revenue: { current: 1000000, target: 1200000, variance: -16.7 },
                ebitda: { current: 200000, target: 240000, margin: 20 },
                workingCapital: { current: 150000, target: 120000, days: 55 },
                debtToEquity: { current: 0.3, target: 0.25, trend: 'stable' }
            }
        }
    }
};
```

## 🚨 Troubleshooting

### Bridge Won't Start
- Check if port 3001 is already in use
- Verify MCP server path is correct
- Ensure parent MCP server builds successfully

### Connection Issues
- Check firewall settings
- Verify network connectivity
- Ensure CORS configuration allows your domain

### MCP Server Crashes
- Check the web interface for error logs
- Use the restart button in web interface
- Check MCP server dependencies

## 🎯 Benefits

### Remote Access
- Run MCP server on powerful server hardware
- Access from multiple client machines
- Centralized financial data processing

### Scalability
- Handle multiple concurrent connections
- Share resources across users
- Load balance across multiple servers

### Monitoring
- Real-time status monitoring
- Connection tracking
- Health check endpoints

## 📚 Architecture

```
Claude Desktop  →  HTTP/WebSocket  →  Bridge  →  stdio  →  MCP Server
     (Client)           (Network)        (Proxy)       (Process)
```

The bridge acts as a **protocol translator**:
- Receives HTTP/WebSocket requests from clients
- Forwards them as stdio messages to MCP server
- Returns responses back to clients
- Manages MCP server lifecycle

This enables the **stdio-based MCP server** to work over **network connections**! 🎉