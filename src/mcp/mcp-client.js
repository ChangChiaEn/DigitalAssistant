/**
 * MCP (Model Context Protocol) Client
 * Connects to local MCP servers and exposes their tools as skills.
 *
 * MCP servers run as local processes and communicate via stdio or HTTP.
 * This client manages connections and translates MCP tools into assistant skills.
 */

class MCPClient {
  constructor() {
    this.servers = new Map(); // serverId -> { config, tools, status }
  }

  /**
   * Register an MCP server configuration.
   * Config format:
   * {
   *   id: 'filesystem',
   *   name: 'File System MCP',
   *   command: 'npx',
   *   args: ['-y', '@anthropic/mcp-filesystem'],
   *   env: { ALLOWED_DIRS: 'C:\\Users' }
   * }
   */
  registerServer(config) {
    this.servers.set(config.id, {
      config,
      tools: [],
      status: 'registered',
    });
  }

  /**
   * Connect to an MCP server by starting its process.
   * For MVP, we use HTTP-based MCP servers that expose /tools and /execute.
   */
  async connectServer(serverId) {
    const server = this.servers.get(serverId);
    if (!server) throw new Error(`MCP server not found: ${serverId}`);

    server.status = 'connecting';

    try {
      // Start the MCP server process via Electron main process
      const result = await window.electronAPI.runCommand(
        `${server.config.command} ${(server.config.args || []).join(' ')}`
      );

      // For HTTP-based servers, discover tools
      if (server.config.url) {
        const response = await fetch(`${server.config.url}/tools`);
        const data = await response.json();
        server.tools = data.tools || [];
      }

      server.status = 'connected';
      return server.tools;
    } catch (err) {
      server.status = 'error';
      throw err;
    }
  }

  /**
   * Execute a tool on a specific MCP server.
   */
  async executeTool(serverId, toolName, args = {}) {
    const server = this.servers.get(serverId);
    if (!server) throw new Error(`MCP server not found: ${serverId}`);
    if (!server.config.url) throw new Error('MCP server has no URL endpoint');

    const response = await fetch(`${server.config.url}/execute`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tool: toolName, arguments: args }),
    });

    return response.json();
  }

  /**
   * Get all available tools across all connected servers.
   * Returns them in a format suitable for the AI prompt.
   */
  getAllTools() {
    const tools = [];
    for (const [serverId, server] of this.servers) {
      if (server.status !== 'connected') continue;
      for (const tool of server.tools) {
        tools.push({
          server: serverId,
          name: tool.name,
          description: tool.description,
          parameters: tool.inputSchema || tool.parameters || {},
        });
      }
    }
    return tools;
  }

  /**
   * Get status of all registered servers.
   */
  getStatus() {
    const status = {};
    for (const [id, server] of this.servers) {
      status[id] = {
        name: server.config.name,
        status: server.status,
        toolCount: server.tools.length,
      };
    }
    return status;
  }

  /**
   * Load MCP server configs from a JSON file (like claude_desktop_config.json format).
   */
  async loadConfig(configPath) {
    try {
      const result = await window.electronAPI.runCommand(`type "${configPath}"`);
      if (!result.success) return [];

      const config = JSON.parse(result.message);
      const servers = config.mcpServers || {};

      for (const [id, serverConfig] of Object.entries(servers)) {
        this.registerServer({
          id,
          name: serverConfig.name || id,
          command: serverConfig.command,
          args: serverConfig.args || [],
          env: serverConfig.env || {},
          url: serverConfig.url,
        });
      }

      return Object.keys(servers);
    } catch {
      return [];
    }
  }
}

if (typeof module !== 'undefined') module.exports = MCPClient;
