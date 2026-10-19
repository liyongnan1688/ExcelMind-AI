/**
 * 进程内工具网关与 MCP 兼容层 (In-Process Tool Gateway & MCP Protocol Adapter)
 * 
 * 职责：
 * 1. 遵循标准 MCP 规范（tools/list 与 tools/call 协议形状）对外暴露工具清单与 Function Calling 格式；
 * 2. 基于统一工具注册表 (ToolRegistry) 实现单一真实源，消除手写 if/else 分派；
 * 3. 在调度底层 COM 之前统一执行参数 Schema 强校验，拦截非法请求与格式注入；
 * 4. 支持在前端持久化外部自定义 MCP Server 配置以备远期扩展。
 */

import {
  getAllToolDescriptors,
  findToolDescriptor,
  type ToolDescriptor,
} from './toolRegistry';

export interface McpToolSchema {
  name: string;
  description: string;
  inputSchema: {
    type: 'object';
    properties: Record<string, {
      type: string;
      description?: string;
      enum?: string[];
      default?: any;
    }>;
    required?: string[];
  };
}

export interface McpToolCallRequest {
  serverName?: string;
  toolName: string;
  arguments: Record<string, any>;
}

export interface McpToolCallResponse {
  ok: boolean;
  toolName: string;
  blueprintSummary?: string;
  content: Array<{
    type: 'text' | 'json' | 'image';
    text?: string;
    json?: any;
  }>;
  error?: string;
}

export interface McpServerConfig {
  id: string;
  name: string;
  endpointUrl: string;
  enabled: boolean;
  apiKey?: string;
}

/**
 * 从统一注册表派生内置本地 MCP 工具清单 (保持向下兼容)
 */
export const BUILTIN_LOCAL_MCP_TOOLS: McpToolSchema[] = getAllToolDescriptors().map((d) => ({
  name: d.name,
  description: d.description,
  inputSchema: d.inputSchema,
}));

class ToolGatewayService {
  private customServers: McpServerConfig[] = [];

  constructor() {
    this.loadCustomServers();
  }

  private loadCustomServers() {
    if (typeof localStorage !== 'undefined') {
      try {
        const raw = localStorage.getItem('excelmind_mcp_servers');
        if (raw) {
          this.customServers = JSON.parse(raw);
        }
      } catch (e) {
        console.warn('[ToolGateway] 读取外部 MCP 配置异常:', e);
      }
    }
  }

  public getRegisteredServers(): McpServerConfig[] {
    return [...this.customServers];
  }

  public saveServers(servers: McpServerConfig[]) {
    this.customServers = [...servers];
    if (typeof localStorage !== 'undefined') {
      localStorage.setItem('excelmind_mcp_servers', JSON.stringify(this.customServers));
    }
  }

  /**
   * 获取所有可用工具定义 (内置本地工具 + 活跃外部工具)
   */
  public getAllAvailableTools(): McpToolSchema[] {
    return getAllToolDescriptors().map((d) => ({
      name: d.name,
      description: d.description,
      inputSchema: d.inputSchema,
    }));
  }

  /**
   * 将可用工具清单格式化为静态、Prompt Caching 友好的 System Prompt 说明块
   */
  public formatToolsSystemPrompt(): string {
    const tools = this.getAllAvailableTools();
    let text = `【可调用的标准原生与外部 MCP 工具集 (Tool Schemas)】：\n`;
    tools.forEach((t, idx) => {
      text += `${idx + 1}. \`${t.name}\`: ${t.description}\n`;
      text += `   参数规范: ${JSON.stringify(t.inputSchema.properties)}\n`;
      if (t.inputSchema.required && t.inputSchema.required.length > 0) {
        text += `   必填项: [${t.inputSchema.required.join(', ')}]\n`;
      }
    });
    return text.trim();
  }

  /**
   * 分发并执行工具调用请求 (经由统一注册表与前置参数校验)
   */
  public async executeTool(call: McpToolCallRequest): Promise<McpToolCallResponse> {
    const { toolName, arguments: args } = call;
    const descriptor = findToolDescriptor(toolName);

    if (!descriptor) {
      return {
        ok: false,
        toolName,
        content: [],
        error: `未知的工具方法: "${toolName}"。请检查 ToolRegistry 注册清单。`,
      };
    }

    // 前置参数强类型与格式校验
    if (descriptor.validate) {
      const validation = descriptor.validate(args || {});
      if (!validation.valid) {
        return {
          ok: false,
          toolName,
          content: [],
          error: `[参数预检拦截] 工具 "${toolName}" 参数不合法: ${validation.error}`,
        };
      }
    }

    // 调用单一注册中心绑定的执行处理器
    const execRes = await descriptor.execute(args || {});
    return {
      ok: execRes.ok,
      toolName,
      blueprintSummary: execRes.blueprintSummary,
      content: execRes.content,
      error: execRes.error,
    };
  }
}

export const mcpGateway = new ToolGatewayService();
export const toolGateway = mcpGateway; // 语义别名
