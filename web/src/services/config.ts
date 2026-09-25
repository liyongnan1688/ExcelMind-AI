export interface LlmConfig {
  provider: string;
  baseUrl: string;
  apiKey: string;
  model: string;
  systemPromptAddition: string;
}

export interface PresetProvider {
  id: string;
  name: string;
  baseUrl: string;
  model: string;
  description: string;
}

export const PRESET_PROVIDERS: PresetProvider[] = [
  {
    id: 'deepseek',
    name: 'DeepSeek 官方',
    baseUrl: 'https://api.deepseek.com/v1',
    model: 'deepseek-chat',
    description: '性价比高，代码生成能力强，推荐使用',
  },
  {
    id: 'aliyun-qwen',
    name: '阿里通义千问 (Qwen)',
    baseUrl: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
    model: 'qwen-plus',
    description: '阿里云百炼兼容模式，响应速度快',
  },
  {
    id: 'zhipu-glm',
    name: '智谱 AI (GLM-4)',
    baseUrl: 'https://open.bigmodel.cn/api/paas/v4',
    model: 'glm-4-flash',
    description: 'GLM 官方接口，支持高并发',
  },
  {
    id: 'ollama',
    name: '本地 Ollama (私有化)',
    baseUrl: 'http://localhost:11434/v1',
    model: 'qwen2.5-coder:7b',
    description: '本地离线模型，保护企业核心数据',
  },
  {
    id: 'custom',
    name: '自定义 (OpenAI 兼容)',
    baseUrl: 'https://api.openai.com/v1',
    model: 'gpt-4o',
    description: '适配任何兼容 /v1/chat/completions 的接口',
  },
];

const CONFIG_STORAGE_KEY = 'lee_excel_llm_config';

export function loadLlmConfig(): LlmConfig {
  try {
    const raw = localStorage.getItem(CONFIG_STORAGE_KEY);
    if (raw) {
      return JSON.parse(raw);
    }
  } catch (e) {
    console.error('Failed to load llm config:', e);
  }
  return {
    provider: 'deepseek',
    baseUrl: 'https://api.deepseek.com/v1',
    apiKey: '',
    model: 'deepseek-chat',
    systemPromptAddition: '',
  };
}

export function saveLlmConfig(config: LlmConfig): void {
  try {
    localStorage.setItem(CONFIG_STORAGE_KEY, JSON.stringify(config));
  } catch (e) {
    console.error('Failed to save llm config:', e);
  }
}
