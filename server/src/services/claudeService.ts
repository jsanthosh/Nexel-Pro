import Anthropic from '@anthropic-ai/sdk';

interface MessageParam {
  role: 'user' | 'assistant';
  content: string;
}

export class ClaudeService {
  private client: Anthropic;

  constructor() {
    this.client = new Anthropic({
      apiKey: process.env.ANTHROPIC_API_KEY,
    });
  }

  private static readonly ALLOWED_MODELS = new Set([
    'claude-sonnet-4-5-20250929',
    'claude-haiku-4-5-20251001',
    'claude-opus-4-6',
  ]);

  async sendMessage(systemPrompt: string, messages: MessageParam[], model?: string): Promise<string> {
    const selectedModel = model && ClaudeService.ALLOWED_MODELS.has(model)
      ? model
      : 'claude-sonnet-4-5-20250929';

    const response = await this.client.messages.create({
      model: selectedModel,
      max_tokens: 4096,
      system: systemPrompt,
      messages,
    });

    const textBlock = response.content.find(b => b.type === 'text');
    if (!textBlock || textBlock.type !== 'text') {
      throw new Error('No text content in Claude response');
    }
    return textBlock.text;
  }
}
