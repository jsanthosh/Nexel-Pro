import { Router, Request, Response } from 'express';
import { ClaudeService } from '../services/claudeService';
import { buildSystemPrompt } from '../prompts/systemPrompt';

const router = Router();
const claudeService = new ClaudeService();

interface MessageParam {
  role: 'user' | 'assistant';
  content: string;
}

function parseClaudeResponse(rawContent: string): object {
  // Try extracting JSON from ```json ... ``` block
  const jsonMatch = rawContent.match(/```json\s*([\s\S]*?)\s*```/);
  if (jsonMatch) {
    try {
      return JSON.parse(jsonMatch[1]);
    } catch {}
  }

  // Try raw JSON parse
  try {
    return JSON.parse(rawContent);
  } catch {}

  // Fallback: treat as plain text answer
  return {
    actions: [{ type: 'QUERY_RESULT', message: rawContent }],
    explanation: rawContent,
  };
}

router.post('/chat', async (req: Request, res: Response) => {
  const { message, context, history = [], model } = req.body;

  if (!message || typeof message !== 'string') {
    return res.status(400).json({ error: 'message is required' });
  }

  try {
    const systemPrompt = buildSystemPrompt(context || { cells: [], selectedRange: null, rowCount: 100, colCount: 26 });

    const messages: MessageParam[] = [
      ...(history as MessageParam[]).slice(-10),
      { role: 'user', content: message },
    ];

    const rawResponse = await claudeService.sendMessage(systemPrompt, messages, model);
    const parsed = parseClaudeResponse(rawResponse);

    res.json(parsed);
  } catch (err) {
    console.error('Chat error:', err);
    res.status(500).json({ error: 'Failed to process request', details: String(err) });
  }
});

export default router;
