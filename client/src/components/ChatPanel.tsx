import React, { useState, useRef, useEffect } from 'react';
import { ChatMessage } from '../types/spreadsheet';

const MODEL_OPTIONS = [
  { value: 'claude-sonnet-4-5-20250929', label: 'Sonnet 4.5' },
  { value: 'claude-haiku-4-5-20251001', label: 'Haiku 4.5' },
  { value: 'claude-opus-4-6', label: 'Opus 4.6' },
];

interface ChatPanelProps {
  messages: ChatMessage[];
  isLoading: boolean;
  onSend: (message: string) => void;
  model: string;
  onModelChange: (model: string) => void;
}

export default function ChatPanel({ messages, isLoading, onSend, model, onModelChange }: ChatPanelProps) {
  const [input, setInput] = useState('');
  const [collapsed, setCollapsed] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  const handleSend = () => {
    const trimmed = input.trim();
    if (!trimmed || isLoading) return;
    setInput('');
    onSend(trimmed);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  return (
    <div className={`chat-panel${collapsed ? ' chat-panel--collapsed' : ''}`}>
      {/* Collapsed sidebar strip */}
      <div className="chat-toggle-bar" onClick={() => setCollapsed(false)}>
        <div className="chat-toggle-bar-icon">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
          </svg>
        </div>
        <span className="chat-toggle-bar-label">Assistant</span>
      </div>

      {/* Normal panel content */}
      <div className="chat-header">
        <div className="chat-header-top">
          <span className="chat-title">Claude Assistant</span>
          <div className="chat-header-actions">
            <select
              className="model-select"
              value={model}
              onChange={(e) => onModelChange(e.target.value)}
              disabled={isLoading}
            >
              {MODEL_OPTIONS.map(opt => (
                <option key={opt.value} value={opt.value}>{opt.label}</option>
              ))}
            </select>
            <button
              className="chat-collapse-btn"
              onClick={() => setCollapsed(true)}
              title="Minimize chat"
            >
              <svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round">
                <polyline points="10,2 14,2 14,6" />
                <polyline points="6,14 2,14 2,10" />
              </svg>
            </button>
          </div>
        </div>
        <span className="chat-subtitle">Ask questions or give commands</span>
      </div>

      <div className="chat-messages">
        {messages.map((msg) => (
          <div key={msg.id} className={`chat-message ${msg.role}`}>
            <div className="message-bubble">
              {msg.isLoading ? (
                <div className="typing-indicator">
                  <span /><span /><span />
                </div>
              ) : (
                <pre className="message-text">{msg.content}</pre>
              )}
            </div>
          </div>
        ))}
        <div ref={messagesEndRef} />
      </div>

      <div className="chat-input-area">
        <textarea
          className="chat-input"
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={handleKeyDown}
          placeholder='Type a command or question... (e.g. "Create a bar chart from A1:D6")'
          rows={2}
          disabled={isLoading}
        />
        <button
          className="chat-send-btn"
          onClick={handleSend}
          disabled={!input.trim() || isLoading}
        >
          Send
        </button>
      </div>
    </div>
  );
}
