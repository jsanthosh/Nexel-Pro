import React, { useState, useRef, useEffect } from 'react';

interface DocumentTitleProps {
  title: string;
  onRename: (newTitle: string) => void;
}

export default function DocumentTitle({ title, onRename }: DocumentTitleProps) {
  const [isEditing, setIsEditing] = useState(false);
  const [editValue, setEditValue] = useState(title);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => { setEditValue(title); }, [title]);

  useEffect(() => {
    if (isEditing && inputRef.current) {
      inputRef.current.focus();
      inputRef.current.select();
    }
  }, [isEditing]);

  const commit = () => {
    const trimmed = editValue.trim();
    if (trimmed && trimmed !== title) {
      onRename(trimmed);
    } else {
      setEditValue(title);
    }
    setIsEditing(false);
  };

  if (isEditing) {
    return (
      <input
        ref={inputRef}
        className="doc-title-input"
        value={editValue}
        onChange={(e) => setEditValue(e.target.value)}
        onBlur={commit}
        onKeyDown={(e) => {
          if (e.key === 'Enter') commit();
          if (e.key === 'Escape') { setEditValue(title); setIsEditing(false); }
        }}
      />
    );
  }

  return (
    <span
      className="doc-title"
      onClick={() => setIsEditing(true)}
      title="Click to rename"
    >
      {title}
    </span>
  );
}
