import React, { useState, useEffect } from 'react';

interface LogEntry {
  id: number;
  time: string;
  message: string;
}

let logEntries: LogEntry[] = [];
let logId = 0;
let listeners: Array<() => void> = [];

export function log(...args: unknown[]) {
  // Expose log for debugging in DateCellEditor
  if (typeof window !== 'undefined') {
    (window as any).sheetableLog = log;
  }
  const message = args.map(a =>
    typeof a === 'object' ? JSON.stringify(a) : String(a)
  ).join(' ');
  const now = new Date();
  const time = `${now.getMinutes().toString().padStart(2, '0')}:${now.getSeconds().toString().padStart(2, '0')}.${now.getMilliseconds().toString().padStart(3, '0')}`;
  logEntries = [...logEntries.slice(-99), { id: ++logId, time, message }];
  for (const l of listeners) l();
}

export const DebugLogger: React.FC = () => {
  const [entries, setEntries] = useState<LogEntry[]>(logEntries);
  const [collapsed, setCollapsed] = useState(true);

  // Expose open/close for automation/testing
  React.useEffect(() => {
    if (typeof window !== 'undefined') {
      (window as any).openDebugLogger = () => setCollapsed(false);
      (window as any).closeDebugLogger = () => setCollapsed(true);
    }
    return () => {
      if (typeof window !== 'undefined') {
        delete (window as any).openDebugLogger;
        delete (window as any).closeDebugLogger;
      }
    };
  }, []);

  useEffect(() => {
    const update = () => setEntries([...logEntries]);
    listeners.push(update);
    return () => {
      listeners = listeners.filter(l => l !== update);
    };
  }, []);

  return (
    <>
      <button
        onClick={() => setCollapsed(!collapsed)}
        style={{
          position: 'fixed',
          bottom: 4,
          right: 4,
          zIndex: 10000,
          background: '#222',
          color: '#666',
          border: '1px solid #333',
          borderRadius: 4,
          fontSize: '10px',
          fontFamily: 'monospace',
          padding: '2px 6px',
          cursor: 'pointer',
          opacity: 0.5,
          pointerEvents: 'auto',
        }}
        title="Debug Log"
      >
        🪲 {entries.length}
      </button>
      {!collapsed && (
        <div style={{
          position: 'fixed',
          bottom: 28,
          right: 4,
          width: 400,
          maxHeight: 200,
          zIndex: 10000,
          background: '#1a1a2e',
          border: '1px solid #444',
          borderRadius: 4,
          fontFamily: 'monospace',
          fontSize: '11px',
          color: '#ccc',
          overflowY: 'auto',
          padding: '4px 8px',
        }}>
          {entries.map(e => (
            <div key={e.id} style={{ whiteSpace: 'pre-wrap', borderBottom: '1px solid #333', padding: '1px 0' }}>
              <span style={{ color: '#888' }}>{e.time}</span> {e.message}
            </div>
          ))}
        </div>
      )}
    </>
  );
};
