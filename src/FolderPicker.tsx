import React, { useState } from 'react';

interface FolderPickerProps {
  folders: Array<{ id: string; name: string }>;
  onSelect: (folderId: string, folderName: string) => void;
  onClose: () => void;
  onRefresh: () => void;
}

export const FolderPicker: React.FC<FolderPickerProps> = ({
  folders,
  onSelect,
  onClose,
  onRefresh,
}) => {
  const [search, setSearch] = useState('');

  const filtered = folders.filter(f =>
    f.name.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div className="dialog-overlay" onClick={onClose}>
      <div className="dialog folder-picker" onClick={e => e.stopPropagation()}>
        <h2>Select Google Drive Folder</h2>
        <p className="dialog-desc">
          Choose a folder to store your data tables. Sheetable will create CSV files and a config file in this folder.
        </p>
        <div className="folder-search">
          <input
            type="text"
            value={search}
            onChange={e => setSearch(e.target.value)}
            placeholder="Search folders..."
            autoFocus
          />
          <button onClick={onRefresh} className="btn-icon" title="Refresh">
            ↻
          </button>
        </div>
        <div className="folder-list">
          {filtered.length === 0 ? (
            <div className="empty-state">No folders found</div>
          ) : (
            filtered.map(f => (
              <div
                key={f.id}
                className="folder-item"
                onClick={() => onSelect(f.id, f.name)}
              >
                <span className="folder-icon">📁</span>
                <span className="folder-name">{f.name}</span>
              </div>
            ))
          )}
        </div>
        <div className="dialog-actions">
          <button onClick={onClose} className="btn-secondary">Cancel</button>
        </div>
      </div>
    </div>
  );
};
