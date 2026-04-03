import React, { useState, useCallback } from 'react';
import type { ICellRendererParams } from 'ag-grid-community';
import * as drive from './drive';
import { useAlert } from './DialogProvider';

// --- Renderer: show thumbnail in cell ---
export const ImageCellRenderer: React.FC<ICellRendererParams> = (props) => {
  const fileId = props.value;
  if (!fileId) return null;

  const thumbUrl = drive.getThumbnailUrl(fileId, 100);

  return (
    <img
      src={thumbUrl}
      alt=""
      style={{
        maxHeight: 24,
        maxWidth: 80,
        objectFit: 'contain',
        verticalAlign: 'middle',
      }}
      referrerPolicy="no-referrer"
    />
  );
};

// --- Upload helper: programmatic file upload ---
function pickAndUploadImage(
  folderId: string,
  tableName: string,
  onComplete: (fileId: string) => void,
  onError?: (err: unknown) => void,
  showAlert?: (msg: string) => void,
): void {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'image/*';
  input.style.display = 'none';
  document.body.appendChild(input);

  input.addEventListener('change', async () => {
    const file = input.files?.[0];
    document.body.removeChild(input);
    if (!file) return;

    if (!file.type.startsWith('image/')) {
      (showAlert ?? window.alert)('Please select an image file');
      return;
    }

    try {
      // Create/find per-table images subfolder
      const imagesFolderId = await drive.findOrCreateSubfolder(tableName, folderId);
      const fileId = await drive.uploadBinaryFile(file, imagesFolderId);
      onComplete(fileId);
    } catch (err) {
      console.error('Image upload failed:', err);
      (showAlert ?? window.alert)('Failed to upload image');
      onError?.(err);
    }
  });

  input.addEventListener('cancel', () => {
    document.body.removeChild(input);
  });

  input.click();
}

// --- Preview dialog for image cells ---
interface ImagePreviewDialogProps {
  fileId: string | null;
  folderId: string;
  tableName: string;
  onChange: (fileId: string) => void;
  onRemove: () => void;
  onClose: () => void;
  showAlert: (msg: string) => void;
}

const ImagePreviewDialog: React.FC<ImagePreviewDialogProps> = ({
  fileId,
  folderId,
  tableName,
  onChange,
  onRemove,
  onClose,
  showAlert,
}) => {
  const [uploading, setUploading] = useState(false);

  const handleChange = useCallback(() => {
    setUploading(true);
    pickAndUploadImage(
      folderId,
      tableName,
      (newFileId) => {
        setUploading(false);
        onChange(newFileId);
      },
      () => setUploading(false),
      showAlert,
    );
  }, [folderId, tableName, onChange, showAlert]);

  const handleRemove = useCallback(() => {
    onRemove();
  }, [onRemove]);

  return (
    <div className="image-dialog-overlay" onClick={onClose}>
      <div className="image-dialog" onClick={(e) => e.stopPropagation()}>
        {fileId && (
          <img
            className="image-dialog-preview"
            src={drive.getThumbnailUrl(fileId, 800)}
            alt=""
            referrerPolicy="no-referrer"
          />
        )}
        {uploading && <div className="image-dialog-uploading">Uploading...</div>}
        <div className="image-dialog-actions">
          <button className="image-dialog-btn" onClick={handleChange} disabled={uploading}>
            {fileId ? 'Change' : 'Upload'}
          </button>
          {fileId && (
            <button className="image-dialog-btn image-dialog-btn-remove" onClick={handleRemove}>
              Remove
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

// --- Hook for managing the image dialog from SpreadsheetGrid ---
export function useImageDialog() {
  const showAlert = useAlert();
  const [dialogState, setDialogState] = useState<{
    fileId: string | null;
    folderId: string;
    tableName: string;
    onSave: (fileId: string | null) => void;
  } | null>(null);

  const openDialog = useCallback(
    (fileId: string | null, folderId: string, tableName: string, onSave: (fileId: string | null) => void) => {
      // If no image yet, open file picker directly
      if (!fileId) {
        pickAndUploadImage(folderId, tableName, (newFileId) => onSave(newFileId), undefined, (msg) => showAlert(msg));
        return;
      }
      setDialogState({ fileId, folderId, tableName, onSave });
    },
    [showAlert],
  );

  const dialogElement = dialogState ? (
    <ImagePreviewDialog
      fileId={dialogState.fileId}
      folderId={dialogState.folderId}
      tableName={dialogState.tableName}
      showAlert={(msg) => showAlert(msg)}
      onChange={(newFileId) => {
        dialogState.onSave(newFileId);
        setDialogState(null);
      }}
      onRemove={() => {
        dialogState.onSave(null);
        setDialogState(null);
      }}
      onClose={() => setDialogState(null)}
    />
  ) : null;

  return { openDialog, dialogElement };
}
