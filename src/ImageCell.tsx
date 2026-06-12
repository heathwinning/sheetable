import React, { useState, useCallback } from 'react';
import type { ICellRendererParams } from 'ag-grid-community';
import * as api from './api';
import { useAlert } from './DialogProvider';

// --- Renderer: show icon in cell; full preview opens on click ---
export const ImageCellRenderer: React.FC<ICellRendererParams> = (props) => {
  const key = props.value;
  if (!key) return null;

  return <span className="image-cell-indicator" title="Image attached" aria-label="Image attached" />;
};

// --- Upload helper: programmatic file upload ---
function pickAndUploadImage(
  bookId: string,
  onComplete: (key: string) => void,
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
      showAlert?.('Please select an image file');
      return;
    }

    try {
      const { key, uploadUrl } = await api.getUploadUrl(bookId, file.name);
      await api.uploadImage(uploadUrl, file);
      onComplete(key);
    } catch (err) {
      console.error('Image upload failed:', err);
      showAlert?.('Failed to upload image');
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
  imageKey: string | null;
  bookId: string;
  onChange: (key: string) => void;
  onRemove: () => void;
  onClose: () => void;
  showAlert: (msg: string) => void;
}

const ImagePreviewDialog: React.FC<ImagePreviewDialogProps> = ({
  imageKey,
  bookId,
  onChange,
  onRemove,
  onClose,
  showAlert,
}) => {
  const [uploading, setUploading] = useState(false);

  const handleChange = useCallback(() => {
    setUploading(true);
    pickAndUploadImage(
      bookId,
      (newKey) => {
        setUploading(false);
        onChange(newKey);
      },
      () => setUploading(false),
      showAlert,
    );
  }, [bookId, onChange, showAlert]);

  const handleRemove = useCallback(() => {
    onRemove();
  }, [onRemove]);

  return (
    <div className="image-dialog-overlay" onClick={onClose}>
      <div className="image-dialog" onClick={(e) => e.stopPropagation()}>
        {imageKey && (
          <img
            className="image-dialog-preview"
            src={api.imageUrl(bookId, imageKey)}
            alt=""
          />
        )}
        {uploading && <div className="image-dialog-uploading">Uploading...</div>}
        <div className="image-dialog-actions">
          <button className="image-dialog-btn" onClick={handleChange} disabled={uploading}>
            {imageKey ? 'Change' : 'Upload'}
          </button>
          {imageKey && (
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
    imageKey: string | null;
    bookId: string;
    onSave: (key: string | null) => void;
  } | null>(null);

  const openDialog = useCallback(
    (imageKey: string | null, bookId: string, _tableName: string, onSave: (key: string | null) => void) => {
      // If no image yet, open file picker directly
      if (!imageKey) {
        pickAndUploadImage(bookId, (newKey) => onSave(newKey), undefined, (msg) => showAlert(msg));
        return;
      }
      setDialogState({ imageKey, bookId, onSave });
    },
    [showAlert],
  );

  const dialogElement = dialogState ? (
    <ImagePreviewDialog
      imageKey={dialogState.imageKey}
      bookId={dialogState.bookId}
      showAlert={(msg) => showAlert(msg)}
      onChange={(newKey) => {
        dialogState.onSave(newKey);
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
