import React, { useRef, useState } from 'react';
import { UploadCloud, Image as ImageIcon, Clipboard, Sparkles, ArrowRight } from 'lucide-react';
import { generateSampleImages, SampleImageItem } from '../data/sampleImages';

interface UploadAreaProps {
  onImageSelected: (file: File | string, filename?: string) => void;
  onError: (msg: string) => void;
}

export const UploadArea: React.FC<UploadAreaProps> = ({ onImageSelected, onError }) => {
  const [isDragging, setIsDragging] = useState(false);
  const [samples] = useState<SampleImageItem[]>(() => generateSampleImages());
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const validateAndProcess = (file: File) => {
    const validTypes = ['image/jpeg', 'image/png', 'image/webp', 'image/jpg'];
    if (!validTypes.includes(file.type)) {
      onError('Unsupported format. Please upload JPG, PNG, or WEBP images.');
      return;
    }
    if (file.size > 25 * 1024 * 1024) {
      onError('File is too large. Maximum supported size is 25MB.');
      return;
    }
    onImageSelected(file, file.name);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      validateAndProcess(e.dataTransfer.files[0]);
    }
  };

  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      validateAndProcess(e.target.files[0]);
    }
  };

  const handlePasteClipboard = async () => {
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        for (const type of item.types) {
          if (type.startsWith('image/')) {
            const blob = await item.getType(type);
            const file = new File([blob], 'clipboard-image.png', { type });
            validateAndProcess(file);
            return;
          }
        }
      }
      onError('No image found in clipboard. Please copy an image first.');
    } catch {
      onError('Unable to read clipboard. You can press Ctrl+V to paste directly.');
    }
  };

  return (
    <div className="upload-container">
      <div 
        className={`upload-card ${isDragging ? 'dragging' : ''}`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        <input 
          type="file" 
          ref={fileInputRef} 
          onChange={handleFileInput} 
          accept="image/png, image/jpeg, image/webp" 
          style={{ display: 'none' }}
        />

        <div className="upload-icon-wrapper">
          <UploadCloud className="upload-hero-icon" size={54} />
        </div>

        <h1 className="upload-title">Drop your image here</h1>
        <p className="upload-subtitle">
          Or browse from your computer to remove background, erase objects, and customize.
        </p>

        <div className="upload-actions">
          <button 
            className="btn btn-primary"
            onClick={() => fileInputRef.current?.click()}
          >
            <ImageIcon size={18} />
            <span>Select Image</span>
          </button>

          <button 
            className="btn btn-secondary"
            onClick={handlePasteClipboard}
          >
            <Clipboard size={18} />
            <span>Paste from Clipboard</span>
          </button>
        </div>

        <div className="upload-badges">
          <span className="badge">JPG</span>
          <span className="badge">PNG</span>
          <span className="badge">WEBP</span>
          <span className="badge-hint">Up to 25MB • 100% Private & In-Browser</span>
        </div>
      </div>

      {/* Preset demo images */}
      <div className="sample-section">
        <div className="sample-header">
          <Sparkles size={16} className="text-accent" />
          <span>Or try one of these test samples:</span>
        </div>
        <div className="sample-grid">
          {samples.map((sample) => (
            <div 
              key={sample.id}
              className="sample-card"
              onClick={() => onImageSelected(sample.dataUrl, `${sample.id}.png`)}
            >
              <div className="sample-thumb-wrapper">
                <img src={sample.dataUrl} alt={sample.name} className="sample-thumb" />
              </div>
              <div className="sample-info">
                <div className="sample-title">{sample.name}</div>
                <div className="sample-desc">{sample.description}</div>
              </div>
              <div className="sample-action">
                <span>Try this</span>
                <ArrowRight size={14} />
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
};
