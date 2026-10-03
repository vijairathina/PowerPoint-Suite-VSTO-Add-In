import React, { useState } from 'react';
import { X, Download, FileImage, ShieldCheck, Sparkles } from 'lucide-react';
import { formatBytes } from '../services/imageProcessing';

interface ExportModalProps {
  isOpen: boolean;
  onClose: () => void;
  onExport: (format: 'png' | 'jpeg' | 'webp', quality: number) => void;
  width: number;
  height: number;
  filename: string;
}

export const ExportModal: React.FC<ExportModalProps> = ({
  isOpen,
  onClose,
  onExport,
  width,
  height,
  filename,
}) => {
  const [format, setFormat] = useState<'png' | 'jpeg' | 'webp'>('png');
  const [quality, setQuality] = useState(0.92);

  if (!isOpen) return null;

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-content" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <div className="modal-title-row">
            <div className="modal-icon">
              <Download size={20} />
            </div>
            <div>
              <h2 className="modal-title">Export & Download</h2>
              <p className="modal-subtitle">Choose format and quality settings</p>
            </div>
          </div>
          <button className="icon-btn" onClick={onClose}>
            <X size={18} />
          </button>
        </div>

        <div className="modal-body">
          {/* Format selection */}
          <div className="field-group">
            <label className="field-label">File Format</label>
            <div className="format-selector-grid">
              <div 
                className={`format-card ${format === 'png' ? 'selected' : ''}`}
                onClick={() => setFormat('png')}
              >
                <div className="format-title">PNG</div>
                <div className="format-desc">Best for transparency & graphics</div>
              </div>

              <div 
                className={`format-card ${format === 'jpeg' ? 'selected' : ''}`}
                onClick={() => setFormat('jpeg')}
              >
                <div className="format-title">JPG / JPEG</div>
                <div className="format-desc">Smaller file size for photos</div>
              </div>

              <div 
                className={`format-card ${format === 'webp' ? 'selected' : ''}`}
                onClick={() => setFormat('webp')}
              >
                <div className="format-title">WEBP</div>
                <div className="format-desc">Modern compressed web format</div>
              </div>
            </div>
          </div>

          {/* Quality slider (for JPG and WEBP) */}
          {(format === 'jpeg' || format === 'webp') && (
            <div className="slider-group">
              <div className="slider-header">
                <span>Compression Quality</span>
                <span className="slider-val">{Math.round(quality * 100)}%</span>
              </div>
              <input 
                type="range" 
                min={0.2} 
                max={1.0} 
                step={0.05}
                value={quality} 
                onChange={(e) => setQuality(Number(e.target.value))}
              />
            </div>
          )}

          {/* Export summary */}
          <div className="export-summary-box">
            <div className="summary-row">
              <span>Target Dimensions</span>
              <span className="summary-val">{width} × {height} px</span>
            </div>
            <div className="summary-row">
              <span>Aspect Ratio</span>
              <span className="summary-val">{(width / height).toFixed(2)}:1</span>
            </div>
            <div className="summary-row">
              <span>Transparency</span>
              <span className="summary-val">
                {format === 'png' ? 'Preserved (Alpha channel)' : 'Filled with white background'}
              </span>
            </div>
          </div>
        </div>

        <div className="modal-footer">
          <button className="btn btn-secondary" onClick={onClose}>
            Cancel
          </button>
          <button 
            className="btn btn-primary"
            onClick={() => {
              onExport(format, quality);
              onClose();
            }}
          >
            <Download size={16} />
            <span>Download Image</span>
          </button>
        </div>
      </div>
    </div>
  );
};
