import React from 'react';
import { 
  Undo2, 
  Redo2, 
  RotateCcw, 
  ZoomIn, 
  ZoomOut, 
  Maximize2, 
  SplitSquareVertical, 
  Download, 
  Sparkles,
  Layers
} from 'lucide-react';

interface HeaderProps {
  canUndo: boolean;
  canRedo: boolean;
  onUndo: () => void;
  onRedo: () => void;
  onReset: () => void;
  zoom: number;
  onZoomIn: () => void;
  onZoomOut: () => void;
  onZoomFit: () => void;
  onZoom100: () => void;
  isCompareMode: boolean;
  onToggleCompare: () => void;
  onOpenExport: () => void;
  onNewImage: () => void;
}

export const Header: React.FC<HeaderProps> = ({
  canUndo,
  canRedo,
  onUndo,
  onRedo,
  onReset,
  zoom,
  onZoomIn,
  onZoomOut,
  onZoomFit,
  onZoom100,
  isCompareMode,
  onToggleCompare,
  onOpenExport,
  onNewImage,
}) => {
  return (
    <header className="app-header">
      {/* Left: Brand & Home */}
      <div className="header-left">
        <div className="brand" onClick={onNewImage} title="Upload new image / Home">
          <div className="brand-icon">
            <Sparkles size={20} />
          </div>
          <div className="brand-text">
            <span className="brand-title">Image Editor</span>
            <span className="brand-badge">PRO</span>
          </div>
        </div>
      </div>

      {/* Center: History & Reset */}
      <div className="header-center">
        <div className="btn-group">
          <button 
            className="icon-btn" 
            onClick={onUndo} 
            disabled={!canUndo}
            title="Undo (Ctrl+Z)"
          >
            <Undo2 size={18} />
          </button>
          <button 
            className="icon-btn" 
            onClick={onRedo} 
            disabled={!canRedo}
            title="Redo (Ctrl+Y or Ctrl+Shift+Z)"
          >
            <Redo2 size={18} />
          </button>
        </div>

        <div className="header-divider" />

        <button 
          className="btn-header-secondary" 
          onClick={onReset}
          title="Reset image to original state"
        >
          <RotateCcw size={16} />
          <span>Reset</span>
        </button>
      </div>

      {/* Right: Zoom, Before/After, Download */}
      <div className="header-right">
        {/* Zoom Controls */}
        <div className="zoom-controls">
          <button 
            className="icon-btn icon-btn-sm" 
            onClick={onZoomOut} 
            title="Zoom Out"
          >
            <ZoomOut size={16} />
          </button>
          
          <button 
            className="zoom-level-btn"
            onClick={onZoom100}
            title="Click to reset zoom to 100%"
          >
            {Math.round(zoom * 100)}%
          </button>

          <button 
            className="icon-btn icon-btn-sm" 
            onClick={onZoomIn} 
            title="Zoom In"
          >
            <ZoomIn size={16} />
          </button>

          <button 
            className="icon-btn icon-btn-sm" 
            onClick={onZoomFit} 
            title="Fit to Screen"
          >
            <Maximize2 size={16} />
          </button>
        </div>

        <div className="header-divider" />

        {/* Before / After Toggle */}
        <button 
          className={`btn-header-toggle ${isCompareMode ? 'active' : ''}`}
          onClick={onToggleCompare}
          title="Split Before & After comparison"
        >
          <SplitSquareVertical size={16} />
          <span>Before / After</span>
        </button>

        {/* Download Button */}
        <button 
          className="btn btn-primary btn-download"
          onClick={onOpenExport}
        >
          <Download size={17} />
          <span>Download</span>
        </button>
      </div>
    </header>
  );
};
