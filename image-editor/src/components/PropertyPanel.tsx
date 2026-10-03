import React, { useRef } from 'react';
import { 
  ToolType, 
  BackgroundConfig, 
  AdjustmentConfig, 
  CropRect, 
  BrushSettings, 
  ImageMetadata 
} from '../types/editor';
import { 
  Wand2, 
  RotateCcw, 
  RotateCw, 
  FlipHorizontal, 
  FlipVertical, 
  Sparkle, 
  Eraser, 
  Paintbrush, 
  Upload, 
  Sliders, 
  Crop as CropIcon, 
  Check, 
  X,
  Layers,
  Palette,
  Info,
  Maximize
} from 'lucide-react';
import { formatBytes } from '../services/imageProcessing';

interface PropertyPanelProps {
  activeTool: ToolType;
  metadata: ImageMetadata;
  brushSettings: BrushSettings;
  onUpdateBrush: (newSettings: Partial<BrushSettings>) => void;
  backgroundConfig: BackgroundConfig;
  onUpdateBackground: (newConfig: Partial<BackgroundConfig>) => void;
  adjustments: AdjustmentConfig;
  onUpdateAdjustments: (newAdjustments: Partial<AdjustmentConfig>) => void;
  onResetAdjustments: () => void;
  onAutoRemoveBg: () => void;
  onRestoreBg: () => void;
  isRemovingBg: boolean;
  hasMask: boolean;
  isInpainting: boolean;
  onExecuteInpaint: () => void;
  onClearInpaintMask: () => void;
  inpaintHasMark: boolean;
  cropAspect: string;
  onSelectCropAspect: (aspect: string) => void;
  onApplyCrop: () => void;
  onCancelCrop: () => void;
  onRotate: (deg: number) => void;
  onFlip: (h: boolean, v: boolean) => void;
  bgTolerance: number;
  onChangeBgTolerance: (val: number) => void;
  bgFeather: number;
  onChangeBgFeather: (val: number) => void;
  zoom: number;
}

export const PropertyPanel: React.FC<PropertyPanelProps> = ({
  activeTool,
  metadata,
  brushSettings,
  onUpdateBrush,
  backgroundConfig,
  onUpdateBackground,
  adjustments,
  onUpdateAdjustments,
  onResetAdjustments,
  onAutoRemoveBg,
  onRestoreBg,
  isRemovingBg,
  hasMask,
  isInpainting,
  onExecuteInpaint,
  onClearInpaintMask,
  inpaintHasMark,
  cropAspect,
  onSelectCropAspect,
  onApplyCrop,
  onCancelCrop,
  onRotate,
  onFlip,
  bgTolerance,
  onChangeBgTolerance,
  bgFeather,
  onChangeBgFeather,
  zoom,
}) => {
  const bgImageInputRef = useRef<HTMLInputElement>(null);

  const colorSwatches = [
    '#ffffff', '#000000', '#f8fafc', '#64748b', 
    '#3b82f6', '#06b6d4', '#10b981', '#f59e0b', 
    '#ef4444', '#ec4899', '#8b5cf6', '#6366f1'
  ];

  const gradientPresets = [
    { name: 'Sunset', colors: ['#ff7e5f', '#feb47b'] },
    { name: 'Ocean', colors: ['#2b5876', '#4e4376'] },
    { name: 'Neon Glow', colors: ['#8a2387', '#e94057'] },
    { name: 'Aurora', colors: ['#00c6ff', '#0072ff'] },
    { name: 'Emerald', colors: ['#11998e', '#38ef7d'] },
    { name: 'Midnight', colors: ['#0f172a', '#334155'] },
  ];

  const handleCustomBgImageUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      const reader = new FileReader();
      reader.onload = (event) => {
        if (event.target?.result) {
          onUpdateBackground({
            type: 'image',
            imageUrl: event.target.result as string,
            imageScale: 1.0,
            imageX: 0,
            imageY: 0,
            imageBlur: 0
          });
        }
      };
      reader.readAsDataURL(file);
    }
  };

  return (
    <aside className="editor-properties">
      {/* 1. IMAGE PROPERTIES (When select/move or default) */}
      {(activeTool === 'select' || activeTool === 'move') && (
        <div className="panel-section">
          <div className="panel-title">
            <Info size={16} />
            <span>Image Properties</span>
          </div>

          <div className="prop-list">
            <div className="prop-row">
              <span className="prop-label">Name</span>
              <span className="prop-value truncate" title={metadata.name}>{metadata.name}</span>
            </div>
            <div className="prop-row">
              <span className="prop-label">Dimensions</span>
              <span className="prop-value">{metadata.currentWidth} × {metadata.currentHeight} px</span>
            </div>
            <div className="prop-row">
              <span className="prop-label">Megapixels</span>
              <span className="prop-value">{((metadata.currentWidth * metadata.currentHeight) / 1000000).toFixed(2)} MP</span>
            </div>
            <div className="prop-row">
              <span className="prop-label">Format</span>
              <span className="prop-value uppercase">{metadata.format}</span>
            </div>
            <div className="prop-row">
              <span className="prop-label">File Size</span>
              <span className="prop-value">{formatBytes(metadata.sizeBytes)}</span>
            </div>
            <div className="prop-row">
              <span className="prop-label">Canvas Zoom</span>
              <span className="prop-value">{Math.round(zoom * 100)}%</span>
            </div>
          </div>

          <div className="panel-divider" />

          <div className="panel-subtitle">Quick AI Actions</div>
          <div className="action-stack">
            <button 
              className="btn btn-action-primary"
              onClick={onAutoRemoveBg}
              disabled={isRemovingBg}
            >
              <Wand2 size={16} />
              <span>{isRemovingBg ? 'Processing AI...' : 'Auto Remove Background'}</span>
            </button>
          </div>
        </div>
      )}

      {/* 2. REMOVE BACKGROUND TOOL */}
      {activeTool === 'remove-bg' && (
        <div className="panel-section">
          <div className="panel-title">
            <Wand2 size={16} />
            <span>Remove Background</span>
          </div>

          <p className="panel-desc">
            Instantly separate foreground subjects from the background. Refine edges below or use the Eraser/Restore brushes.
          </p>

          <div className="action-stack">
            <button 
              className="btn btn-primary btn-block"
              onClick={onAutoRemoveBg}
              disabled={isRemovingBg}
            >
              <Wand2 size={17} />
              <span>{isRemovingBg ? 'Detecting Subject...' : 'Remove Background'}</span>
            </button>

            {hasMask && (
              <button 
                className="btn btn-secondary btn-block"
                onClick={onRestoreBg}
              >
                <RotateCcw size={15} />
                <span>Restore Original Background</span>
              </button>
            )}
          </div>

          <div className="panel-divider" />

          <div className="panel-subtitle">Fine-Tune Detection</div>

          <div className="slider-group">
            <div className="slider-header">
              <span>Tolerance</span>
              <span className="slider-val">{bgTolerance}</span>
            </div>
            <input 
              type="range" 
              min={10} 
              max={80} 
              value={bgTolerance} 
              onChange={(e) => onChangeBgTolerance(Number(e.target.value))}
            />
          </div>

          <div className="slider-group">
            <div className="slider-header">
              <span>Edge Feathering</span>
              <span className="slider-val">{bgFeather}px</span>
            </div>
            <input 
              type="range" 
              min={0} 
              max={15} 
              value={bgFeather} 
              onChange={(e) => onChangeBgFeather(Number(e.target.value))}
            />
          </div>

          <div className="help-box">
            <Sparkle size={14} className="text-accent" />
            <span>Tip: Switch to Eraser or Restore brush to manually refine fine hair or tricky edges.</span>
          </div>
        </div>
      )}

      {/* 3. ERASER & RESTORE TOOLS */}
      {(activeTool === 'eraser' || activeTool === 'restore') && (
        <div className="panel-section">
          <div className="panel-title">
            {activeTool === 'eraser' ? <Eraser size={16} /> : <Paintbrush size={16} />}
            <span>{activeTool === 'eraser' ? 'Eraser Tool' : 'Restore Brush'}</span>
          </div>

          <p className="panel-desc">
            {activeTool === 'eraser' 
              ? 'Paint over pixels to erase them from the foreground layer.'
              : 'Paint over transparent areas to restore the original background image.'}
          </p>

          <div className="slider-group">
            <div className="slider-header">
              <span>Brush Size</span>
              <span className="slider-val">{brushSettings.size}px</span>
            </div>
            <input 
              type="range" 
              min={5} 
              max={250} 
              value={brushSettings.size} 
              onChange={(e) => onUpdateBrush({ size: Number(e.target.value) })}
            />
            {/* Visual preview dot */}
            <div className="brush-preview-circle-wrap">
              <div 
                className="brush-preview-circle" 
                style={{ 
                  width: `${Math.min(60, brushSettings.size)}px`, 
                  height: `${Math.min(60, brushSettings.size)}px` 
                }} 
              />
            </div>
          </div>

          <div className="slider-group">
            <div className="slider-header">
              <span>Hardness</span>
              <span className="slider-val">{brushSettings.hardness}%</span>
            </div>
            <input 
              type="range" 
              min={0} 
              max={100} 
              value={brushSettings.hardness} 
              onChange={(e) => onUpdateBrush({ hardness: Number(e.target.value) })}
            />
          </div>

          <div className="slider-group">
            <div className="slider-header">
              <span>Opacity</span>
              <span className="slider-val">{brushSettings.opacity}%</span>
            </div>
            <input 
              type="range" 
              min={5} 
              max={100} 
              value={brushSettings.opacity} 
              onChange={(e) => onUpdateBrush({ opacity: Number(e.target.value) })}
            />
          </div>
        </div>
      )}

      {/* 4. MARK & REMOVE (INPAINTING) */}
      {activeTool === 'mark-remove' && (
        <div className="panel-section">
          <div className="panel-title">
            <Sparkle size={16} />
            <span>Mark & Remove Object</span>
          </div>

          <p className="panel-desc">
            Highlight unwanted people, watermarks, text, or blemishes with the brush, then click <strong>Remove Object</strong>.
          </p>

          <div className="slider-group">
            <div className="slider-header">
              <span>Marker Brush Size</span>
              <span className="slider-val">{brushSettings.size}px</span>
            </div>
            <input 
              type="range" 
              min={5} 
              max={150} 
              value={brushSettings.size} 
              onChange={(e) => onUpdateBrush({ size: Number(e.target.value) })}
            />
          </div>

          <div className="action-stack">
            <button 
              className="btn btn-primary btn-block btn-glow"
              onClick={onExecuteInpaint}
              disabled={isInpainting || !inpaintHasMark}
            >
              <Sparkle size={16} />
              <span>{isInpainting ? 'Inpainting Seamlessly...' : 'Remove Marked Object'}</span>
            </button>

            <button 
              className="btn btn-secondary btn-block"
              onClick={onClearInpaintMask}
              disabled={!inpaintHasMark}
            >
              <X size={15} />
              <span>Clear Selection</span>
            </button>
          </div>

          <div className="help-box">
            <span>The algorithm naturally fills in background textures, preserving lighting and adjacent patterns.</span>
          </div>
        </div>
      )}

      {/* 5. BACKGROUND CUSTOMIZER */}
      {activeTool === 'background' && (
        <div className="panel-section">
          <div className="panel-title">
            <Palette size={16} />
            <span>Background Layer</span>
          </div>

          {/* Type selector tabs */}
          <div className="tab-group">
            <button 
              className={`tab-btn ${backgroundConfig.type === 'transparent' ? 'active' : ''}`}
              onClick={() => onUpdateBackground({ type: 'transparent' })}
            >
              Transparent
            </button>
            <button 
              className={`tab-btn ${backgroundConfig.type === 'color' ? 'active' : ''}`}
              onClick={() => onUpdateBackground({ type: 'color' })}
            >
              Solid
            </button>
            <button 
              className={`tab-btn ${backgroundConfig.type === 'gradient' ? 'active' : ''}`}
              onClick={() => onUpdateBackground({ type: 'gradient' })}
            >
              Gradient
            </button>
            <button 
              className={`tab-btn ${backgroundConfig.type === 'image' ? 'active' : ''}`}
              onClick={() => onUpdateBackground({ type: 'image' })}
            >
              Image
            </button>
          </div>

          {/* Solid Color Controls */}
          {backgroundConfig.type === 'color' && (
            <div className="bg-subcontrols">
              <div className="panel-subtitle">Preset Colors</div>
              <div className="swatch-grid">
                {colorSwatches.map((color) => (
                  <button
                    key={color}
                    className={`swatch-btn ${backgroundConfig.color.toLowerCase() === color.toLowerCase() ? 'selected' : ''}`}
                    style={{ backgroundColor: color }}
                    onClick={() => onUpdateBackground({ color })}
                  />
                ))}
              </div>

              <div className="color-picker-row">
                <span className="prop-label">Custom Color</span>
                <div className="color-input-wrapper">
                  <input 
                    type="color" 
                    value={backgroundConfig.color}
                    onChange={(e) => onUpdateBackground({ color: e.target.value })}
                  />
                  <span className="color-hex">{backgroundConfig.color}</span>
                </div>
              </div>
            </div>
          )}

          {/* Gradient Controls */}
          {backgroundConfig.type === 'gradient' && (
            <div className="bg-subcontrols">
              <div className="panel-subtitle">Gradient Presets</div>
              <div className="gradient-grid">
                {gradientPresets.map((preset) => (
                  <button
                    key={preset.name}
                    className="gradient-card"
                    style={{ background: `linear-gradient(135deg, ${preset.colors[0]}, ${preset.colors[1]})` }}
                    onClick={() => onUpdateBackground({ 
                      gradient: { 
                        ...backgroundConfig.gradient, 
                        colors: preset.colors 
                      } 
                    })}
                  >
                    <span>{preset.name}</span>
                  </button>
                ))}
              </div>

              <div className="panel-divider" />

              <div className="slider-group">
                <div className="slider-header">
                  <span>Gradient Angle</span>
                  <span className="slider-val">{backgroundConfig.gradient.angle}°</span>
                </div>
                <input 
                  type="range" 
                  min={0} 
                  max={360} 
                  value={backgroundConfig.gradient.angle} 
                  onChange={(e) => onUpdateBackground({ 
                    gradient: { ...backgroundConfig.gradient, angle: Number(e.target.value) } 
                  })}
                />
              </div>

              <div className="color-picker-row">
                <span className="prop-label">Color 1</span>
                <input 
                  type="color" 
                  value={backgroundConfig.gradient.colors[0] || '#ffffff'}
                  onChange={(e) => {
                    const colors = [...backgroundConfig.gradient.colors];
                    colors[0] = e.target.value;
                    onUpdateBackground({ gradient: { ...backgroundConfig.gradient, colors } });
                  }}
                />
              </div>

              <div className="color-picker-row">
                <span className="prop-label">Color 2</span>
                <input 
                  type="color" 
                  value={backgroundConfig.gradient.colors[1] || '#000000'}
                  onChange={(e) => {
                    const colors = [...backgroundConfig.gradient.colors];
                    colors[1] = e.target.value;
                    onUpdateBackground({ gradient: { ...backgroundConfig.gradient, colors } });
                  }}
                />
              </div>
            </div>
          )}

          {/* Custom Image Background */}
          {backgroundConfig.type === 'image' && (
            <div className="bg-subcontrols">
              <input 
                type="file" 
                ref={bgImageInputRef} 
                onChange={handleCustomBgImageUpload}
                accept="image/*"
                style={{ display: 'none' }}
              />
              <button 
                className="btn btn-secondary btn-block"
                onClick={() => bgImageInputRef.current?.click()}
              >
                <Upload size={16} />
                <span>Upload Background Image</span>
              </button>

              {backgroundConfig.imageUrl && (
                <>
                  <div className="slider-group">
                    <div className="slider-header">
                      <span>Scale</span>
                      <span className="slider-val">{Math.round(backgroundConfig.imageScale * 100)}%</span>
                    </div>
                    <input 
                      type="range" 
                      min={0.2} 
                      max={3.0} 
                      step={0.05}
                      value={backgroundConfig.imageScale} 
                      onChange={(e) => onUpdateBackground({ imageScale: Number(e.target.value) })}
                    />
                  </div>

                  <div className="slider-group">
                    <div className="slider-header">
                      <span>Background Blur</span>
                      <span className="slider-val">{backgroundConfig.imageBlur}px</span>
                    </div>
                    <input 
                      type="range" 
                      min={0} 
                      max={30} 
                      value={backgroundConfig.imageBlur} 
                      onChange={(e) => onUpdateBackground({ imageBlur: Number(e.target.value) })}
                    />
                  </div>
                </>
              )}
            </div>
          )}
        </div>
      )}

      {/* 6. ADJUSTMENTS & FILTERS */}
      {activeTool === 'adjustments' && (
        <div className="panel-section">
          <div className="panel-title-row">
            <div className="panel-title">
              <Sliders size={16} />
              <span>Adjustments</span>
            </div>
            <button className="btn-text" onClick={onResetAdjustments}>
              Reset
            </button>
          </div>

          <div className="slider-group">
            <div className="slider-header">
              <span>Brightness</span>
              <span className="slider-val">{adjustments.brightness}</span>
            </div>
            <input 
              type="range" 
              min={-100} 
              max={100} 
              value={adjustments.brightness} 
              onChange={(e) => onUpdateAdjustments({ brightness: Number(e.target.value) })}
            />
          </div>

          <div className="slider-group">
            <div className="slider-header">
              <span>Contrast</span>
              <span className="slider-val">{adjustments.contrast}</span>
            </div>
            <input 
              type="range" 
              min={-100} 
              max={100} 
              value={adjustments.contrast} 
              onChange={(e) => onUpdateAdjustments({ contrast: Number(e.target.value) })}
            />
          </div>

          <div className="slider-group">
            <div className="slider-header">
              <span>Saturation</span>
              <span className="slider-val">{adjustments.saturation}</span>
            </div>
            <input 
              type="range" 
              min={-100} 
              max={100} 
              value={adjustments.saturation} 
              onChange={(e) => onUpdateAdjustments({ saturation: Number(e.target.value) })}
            />
          </div>

          <div className="slider-group">
            <div className="slider-header">
              <span>Exposure</span>
              <span className="slider-val">{adjustments.exposure}</span>
            </div>
            <input 
              type="range" 
              min={-100} 
              max={100} 
              value={adjustments.exposure} 
              onChange={(e) => onUpdateAdjustments({ exposure: Number(e.target.value) })}
            />
          </div>

          <div className="slider-group">
            <div className="slider-header">
              <span>Blur</span>
              <span className="slider-val">{adjustments.blur}px</span>
            </div>
            <input 
              type="range" 
              min={0} 
              max={30} 
              value={adjustments.blur} 
              onChange={(e) => onUpdateAdjustments({ blur: Number(e.target.value) })}
            />
          </div>

          <div className="slider-group">
            <div className="slider-header">
              <span>Sharpen</span>
              <span className="slider-val">{adjustments.sharpen}</span>
            </div>
            <input 
              type="range" 
              min={0} 
              max={100} 
              value={adjustments.sharpen} 
              onChange={(e) => onUpdateAdjustments({ sharpen: Number(e.target.value) })}
            />
          </div>

          <div className="slider-group">
            <div className="slider-header">
              <span>Opacity</span>
              <span className="slider-val">{adjustments.opacity}%</span>
            </div>
            <input 
              type="range" 
              min={10} 
              max={100} 
              value={adjustments.opacity} 
              onChange={(e) => onUpdateAdjustments({ opacity: Number(e.target.value) })}
            />
          </div>
        </div>
      )}

      {/* 7. CROP */}
      {activeTool === 'crop' && (
        <div className="panel-section">
          <div className="panel-title">
            <CropIcon size={16} />
            <span>Crop & Resize</span>
          </div>

          <div className="panel-subtitle">Aspect Ratio</div>
          <div className="aspect-ratio-grid">
            {['free', '1:1', '4:3', '16:9', 'original'].map((aspect) => (
              <button
                key={aspect}
                className={`aspect-btn ${cropAspect === aspect ? 'active' : ''}`}
                onClick={() => onSelectCropAspect(aspect)}
              >
                {aspect.toUpperCase()}
              </button>
            ))}
          </div>

          <div className="action-stack">
            <button className="btn btn-primary btn-block" onClick={onApplyCrop}>
              <Check size={16} />
              <span>Apply Crop</span>
            </button>
            <button className="btn btn-secondary btn-block" onClick={onCancelCrop}>
              <X size={16} />
              <span>Cancel</span>
            </button>
          </div>
        </div>
      )}

      {/* 8. TRANSFORM (ROTATE & FLIP) */}
      {activeTool === 'transform' && (
        <div className="panel-section">
          <div className="panel-title">
            <RotateCw size={16} />
            <span>Rotate & Flip</span>
          </div>

          <div className="panel-subtitle">Rotate</div>
          <div className="btn-row">
            <button className="btn btn-secondary flex-1" onClick={() => onRotate(-90)}>
              <RotateCcw size={16} />
              <span>Left 90°</span>
            </button>
            <button className="btn btn-secondary flex-1" onClick={() => onRotate(90)}>
              <RotateCw size={16} />
              <span>Right 90°</span>
            </button>
          </div>

          <div className="panel-subtitle">Flip</div>
          <div className="btn-row">
            <button className="btn btn-secondary flex-1" onClick={() => onFlip(true, false)}>
              <FlipHorizontal size={16} />
              <span>Horizontal</span>
            </button>
            <button className="btn btn-secondary flex-1" onClick={() => onFlip(false, true)}>
              <FlipVertical size={16} />
              <span>Vertical</span>
            </button>
          </div>
        </div>
      )}
    </aside>
  );
};
