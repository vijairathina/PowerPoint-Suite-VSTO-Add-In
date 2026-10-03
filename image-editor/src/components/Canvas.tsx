import React, { useRef, useEffect, useState, useCallback } from 'react';
import { 
  ToolType, 
  BackgroundConfig, 
  AdjustmentConfig, 
  CropRect, 
  BrushSettings 
} from '../types/editor';
import { 
  applySharpen, 
  getFilterString, 
  renderBackground 
} from '../services/imageProcessing';
import { applyAlphaMask, applyMaskBrush } from '../services/backgroundRemoval';

interface CanvasProps {
  activeTool: ToolType;
  baseCanvas: HTMLCanvasElement | null;
  maskCanvas: HTMLCanvasElement | null;
  inpaintMaskCanvas: HTMLCanvasElement | null;
  backgroundConfig: BackgroundConfig;
  bgImageElement: HTMLImageElement | null;
  adjustments: AdjustmentConfig;
  brushSettings: BrushSettings;
  zoom: number;
  panX: number;
  panY: number;
  onPanChange: (x: number, y: number) => void;
  onZoomChange: (zoom: number) => void;
  onMaskUpdate: () => void;
  onInpaintMaskUpdate: () => void;
  cropRect: CropRect | null;
  onCropChange: (rect: CropRect) => void;
  isCompareMode: boolean;
  originalImageCanvas: HTMLCanvasElement | null;
}

export const Canvas: React.FC<CanvasProps> = ({
  activeTool,
  baseCanvas,
  maskCanvas,
  inpaintMaskCanvas,
  backgroundConfig,
  bgImageElement,
  adjustments,
  brushSettings,
  zoom,
  panX,
  panY,
  onPanChange,
  onZoomChange,
  onMaskUpdate,
  onInpaintMaskUpdate,
  cropRect,
  onCropChange,
  isCompareMode,
  originalImageCanvas,
}) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const displayCanvasRef = useRef<HTMLCanvasElement>(null);
  const overlayCanvasRef = useRef<HTMLCanvasElement>(null);

  const [isDrawing, setIsDrawing] = useState(false);
  const [isPanning, setIsPanning] = useState(false);
  const [panStart, setPanStart] = useState({ x: 0, y: 0 });
  const [cursorPos, setCursorPos] = useState<{ x: number; y: number } | null>(null);
  const [compareSplit, setCompareSplit] = useState(0.5); // 0 to 1
  const [isDraggingSplit, setIsDraggingSplit] = useState(false);

  // Crop interaction state
  const [cropDragMode, setCropDragMode] = useState<string | null>(null);
  const [cropDragStart, setCropDragStart] = useState({ x: 0, y: 0, rect: cropRect });

  // Get image coordinates from mouse event
  const getImageCoordinates = useCallback((e: React.MouseEvent): { x: number; y: number } | null => {
    if (!baseCanvas || !displayCanvasRef.current) return null;
    const rect = displayCanvasRef.current.getBoundingClientRect();
    const clientX = e.clientX - rect.left;
    const clientY = e.clientY - rect.top;

    const scaleX = baseCanvas.width / rect.width;
    const scaleY = baseCanvas.height / rect.height;

    const x = Math.round(clientX * scaleX);
    const y = Math.round(clientY * scaleY);

    return { x, y };
  }, [baseCanvas]);

  // Render display canvas
  useEffect(() => {
    if (!baseCanvas || !displayCanvasRef.current) return;

    const canvas = displayCanvasRef.current;
    const width = baseCanvas.width;
    const height = baseCanvas.height;

    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d')!;

    // 1. Render Background
    renderBackground(ctx, width, height, backgroundConfig, bgImageElement);

    // 2. Render Foreground with Mask
    const fgCanvas = applyAlphaMask(baseCanvas, maskCanvas);

    ctx.save();
    ctx.filter = getFilterString(adjustments);
    ctx.drawImage(fgCanvas, 0, 0);
    ctx.restore();

    if (adjustments.sharpen > 0) {
      applySharpen(ctx, width, height, adjustments.sharpen);
    }

    // 3. Before/After Split rendering if compare mode is active
    if (isCompareMode && originalImageCanvas) {
      const splitX = Math.round(width * compareSplit);
      ctx.save();
      ctx.beginPath();
      ctx.rect(0, 0, splitX, height);
      ctx.clip();
      ctx.drawImage(originalImageCanvas, 0, 0);
      ctx.restore();

      // Divider line
      ctx.save();
      ctx.strokeStyle = '#ffffff';
      ctx.lineWidth = 3;
      ctx.beginPath();
      ctx.moveTo(splitX, 0);
      ctx.lineTo(splitX, height);
      ctx.stroke();

      // Divider handle
      ctx.fillStyle = '#3b82f6';
      ctx.beginPath();
      ctx.arc(splitX, height / 2, 14, 0, Math.PI * 2);
      ctx.fill();
      ctx.strokeStyle = '#ffffff';
      ctx.lineWidth = 2;
      ctx.stroke();
      ctx.restore();
    }
  }, [baseCanvas, maskCanvas, backgroundConfig, bgImageElement, adjustments, isCompareMode, compareSplit, originalImageCanvas]);

  // Render overlay canvas (for brush strokes, inpainting markers, and crop box)
  useEffect(() => {
    if (!baseCanvas || !overlayCanvasRef.current) return;

    const canvas = overlayCanvasRef.current;
    const width = baseCanvas.width;
    const height = baseCanvas.height;

    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d')!;
    ctx.clearRect(0, 0, width, height);

    // Render inpainting mask overlay (glowing red)
    if (inpaintMaskCanvas && (activeTool === 'mark-remove' || activeTool === 'select')) {
      ctx.save();
      ctx.globalAlpha = 0.55;
      ctx.fillStyle = '#ef4444';
      ctx.drawImage(inpaintMaskCanvas, 0, 0);
      ctx.restore();
    }

    // Render Crop rect
    if (activeTool === 'crop' && cropRect) {
      ctx.save();
      // Darken outside
      ctx.fillStyle = 'rgba(0, 0, 0, 0.55)';
      ctx.fillRect(0, 0, width, cropRect.y);
      ctx.fillRect(0, cropRect.y + cropRect.height, width, height - (cropRect.y + cropRect.height));
      ctx.fillRect(0, cropRect.y, cropRect.x, cropRect.height);
      ctx.fillRect(cropRect.x + cropRect.width, cropRect.y, width - (cropRect.x + cropRect.width), cropRect.height);

      // Crop border
      ctx.strokeStyle = '#ffffff';
      ctx.lineWidth = 2;
      ctx.strokeRect(cropRect.x, cropRect.y, cropRect.width, cropRect.height);

      // Rule of thirds grid lines
      ctx.strokeStyle = 'rgba(255, 255, 255, 0.35)';
      ctx.lineWidth = 1;
      const thirdW = cropRect.width / 3;
      const thirdH = cropRect.height / 3;
      for (let i = 1; i <= 2; i++) {
        ctx.beginPath();
        ctx.moveTo(cropRect.x + thirdW * i, cropRect.y);
        ctx.lineTo(cropRect.x + thirdW * i, cropRect.y + cropRect.height);
        ctx.stroke();

        ctx.beginPath();
        ctx.moveTo(cropRect.x, cropRect.y + thirdH * i);
        ctx.lineTo(cropRect.x + cropRect.width, cropRect.y + thirdH * i);
        ctx.stroke();
      }

      // Corner handles
      const handleSize = 14;
      ctx.fillStyle = '#ffffff';
      const corners = [
        [cropRect.x, cropRect.y],
        [cropRect.x + cropRect.width, cropRect.y],
        [cropRect.x, cropRect.y + cropRect.height],
        [cropRect.x + cropRect.width, cropRect.y + cropRect.height],
      ];
      for (const [cx, cy] of corners) {
        ctx.fillRect(cx - handleSize / 2, cy - handleSize / 2, handleSize, handleSize);
      }

      ctx.restore();
    }
  }, [baseCanvas, inpaintMaskCanvas, activeTool, cropRect]);

  // Handle Wheel Zoom
  const handleWheel = (e: React.WheelEvent) => {
    e.preventDefault();
    const factor = e.deltaY < 0 ? 1.1 : 0.9;
    const newZoom = Math.min(5.0, Math.max(0.1, zoom * factor));
    onZoomChange(newZoom);
  };

  // Mouse Down
  const handleMouseDown = (e: React.MouseEvent) => {
    if (e.button === 1 || e.altKey || activeTool === 'move') {
      // Middle click or Alt or Move tool = Pan
      setIsPanning(true);
      setPanStart({ x: e.clientX - panX, y: e.clientY - panY });
      return;
    }

    if (isCompareMode) {
      setIsDraggingSplit(true);
      return;
    }

    const coords = getImageCoordinates(e);
    if (!coords) return;

    if (activeTool === 'eraser' || activeTool === 'restore') {
      if (!maskCanvas) return;
      setIsDrawing(true);
      applyMaskBrush(
        maskCanvas,
        coords.x,
        coords.y,
        brushSettings.size,
        activeTool === 'eraser' ? 'erase' : 'restore',
        brushSettings.hardness,
        brushSettings.opacity
      );
      onMaskUpdate();
    } else if (activeTool === 'mark-remove') {
      if (!inpaintMaskCanvas) return;
      setIsDrawing(true);
      const ctx = inpaintMaskCanvas.getContext('2d')!;
      ctx.fillStyle = '#ef4444';
      ctx.beginPath();
      ctx.arc(coords.x, coords.y, brushSettings.size / 2, 0, Math.PI * 2);
      ctx.fill();
      onInpaintMaskUpdate();
    }
  };

  // Mouse Move
  const handleMouseMove = (e: React.MouseEvent) => {
    if (isPanning) {
      onPanChange(e.clientX - panStart.x, e.clientY - panStart.y);
      return;
    }

    if (isDraggingSplit && baseCanvas && displayCanvasRef.current) {
      const rect = displayCanvasRef.current.getBoundingClientRect();
      const relX = (e.clientX - rect.left) / rect.width;
      setCompareSplit(Math.max(0.05, Math.min(0.95, relX)));
      return;
    }

    const coords = getImageCoordinates(e);
    if (coords) {
      setCursorPos(coords);

      if (isDrawing) {
        if (activeTool === 'eraser' || activeTool === 'restore') {
          if (maskCanvas) {
            applyMaskBrush(
              maskCanvas,
              coords.x,
              coords.y,
              brushSettings.size,
              activeTool === 'eraser' ? 'erase' : 'restore',
              brushSettings.hardness,
              brushSettings.opacity
            );
            onMaskUpdate();
          }
        } else if (activeTool === 'mark-remove') {
          if (inpaintMaskCanvas) {
            const ctx = inpaintMaskCanvas.getContext('2d')!;
            ctx.fillStyle = '#ef4444';
            ctx.beginPath();
            ctx.arc(coords.x, coords.y, brushSettings.size / 2, 0, Math.PI * 2);
            ctx.fill();
            onInpaintMaskUpdate();
          }
        }
      }
    }
  };

  const handleMouseUp = () => {
    setIsDrawing(false);
    setIsPanning(false);
    setIsDraggingSplit(false);
  };

  const handleMouseLeave = () => {
    setIsDrawing(false);
    setIsPanning(false);
    setIsDraggingSplit(false);
    setCursorPos(null);
  };

  const showBrushCursor = 
    (activeTool === 'eraser' || activeTool === 'restore' || activeTool === 'mark-remove') && 
    cursorPos !== null && 
    !isPanning;

  return (
    <main 
      className={`editor-viewport ${activeTool === 'move' ? 'cursor-grab' : ''} ${isPanning ? 'cursor-grabbing' : ''}`}
      ref={containerRef}
      onWheel={handleWheel}
      onMouseDown={handleMouseDown}
      onMouseMove={handleMouseMove}
      onMouseUp={handleMouseUp}
      onMouseLeave={handleMouseLeave}
    >
      <div 
        className="canvas-stage"
        style={{
          transform: `translate(${panX}px, ${panY}px) scale(${zoom})`,
          transformOrigin: 'center center',
        }}
      >
        {/* Checkerboard Background for transparency */}
        <div className="checkerboard-backdrop">
          <canvas ref={displayCanvasRef} className="render-canvas" />
          <canvas ref={overlayCanvasRef} className="overlay-canvas" />
        </div>

        {/* Live Brush Circle Cursor */}
        {showBrushCursor && cursorPos && displayCanvasRef.current && (
          <div 
            className="brush-cursor-indicator"
            style={{
              left: `${cursorPos.x}px`,
              top: `${cursorPos.y}px`,
              width: `${brushSettings.size}px`,
              height: `${brushSettings.size}px`,
              borderColor: activeTool === 'mark-remove' ? '#ef4444' : '#ffffff',
            }}
          />
        )}
      </div>

      {/* Floating Info badge at bottom-left of canvas */}
      {baseCanvas && (
        <div className="canvas-status-pill">
          <span>{baseCanvas.width} × {baseCanvas.height} px</span>
          <span className="dot">•</span>
          <span>{Math.round(zoom * 100)}%</span>
          {isCompareMode && (
            <>
              <span className="dot">•</span>
              <span className="text-accent font-semibold">Before (Left) / After (Right)</span>
            </>
          )}
        </div>
      )}
    </main>
  );
};
