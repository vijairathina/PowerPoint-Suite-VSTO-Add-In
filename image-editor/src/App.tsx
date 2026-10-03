import React, { useState, useEffect, useRef, useCallback } from 'react';
import { 
  ToolType, 
  BackgroundConfig, 
  AdjustmentConfig, 
  CropRect, 
  BrushSettings, 
  ImageMetadata, 
  HistoryStep 
} from './types/editor';
import { 
  createCanvas, 
  cropCanvas, 
  loadImage, 
  transformCanvas 
} from './services/imageProcessing';
import { 
  removeBackgroundSmart, 
  applyAlphaMask 
} from './services/backgroundRemoval';
import { inpaintObjectRemoval } from './services/inpainting';
import { 
  createCompositeCanvas, 
  downloadCanvas 
} from './services/exportService';

import { Header } from './components/Header';
import { Toolbar } from './components/Toolbar';
import { Canvas } from './components/Canvas';
import { PropertyPanel } from './components/PropertyPanel';
import { UploadArea } from './components/UploadArea';
import { ExportModal } from './components/ExportModal';
import { Toast, ToastMessage } from './components/Toast';

const defaultAdjustments: AdjustmentConfig = {
  brightness: 0,
  contrast: 0,
  saturation: 0,
  exposure: 0,
  blur: 0,
  sharpen: 0,
  opacity: 100,
};

const defaultBackground: BackgroundConfig = {
  type: 'transparent',
  color: '#3b82f6',
  gradient: {
    type: 'linear',
    angle: 135,
    colors: ['#ff7e5f', '#feb47b'],
  },
  imageScale: 1.0,
  imageX: 0,
  imageY: 0,
  imageBlur: 0,
};

const defaultBrush: BrushSettings = {
  size: 35,
  hardness: 80,
  opacity: 100,
  feather: 2,
};

export const App: React.FC = () => {
  // App view state
  const [hasImage, setHasImage] = useState(false);
  const [metadata, setMetadata] = useState<ImageMetadata>({
    name: 'untitled.png',
    originalWidth: 0,
    originalHeight: 0,
    currentWidth: 0,
    currentHeight: 0,
    sizeBytes: 0,
    format: 'png',
  });

  // Main Canvases
  const [originalCanvas, setOriginalCanvas] = useState<HTMLCanvasElement | null>(null);
  const [baseCanvas, setBaseCanvas] = useState<HTMLCanvasElement | null>(null);
  const [maskCanvas, setMaskCanvas] = useState<HTMLCanvasElement | null>(null);
  const [inpaintMaskCanvas, setInpaintMaskCanvas] = useState<HTMLCanvasElement | null>(null);
  const [inpaintHasMark, setInpaintHasMark] = useState(false);

  // Settings
  const [activeTool, setActiveTool] = useState<ToolType>('select');
  const [backgroundConfig, setBackgroundConfig] = useState<BackgroundConfig>(defaultBackground);
  const [bgImageElement, setBgImageElement] = useState<HTMLImageElement | null>(null);
  const [adjustments, setAdjustments] = useState<AdjustmentConfig>(defaultAdjustments);
  const [brushSettings, setBrushSettings] = useState<BrushSettings>(defaultBrush);

  // Background removal options
  const [bgTolerance, setBgTolerance] = useState(38);
  const [bgFeather, setBgFeather] = useState(3);
  const [isRemovingBg, setIsRemovingBg] = useState(false);

  // Inpainting state
  const [isInpainting, setIsInpainting] = useState(false);

  // Crop state
  const [cropRect, setCropRect] = useState<CropRect | null>(null);
  const [cropAspect, setCropAspect] = useState<string>('free');

  // Canvas Viewport (Zoom & Pan)
  const [zoom, setZoom] = useState(1.0);
  const [panX, setPanX] = useState(0);
  const [panY, setPanY] = useState(0);
  const [isCompareMode, setIsCompareMode] = useState(false);

  // Modals & Feedback
  const [isExportOpen, setIsExportOpen] = useState(false);
  const [toasts, setToasts] = useState<ToastMessage[]>([]);

  // History system
  const [history, setHistory] = useState<HistoryStep[]>([]);
  const [historyIndex, setHistoryIndex] = useState(-1);

  const addToast = (type: 'success' | 'error' | 'info', message: string) => {
    const id = Math.random().toString(36).substring(2, 9);
    setToasts((prev) => [...prev, { id, type, message }]);
    setTimeout(() => {
      setToasts((prev) => prev.filter((t) => t.id !== id));
    }, 4500);
  };

  const removeToast = (id: string) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  };

  // Push history step
  const pushHistory = useCallback((description: string, customBase?: HTMLCanvasElement | null, customMask?: HTMLCanvasElement | null) => {
    const activeBase = customBase || baseCanvas;
    const activeMask = customMask !== undefined ? customMask : maskCanvas;
    if (!activeBase) return;

    const step: HistoryStep = {
      description,
      timestamp: Date.now(),
      baseImageDataUrl: activeBase.toDataURL(),
      maskDataUrl: activeMask ? activeMask.toDataURL() : null,
      adjustments: { ...adjustments },
      background: { ...backgroundConfig },
      rotation: 0,
      flipH: false,
      flipV: false,
    };

    setHistory((prev) => {
      const trimmed = prev.slice(0, historyIndex + 1);
      if (trimmed.length > 20) trimmed.shift(); // Limit to 20 steps for memory
      return [...trimmed, step];
    });
    setHistoryIndex((prev) => Math.min(19, prev + 1));
  }, [baseCanvas, maskCanvas, adjustments, backgroundConfig, historyIndex]);

  // Undo
  const handleUndo = async () => {
    if (historyIndex <= 0) return;
    const prevIndex = historyIndex - 1;
    const step = history[prevIndex];
    if (!step) return;

    if (step.baseImageDataUrl) {
      const img = await loadImage(step.baseImageDataUrl);
      const bCanvas = createCanvas(img.width, img.height);
      bCanvas.getContext('2d')!.drawImage(img, 0, 0);
      setBaseCanvas(bCanvas);
      setMetadata((prev) => ({ ...prev, currentWidth: img.width, currentHeight: img.height }));
    }

    if (step.maskDataUrl) {
      const mImg = await loadImage(step.maskDataUrl);
      const mCanvas = createCanvas(mImg.width, mImg.height);
      mCanvas.getContext('2d')!.drawImage(mImg, 0, 0);
      setMaskCanvas(mCanvas);
    } else {
      setMaskCanvas(null);
    }

    setAdjustments(step.adjustments);
    setBackgroundConfig(step.background);
    setHistoryIndex(prevIndex);
    addToast('info', `Undo: ${step.description}`);
  };

  // Redo
  const handleRedo = async () => {
    if (historyIndex >= history.length - 1) return;
    const nextIndex = historyIndex + 1;
    const step = history[nextIndex];
    if (!step) return;

    if (step.baseImageDataUrl) {
      const img = await loadImage(step.baseImageDataUrl);
      const bCanvas = createCanvas(img.width, img.height);
      bCanvas.getContext('2d')!.drawImage(img, 0, 0);
      setBaseCanvas(bCanvas);
      setMetadata((prev) => ({ ...prev, currentWidth: img.width, currentHeight: img.height }));
    }

    if (step.maskDataUrl) {
      const mImg = await loadImage(step.maskDataUrl);
      const mCanvas = createCanvas(mImg.width, mImg.height);
      mCanvas.getContext('2d')!.drawImage(mImg, 0, 0);
      setMaskCanvas(mCanvas);
    } else {
      setMaskCanvas(null);
    }

    setAdjustments(step.adjustments);
    setBackgroundConfig(step.background);
    setHistoryIndex(nextIndex);
    addToast('info', `Redo: ${step.description}`);
  };

  // Reset to original image
  const handleReset = () => {
    if (!originalCanvas) return;
    const resetBase = createCanvas(originalCanvas.width, originalCanvas.height);
    resetBase.getContext('2d')!.drawImage(originalCanvas, 0, 0);
    setBaseCanvas(resetBase);
    setMaskCanvas(null);
    setAdjustments(defaultAdjustments);
    setBackgroundConfig(defaultBackground);
    setZoom(1.0);
    setPanX(0);
    setPanY(0);
    setCropRect(null);
    if (inpaintMaskCanvas) {
      const ctx = inpaintMaskCanvas.getContext('2d')!;
      ctx.clearRect(0, 0, inpaintMaskCanvas.width, inpaintMaskCanvas.height);
      setInpaintHasMark(false);
    }
    pushHistory('Reset to Original', resetBase, null);
    addToast('info', 'Reset all modifications to original image.');
  };

  // Load new image from file or URL
  const handleImageLoaded = async (fileOrUrl: File | string, filename?: string) => {
    try {
      let dataUrl: string;
      let name = filename || 'image.png';
      let size = 0;
      let format = 'png';

      if (typeof fileOrUrl === 'string') {
        dataUrl = fileOrUrl;
        format = fileOrUrl.includes('jpeg') || fileOrUrl.includes('jpg') ? 'jpg' : 'png';
        size = Math.round(dataUrl.length * 0.75);
      } else {
        name = fileOrUrl.name;
        size = fileOrUrl.size;
        format = fileOrUrl.type.split('/')[1] || 'png';
        dataUrl = await new Promise((resolve) => {
          const reader = new FileReader();
          reader.onload = (e) => resolve(e.target?.result as string);
          reader.readAsDataURL(fileOrUrl);
        });
      }

      const img = await loadImage(dataUrl);

      // Create original and base canvases
      const orig = createCanvas(img.width, img.height);
      orig.getContext('2d')!.drawImage(img, 0, 0);
      const base = createCanvas(img.width, img.height);
      base.getContext('2d')!.drawImage(img, 0, 0);

      // Inpaint mask canvas
      const inpaintCanvas = createCanvas(img.width, img.height);

      setOriginalCanvas(orig);
      setBaseCanvas(base);
      setMaskCanvas(null);
      setInpaintMaskCanvas(inpaintCanvas);
      setInpaintHasMark(false);

      setMetadata({
        name,
        originalWidth: img.width,
        originalHeight: img.height,
        currentWidth: img.width,
        currentHeight: img.height,
        sizeBytes: size,
        format,
      });

      // Fit to screen initial zoom
      const containerW = window.innerWidth - 380;
      const containerH = window.innerHeight - 120;
      const scaleW = (containerW * 0.78) / img.width;
      const scaleH = (containerH * 0.78) / img.height;
      const fitZoom = Math.min(1.0, Math.min(scaleW, scaleH));

      setZoom(Math.max(0.2, fitZoom));
      setPanX(0);
      setPanY(0);
      setAdjustments(defaultAdjustments);
      setBackgroundConfig(defaultBackground);
      setActiveTool('select');
      setHasImage(true);

      // Initialize history
      const initialStep: HistoryStep = {
        description: 'Uploaded Image',
        timestamp: Date.now(),
        baseImageDataUrl: base.toDataURL(),
        maskDataUrl: null,
        adjustments: { ...defaultAdjustments },
        background: { ...defaultBackground },
        rotation: 0,
        flipH: false,
        flipV: false,
      };
      setHistory([initialStep]);
      setHistoryIndex(0);

      addToast('success', `Loaded "${name}" (${img.width}×${img.height} px)`);
    } catch (err: any) {
      addToast('error', `Failed to open image: ${err.message}`);
    }
  };

  // Automatic Background Removal
  const handleAutoRemoveBg = () => {
    if (!baseCanvas) return;
    setIsRemovingBg(true);

    setTimeout(() => {
      try {
        const mask = removeBackgroundSmart(baseCanvas, {
          tolerance: bgTolerance,
          feather: bgFeather,
        });
        setMaskCanvas(mask);
        pushHistory('Auto Remove Background', baseCanvas, mask);
        addToast('success', 'Background removed! Select Background tab to add colors or gradients.');
        setActiveTool('background');
      } catch (err: any) {
        addToast('error', `Background removal failed: ${err.message}`);
      } finally {
        setIsRemovingBg(false);
      }
    }, 80);
  };

  // Restore Background
  const handleRestoreBg = () => {
    setMaskCanvas(null);
    pushHistory('Restore Background', baseCanvas, null);
    addToast('info', 'Background restored to original.');
  };

  // Execute Mark & Remove (Inpainting)
  const handleExecuteInpaint = async () => {
    if (!baseCanvas || !inpaintMaskCanvas || !inpaintHasMark) return;
    setIsInpainting(true);

    try {
      const inpainted = await inpaintObjectRemoval(baseCanvas, inpaintMaskCanvas);
      setBaseCanvas(inpainted);

      // Clear inpaint mask
      const ctx = inpaintMaskCanvas.getContext('2d')!;
      ctx.clearRect(0, 0, inpaintMaskCanvas.width, inpaintMaskCanvas.height);
      setInpaintHasMark(false);

      pushHistory('Mark & Remove Object', inpainted, maskCanvas);
      addToast('success', 'Object removed seamlessly!');
    } catch (err: any) {
      addToast('error', `Inpainting failed: ${err.message}`);
    } finally {
      setIsInpainting(false);
    }
  };

  const handleClearInpaintMask = () => {
    if (!inpaintMaskCanvas) return;
    const ctx = inpaintMaskCanvas.getContext('2d')!;
    ctx.clearRect(0, 0, inpaintMaskCanvas.width, inpaintMaskCanvas.height);
    setInpaintHasMark(false);
    addToast('info', 'Marked area cleared.');
  };

  // Rotate & Flip
  const handleRotate = (deg: number) => {
    if (!baseCanvas) return;
    const newBase = transformCanvas(baseCanvas, deg, false, false);
    let newMask = maskCanvas ? transformCanvas(maskCanvas, deg, false, false) : null;
    let newInpaint = inpaintMaskCanvas ? transformCanvas(inpaintMaskCanvas, deg, false, false) : null;

    setBaseCanvas(newBase);
    setMaskCanvas(newMask);
    if (newInpaint) setInpaintMaskCanvas(newInpaint);
    setMetadata((prev) => ({ ...prev, currentWidth: newBase.width, currentHeight: newBase.height }));
    pushHistory(`Rotated ${deg > 0 ? 'Right' : 'Left'} 90°`, newBase, newMask);
  };

  const handleFlip = (flipH: boolean, flipV: boolean) => {
    if (!baseCanvas) return;
    const newBase = transformCanvas(baseCanvas, 0, flipH, flipV);
    let newMask = maskCanvas ? transformCanvas(maskCanvas, 0, flipH, flipV) : null;
    let newInpaint = inpaintMaskCanvas ? transformCanvas(inpaintMaskCanvas, 0, flipH, flipV) : null;

    setBaseCanvas(newBase);
    setMaskCanvas(newMask);
    if (newInpaint) setInpaintMaskCanvas(newInpaint);
    pushHistory(`Flipped ${flipH ? 'Horizontally' : 'Vertically'}`, newBase, newMask);
  };

  // Crop tool initial setup
  useEffect(() => {
    if (activeTool === 'crop' && baseCanvas) {
      const marginX = Math.round(baseCanvas.width * 0.1);
      const marginY = Math.round(baseCanvas.height * 0.1);
      setCropRect({
        x: marginX,
        y: marginY,
        width: baseCanvas.width - marginX * 2,
        height: baseCanvas.height - marginY * 2,
      });
    } else {
      setCropRect(null);
    }
  }, [activeTool, baseCanvas]);

  const handleSelectCropAspect = (aspect: string) => {
    if (!baseCanvas) return;
    setCropAspect(aspect);
    const w = baseCanvas.width;
    const h = baseCanvas.height;

    let targetW = Math.round(w * 0.8);
    let targetH = Math.round(h * 0.8);

    if (aspect === '1:1') {
      const size = Math.min(targetW, targetH);
      targetW = size;
      targetH = size;
    } else if (aspect === '4:3') {
      targetH = Math.round((targetW * 3) / 4);
    } else if (aspect === '16:9') {
      targetH = Math.round((targetW * 9) / 16);
    } else if (aspect === 'original') {
      targetW = Math.round(w * 0.85);
      targetH = Math.round(h * 0.85);
    }

    setCropRect({
      x: Math.round((w - targetW) / 2),
      y: Math.round((h - targetH) / 2),
      width: targetW,
      height: targetH,
    });
  };

  const handleApplyCrop = () => {
    if (!baseCanvas || !cropRect) return;
    const newBase = cropCanvas(baseCanvas, cropRect);
    let newMask = maskCanvas ? cropCanvas(maskCanvas, cropRect) : null;
    const newInpaint = createCanvas(cropRect.width, cropRect.height);

    setBaseCanvas(newBase);
    setMaskCanvas(newMask);
    setInpaintMaskCanvas(newInpaint);
    setCropRect(null);
    setActiveTool('select');
    setMetadata((prev) => ({ ...prev, currentWidth: newBase.width, currentHeight: newBase.height }));
    pushHistory('Crop Image', newBase, newMask);
    addToast('success', `Cropped to ${cropRect.width} × ${cropRect.height} px`);
  };

  // Custom background image loader
  useEffect(() => {
    if (backgroundConfig.type === 'image' && backgroundConfig.imageUrl) {
      loadImage(backgroundConfig.imageUrl)
        .then((img) => setBgImageElement(img))
        .catch(() => addToast('error', 'Failed to load background image'));
    } else {
      setBgImageElement(null);
    }
  }, [backgroundConfig.type, backgroundConfig.imageUrl]);

  // Export & Download
  const handleExport = (format: 'png' | 'jpeg' | 'webp', quality: number) => {
    if (!baseCanvas) return;
    const composite = createCompositeCanvas(
      baseCanvas,
      maskCanvas,
      backgroundConfig,
      bgImageElement,
      adjustments
    );
    downloadCanvas(composite, {
      format,
      quality,
      filename: metadata.name,
    });
    addToast('success', `Exported as ${format.toUpperCase()}!`);
  };

  // Keyboard Shortcuts
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Don't trigger if inside input or textarea
      if (['INPUT', 'TEXTAREA'].includes((e.target as HTMLElement)?.tagName)) return;

      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        if (e.shiftKey) handleRedo();
        else handleUndo();
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
        e.preventDefault();
        handleRedo();
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        setIsExportOpen(true);
      } else if (e.key === 'Escape') {
        if (activeTool === 'crop') setActiveTool('select');
        setIsExportOpen(false);
      } else if (e.key.toLowerCase() === 'v') {
        setActiveTool('select');
      } else if (e.key.toLowerCase() === 'h') {
        setActiveTool('move');
      } else if (e.key.toLowerCase() === 'c') {
        setActiveTool('crop');
      } else if (e.key.toLowerCase() === 'b') {
        setActiveTool('remove-bg');
      } else if (e.key.toLowerCase() === 'e') {
        setActiveTool('eraser');
      } else if (e.key.toLowerCase() === 'm') {
        setActiveTool('mark-remove');
      } else if (e.key.toLowerCase() === 'r') {
        setActiveTool('restore');
      } else if (e.key.toLowerCase() === 'g') {
        setActiveTool('background');
      } else if (e.key.toLowerCase() === 'a') {
        setActiveTool('adjustments');
      } else if (e.key.toLowerCase() === 't') {
        setActiveTool('transform');
      } else if (e.key === '0') {
        // Fit zoom
        if (baseCanvas) {
          const containerW = window.innerWidth - 380;
          const containerH = window.innerHeight - 120;
          const scaleW = (containerW * 0.78) / baseCanvas.width;
          const scaleH = (containerH * 0.78) / baseCanvas.height;
          setZoom(Math.min(1.0, Math.min(scaleW, scaleH)));
          setPanX(0);
          setPanY(0);
        }
      } else if (e.key === '1') {
        setZoom(1.0);
        setPanX(0);
        setPanY(0);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [baseCanvas, historyIndex, history, activeTool]);

  // Window drag & drop / paste listener
  useEffect(() => {
    const handleWindowPaste = async (e: ClipboardEvent) => {
      const items = e.clipboardData?.items;
      if (!items) return;
      for (let i = 0; i < items.length; i++) {
        if (items[i].type.startsWith('image/')) {
          const file = items[i].getAsFile();
          if (file) handleImageLoaded(file, 'pasted-image.png');
          break;
        }
      }
    };

    const handleWindowDrop = (e: DragEvent) => {
      e.preventDefault();
      if (e.dataTransfer?.files && e.dataTransfer.files.length > 0) {
        const file = e.dataTransfer.files[0];
        if (file.type.startsWith('image/')) {
          handleImageLoaded(file);
        }
      }
    };

    const handleWindowDragOver = (e: DragEvent) => {
      e.preventDefault();
    };

    window.addEventListener('paste', handleWindowPaste);
    window.addEventListener('drop', handleWindowDrop);
    window.addEventListener('dragover', handleWindowDragOver);

    return () => {
      window.removeEventListener('paste', handleWindowPaste);
      window.removeEventListener('drop', handleWindowDrop);
      window.removeEventListener('dragover', handleWindowDragOver);
    };
  }, []);

  return (
    <div className="app-root">
      <Toast toasts={toasts} onDismiss={removeToast} />

      {!hasImage ? (
        <UploadArea 
          onImageSelected={handleImageLoaded}
          onError={(msg) => addToast('error', msg)}
        />
      ) : (
        <div className="editor-shell">
          <Header
            canUndo={historyIndex > 0}
            canRedo={historyIndex < history.length - 1}
            onUndo={handleUndo}
            onRedo={handleRedo}
            onReset={handleReset}
            zoom={zoom}
            onZoomIn={() => setZoom((z) => Math.min(5.0, z * 1.2))}
            onZoomOut={() => setZoom((z) => Math.max(0.1, z * 0.8))}
            onZoomFit={() => {
              if (baseCanvas) {
                const containerW = window.innerWidth - 380;
                const containerH = window.innerHeight - 120;
                const scaleW = (containerW * 0.78) / baseCanvas.width;
                const scaleH = (containerH * 0.78) / baseCanvas.height;
                setZoom(Math.min(1.0, Math.min(scaleW, scaleH)));
                setPanX(0);
                setPanY(0);
              }
            }}
            onZoom100={() => {
              setZoom(1.0);
              setPanX(0);
              setPanY(0);
            }}
            isCompareMode={isCompareMode}
            onToggleCompare={() => setIsCompareMode((prev) => !prev)}
            onOpenExport={() => setIsExportOpen(true)}
            onNewImage={() => setHasImage(false)}
          />

          <div className="editor-workspace">
            <Toolbar 
              activeTool={activeTool}
              onSelectTool={setActiveTool}
              hasRemovedBg={maskCanvas !== null}
            />

            <Canvas 
              activeTool={activeTool}
              baseCanvas={baseCanvas}
              maskCanvas={maskCanvas}
              inpaintMaskCanvas={inpaintMaskCanvas}
              backgroundConfig={backgroundConfig}
              bgImageElement={bgImageElement}
              adjustments={adjustments}
              brushSettings={brushSettings}
              zoom={zoom}
              panX={panX}
              panY={panY}
              onPanChange={(x, y) => {
                setPanX(x);
                setPanY(y);
              }}
              onZoomChange={setZoom}
              onMaskUpdate={() => {
                // Mask updated via brush
              }}
              onInpaintMaskUpdate={() => {
                setInpaintHasMark(true);
              }}
              cropRect={cropRect}
              onCropChange={setCropRect}
              isCompareMode={isCompareMode}
              originalImageCanvas={originalCanvas}
            />

            <PropertyPanel 
              activeTool={activeTool}
              metadata={metadata}
              brushSettings={brushSettings}
              onUpdateBrush={(b) => setBrushSettings((prev) => ({ ...prev, ...b }))}
              backgroundConfig={backgroundConfig}
              onUpdateBackground={(b) => {
                setBackgroundConfig((prev) => ({ ...prev, ...b }));
              }}
              adjustments={adjustments}
              onUpdateAdjustments={(a) => setAdjustments((prev) => ({ ...prev, ...a }))}
              onResetAdjustments={() => setAdjustments(defaultAdjustments)}
              onAutoRemoveBg={handleAutoRemoveBg}
              onRestoreBg={handleRestoreBg}
              isRemovingBg={isRemovingBg}
              hasMask={maskCanvas !== null}
              isInpainting={isInpainting}
              onExecuteInpaint={handleExecuteInpaint}
              onClearInpaintMask={handleClearInpaintMask}
              inpaintHasMark={inpaintHasMark}
              cropAspect={cropAspect}
              onSelectCropAspect={handleSelectCropAspect}
              onApplyCrop={handleApplyCrop}
              onCancelCrop={() => {
                setCropRect(null);
                setActiveTool('select');
              }}
              onRotate={handleRotate}
              onFlip={handleFlip}
              bgTolerance={bgTolerance}
              onChangeBgTolerance={setBgTolerance}
              bgFeather={bgFeather}
              onChangeBgFeather={setBgFeather}
              zoom={zoom}
            />
          </div>

          <ExportModal 
            isOpen={isExportOpen}
            onClose={() => setIsExportOpen(false)}
            onExport={handleExport}
            width={metadata.currentWidth}
            height={metadata.currentHeight}
            filename={metadata.name}
          />
        </div>
      )}
    </div>
  );
};

export default App;
