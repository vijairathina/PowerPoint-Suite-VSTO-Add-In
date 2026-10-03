import { AdjustmentConfig, BackgroundConfig } from '../types/editor';
import { applySharpen, createCanvas, getFilterString, renderBackground } from './imageProcessing';
import { applyAlphaMask } from './backgroundRemoval';

export interface ExportOptions {
  format: 'png' | 'jpeg' | 'webp';
  quality: number; // 0.1 to 1.0
  filename: string;
}

/**
 * Creates the final high-resolution composite canvas
 */
export function createCompositeCanvas(
  baseImageCanvas: HTMLCanvasElement,
  maskCanvas: HTMLCanvasElement | null,
  background: BackgroundConfig,
  bgImageElement: HTMLImageElement | null,
  adjustments: AdjustmentConfig
): HTMLCanvasElement {
  const width = baseImageCanvas.width;
  const height = baseImageCanvas.height;

  const exportCanvas = createCanvas(width, height);
  const ctx = exportCanvas.getContext('2d')!;

  // 1. Render Background layer
  renderBackground(ctx, width, height, background, bgImageElement);

  // 2. Prepare foreground with alpha mask
  const fgCanvas = applyAlphaMask(baseImageCanvas, maskCanvas);

  // 3. Draw foreground with adjustments
  ctx.save();
  ctx.filter = getFilterString(adjustments);
  ctx.drawImage(fgCanvas, 0, 0);
  ctx.restore();

  // 4. Sharpen if required
  if (adjustments.sharpen > 0) {
    applySharpen(ctx, width, height, adjustments.sharpen);
  }

  return exportCanvas;
}

/**
 * Trigger file download in browser
 */
export function downloadCanvas(
  canvas: HTMLCanvasElement,
  options: ExportOptions
) {
  const mimeType = options.format === 'png' ? 'image/png' : options.format === 'jpeg' ? 'image/jpeg' : 'image/webp';
  
  // If JPEG, fill transparent background with white if transparent was selected
  let finalCanvas = canvas;
  if (options.format === 'jpeg') {
    finalCanvas = createCanvas(canvas.width, canvas.height);
    const fCtx = finalCanvas.getContext('2d')!;
    fCtx.fillStyle = '#ffffff';
    fCtx.fillRect(0, 0, canvas.width, canvas.height);
    fCtx.drawImage(canvas, 0, 0);
  }

  const dataUrl = finalCanvas.toDataURL(mimeType, options.quality);
  const link = document.createElement('a');
  const baseName = options.filename.replace(/\.[^/.]+$/, "");
  link.download = `${baseName}_edited.${options.format === 'jpeg' ? 'jpg' : options.format}`;
  link.href = dataUrl;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
