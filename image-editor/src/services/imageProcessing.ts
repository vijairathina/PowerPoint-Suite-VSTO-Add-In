import { AdjustmentConfig, BackgroundConfig, CropRect } from '../types/editor';

/**
 * Format bytes into human readable string
 */
export function formatBytes(bytes: number, decimals = 1): string {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(dm))} ${sizes[i]}`;
}

/**
 * Load HTMLImageElement from a URL or Data URI
 */
export function loadImage(src: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = () => resolve(img);
    img.onerror = (err) => reject(new Error('Failed to load image: ' + err));
    img.src = src;
  });
}

/**
 * Create a canvas with given dimensions
 */
export function createCanvas(width: number, height: number): HTMLCanvasElement {
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(width));
  canvas.height = Math.max(1, Math.round(height));
  return canvas;
}

/**
 * Apply CSS filter string from AdjustmentConfig
 */
export function getFilterString(adj: AdjustmentConfig): string {
  const brightness = 1 + adj.brightness / 100;
  const contrast = 1 + adj.contrast / 100;
  const saturate = 1 + adj.saturation / 100;
  const exposure = Math.pow(2, adj.exposure / 50); // exposure multiplier
  const totalBrightness = brightness * exposure;
  const blur = adj.blur > 0 ? `blur(${adj.blur}px)` : '';
  const opacity = adj.opacity / 100;

  return `brightness(${totalBrightness}) contrast(${contrast}) saturate(${saturate}) opacity(${opacity}) ${blur}`.trim();
}

/**
 * Apply a convolution matrix (e.g. for sharpening)
 */
export function applySharpen(ctx: CanvasRenderingContext2D, width: number, height: number, amount: number) {
  if (amount <= 0) return;
  const strength = (amount / 100) * 1.5;
  const imgData = ctx.getImageData(0, 0, width, height);
  const data = imgData.data;
  const copy = new Uint8ClampedArray(data);

  // 3x3 Laplacian / Unsharp mask kernel
  // [ 0, -s,  0 ]
  // [-s, 1+4s, -s]
  // [ 0, -s,  0 ]
  const centerWeight = 1 + 4 * strength;
  const edgeWeight = -strength;

  for (let y = 1; y < height - 1; y++) {
    for (let x = 1; x < width - 1; x++) {
      const idx = (y * width + x) * 4;
      for (let c = 0; c < 3; c++) {
        const top = copy[((y - 1) * width + x) * 4 + c];
        const bottom = copy[((y + 1) * width + x) * 4 + c];
        const left = copy[(y * width + (x - 1)) * 4 + c];
        const right = copy[(y * width + (x + 1)) * 4 + c];
        const center = copy[idx + c];

        const val = center * centerWeight + (top + bottom + left + right) * edgeWeight;
        data[idx + c] = Math.min(255, Math.max(0, val));
      }
    }
  }

  ctx.putImageData(imgData, 0, 0);
}

/**
 * Render background onto a target canvas context
 */
export function renderBackground(
  ctx: CanvasRenderingContext2D,
  width: number,
  height: number,
  bgConfig: BackgroundConfig,
  bgImage?: HTMLImageElement | null
) {
  ctx.save();
  ctx.clearRect(0, 0, width, height);

  if (bgConfig.type === 'transparent') {
    ctx.restore();
    return;
  }

  if (bgConfig.type === 'color') {
    ctx.fillStyle = bgConfig.color;
    ctx.fillRect(0, 0, width, height);
  } else if (bgConfig.type === 'gradient') {
    let grad: CanvasGradient;
    const { angle, colors, type } = bgConfig.gradient;
    const c1 = colors[0] || '#ffffff';
    const c2 = colors[1] || '#000000';

    if (type === 'linear') {
      const rad = ((angle - 90) * Math.PI) / 180;
      const x1 = width / 2 - (Math.cos(rad) * width) / 2;
      const y1 = height / 2 - (Math.sin(rad) * height) / 2;
      const x2 = width / 2 + (Math.cos(rad) * width) / 2;
      const y2 = height / 2 + (Math.sin(rad) * height) / 2;
      grad = ctx.createLinearGradient(x1, y1, x2, y2);
    } else {
      const cx = width / 2;
      const cy = height / 2;
      const r = Math.max(width, height) / 2;
      grad = ctx.createRadialGradient(cx, cy, 0, cx, cy, r);
    }

    grad.addColorStop(0, c1);
    grad.addColorStop(1, c2);
    ctx.fillStyle = grad;
    ctx.fillRect(0, 0, width, height);
  } else if (bgConfig.type === 'image' && bgImage) {
    if (bgConfig.imageBlur > 0) {
      ctx.filter = `blur(${bgConfig.imageBlur}px)`;
    }
    const scale = bgConfig.imageScale || 1.0;
    const drawW = bgImage.width * scale;
    const drawH = bgImage.height * scale;
    const offsetX = (width - drawW) / 2 + (bgConfig.imageX / 100) * width;
    const offsetY = (height - drawH) / 2 + (bgConfig.imageY / 100) * height;

    ctx.drawImage(bgImage, offsetX, offsetY, drawW, drawH);
    ctx.filter = 'none';
  }

  ctx.restore();
}

/**
 * Rotate and flip a canvas image, returning a newly dimensioned canvas
 */
export function transformCanvas(
  sourceCanvas: HTMLCanvasElement,
  rotationDegrees: number, // 0, 90, 180, 270
  flipH: boolean,
  flipV: boolean
): HTMLCanvasElement {
  const rad = (rotationDegrees * Math.PI) / 180;
  const isPerpendicular = rotationDegrees % 180 !== 0;

  const newWidth = isPerpendicular ? sourceCanvas.height : sourceCanvas.width;
  const newHeight = isPerpendicular ? sourceCanvas.width : sourceCanvas.height;

  const resultCanvas = createCanvas(newWidth, newHeight);
  const ctx = resultCanvas.getContext('2d')!;

  ctx.save();
  ctx.translate(newWidth / 2, newHeight / 2);
  ctx.rotate(rad);
  ctx.scale(flipH ? -1 : 1, flipV ? -1 : 1);
  ctx.drawImage(sourceCanvas, -sourceCanvas.width / 2, -sourceCanvas.height / 2);
  ctx.restore();

  return resultCanvas;
}

/**
 * Crop a canvas
 */
export function cropCanvas(sourceCanvas: HTMLCanvasElement, cropRect: CropRect): HTMLCanvasElement {
  const result = createCanvas(cropRect.width, cropRect.height);
  const ctx = result.getContext('2d')!;
  ctx.drawImage(
    sourceCanvas,
    cropRect.x,
    cropRect.y,
    cropRect.width,
    cropRect.height,
    0,
    0,
    cropRect.width,
    cropRect.height
  );
  return result;
}
