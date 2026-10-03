import { createCanvas } from './imageProcessing';

export interface BgRemovalOptions {
  tolerance?: number; // 10 to 100, default 35
  feather?: number; // 0 to 20, default 3
  smoothness?: number; // 1 to 5, default 2
}

/**
 * Calculates color distance in Lab/RGB space between two pixels
 */
function colorDistance(r1: number, g1: number, b1: number, r2: number, g2: number, b2: number): number {
  // Weighted Euclidean distance for human perception (redmean formula)
  const rmean = (r1 + r2) / 2;
  const dr = r1 - r2;
  const dg = g1 - g2;
  const db = b1 - b2;
  return Math.sqrt(
    (((512 + rmean) * dr * dr) >> 8) +
    4 * dg * dg +
    (((767 - rmean) * db * db) >> 8)
  );
}

/**
 * Automatically detects and segments foreground from background.
 * Returns an alpha mask canvas where 255 = foreground, 0 = background.
 */
export function removeBackgroundSmart(
  imageCanvas: HTMLCanvasElement,
  options: BgRemovalOptions = {}
): HTMLCanvasElement {
  const width = imageCanvas.width;
  const height = imageCanvas.height;
  const ctx = imageCanvas.getContext('2d')!;
  const imgData = ctx.getImageData(0, 0, width, height);
  const data = imgData.data;

  const tolerance = (options.tolerance ?? 38) * 1.8;
  const feather = options.feather ?? 3;

  // Mask array: 1 = foreground, 0 = background
  const mask = new Uint8Array(width * height);
  mask.fill(1); // default all foreground

  // Sample border pixels to detect background color palette
  const borderSamples: [number, number, number][] = [];
  const stepX = Math.max(1, Math.floor(width / 60));
  const stepY = Math.max(1, Math.floor(height / 60));

  // Top & bottom edges
  for (let x = 0; x < width; x += stepX) {
    const topIdx = x * 4;
    borderSamples.push([data[topIdx], data[topIdx + 1], data[topIdx + 2]]);
    const botIdx = ((height - 1) * width + x) * 4;
    borderSamples.push([data[botIdx], data[botIdx + 1], data[botIdx + 2]]);
  }
  // Left & right edges
  for (let y = 0; y < height; y += stepY) {
    const leftIdx = (y * width) * 4;
    borderSamples.push([data[leftIdx], data[leftIdx + 1], data[leftIdx + 2]]);
    const rightIdx = (y * width + (width - 1)) * 4;
    borderSamples.push([data[rightIdx], data[rightIdx + 1], data[rightIdx + 2]]);
  }

  // Clustered boundary colors (k-means-like representative colors)
  const bgClusters: [number, number, number][] = [];
  for (const sample of borderSamples) {
    let matched = false;
    for (const cluster of bgClusters) {
      if (colorDistance(sample[0], sample[1], sample[2], cluster[0], cluster[1], cluster[2]) < 25) {
        matched = true;
        break;
      }
    }
    if (!matched) {
      bgClusters.push(sample);
      if (bgClusters.length > 12) break;
    }
  }

  // Multi-source Flood Fill from all boundary pixels
  const visited = new Uint8Array(width * height);
  const queue: number[] = [];

  // Seed boundary pixels that match background clusters
  for (let x = 0; x < width; x++) {
    queue.push(x); // top
    queue.push((height - 1) * width + x); // bottom
  }
  for (let y = 1; y < height - 1; y++) {
    queue.push(y * width); // left
    queue.push(y * width + (width - 1)); // right
  }

  let head = 0;
  while (head < queue.length) {
    const idx = queue[head++];
    if (visited[idx]) continue;
    visited[idx] = 1;

    const px = (idx % width);
    const py = Math.floor(idx / width);
    const pOffset = idx * 4;
    const r = data[pOffset];
    const g = data[pOffset + 1];
    const b = data[pOffset + 2];

    // Check if pixel matches any background cluster
    let minBgDist = Infinity;
    for (const cluster of bgClusters) {
      const d = colorDistance(r, g, b, cluster[0], cluster[1], cluster[2]);
      if (d < minBgDist) minBgDist = d;
    }

    // Distance to center (foreground prior: subjects are usually towards center)
    const dx = (px - width / 2) / (width / 2);
    const dy = (py - height / 2) / (height / 2);
    const centerDist = Math.sqrt(dx * dx + dy * dy);
    // Boundary tolerance is higher, center tolerance is stricter
    const effectiveTolerance = tolerance * (0.8 + 0.5 * Math.min(1.2, centerDist));

    if (minBgDist < effectiveTolerance) {
      mask[idx] = 0; // It's background!

      // Push 4-neighbors
      if (px > 0 && !visited[idx - 1]) queue.push(idx - 1);
      if (px < width - 1 && !visited[idx + 1]) queue.push(idx + 1);
      if (py > 0 && !visited[idx - width]) queue.push(idx - width);
      if (py < height - 1 && !visited[idx + width]) queue.push(idx + width);
    }
  }

  // Create alpha mask canvas
  const maskCanvas = createCanvas(width, height);
  const maskCtx = maskCanvas.getContext('2d')!;
  const maskImgData = maskCtx.createImageData(width, height);
  const maskData = maskImgData.data;

  // Fill initial mask
  for (let i = 0; i < width * height; i++) {
    const m = mask[i] ? 255 : 0;
    const offset = i * 4;
    maskData[offset] = 255;
    maskData[offset + 1] = 255;
    maskData[offset + 2] = 255;
    maskData[offset + 3] = m;
  }
  maskCtx.putImageData(maskImgData, 0, 0);

  // Apply feathering / anti-aliasing to edges
  if (feather > 0) {
    const smoothCanvas = createCanvas(width, height);
    const smoothCtx = smoothCanvas.getContext('2d')!;
    smoothCtx.filter = `blur(${feather}px)`;
    smoothCtx.drawImage(maskCanvas, 0, 0);
    return smoothCanvas;
  }

  return maskCanvas;
}

/**
 * Apply eraser or restore brush onto the alpha mask
 */
export function applyMaskBrush(
  maskCanvas: HTMLCanvasElement,
  x: number,
  y: number,
  size: number,
  mode: 'erase' | 'restore',
  hardness: number = 80,
  opacity: number = 100
) {
  const ctx = maskCanvas.getContext('2d')!;
  ctx.save();

  const radius = Math.max(1, size / 2);
  const alpha = Math.min(1, Math.max(0, opacity / 100));

  const grad = ctx.createRadialGradient(x, y, radius * (hardness / 100), x, y, radius);

  if (mode === 'erase') {
    // Destination-out clears the mask (makes transparent)
    ctx.globalCompositeOperation = 'destination-out';
    grad.addColorStop(0, `rgba(0, 0, 0, ${alpha})`);
    grad.addColorStop(1, 'rgba(0, 0, 0, 0)');
    ctx.fillStyle = grad;
    ctx.beginPath();
    ctx.arc(x, y, radius, 0, Math.PI * 2);
    ctx.fill();
  } else {
    // Restore adds opacity back
    ctx.globalCompositeOperation = 'source-over';
    grad.addColorStop(0, `rgba(255, 255, 255, ${alpha})`);
    grad.addColorStop(1, 'rgba(255, 255, 255, 0)');
    ctx.fillStyle = grad;
    ctx.beginPath();
    ctx.arc(x, y, radius, 0, Math.PI * 2);
    ctx.fill();
  }

  ctx.restore();
}

/**
 * Composite the original image with its alpha mask onto an output canvas
 */
export function applyAlphaMask(
  sourceImageCanvas: HTMLCanvasElement,
  maskCanvas: HTMLCanvasElement | null
): HTMLCanvasElement {
  const width = sourceImageCanvas.width;
  const height = sourceImageCanvas.height;
  const output = createCanvas(width, height);
  const ctx = output.getContext('2d')!;

  if (!maskCanvas) {
    ctx.drawImage(sourceImageCanvas, 0, 0);
    return output;
  }

  // Draw source image
  ctx.drawImage(sourceImageCanvas, 0, 0);

  // Apply mask via destination-in
  ctx.globalCompositeOperation = 'destination-in';
  ctx.drawImage(maskCanvas, 0, 0);

  return output;
}
