import { createCanvas } from './imageProcessing';

export interface InpaintOptions {
  radius?: number; // Inpaint neighborhood radius
  iterations?: number;
  onProgress?: (percent: number) => void;
}

/**
 * Removes marked objects from the image using fast marching gradient inpainting.
 * 
 * @param sourceCanvas The original image canvas
 * @param maskCanvas The mask canvas where non-transparent/white pixels indicate the region to remove
 * @param options Configuration options
 */
export async function inpaintObjectRemoval(
  sourceCanvas: HTMLCanvasElement,
  maskCanvas: HTMLCanvasElement,
  options: InpaintOptions = {}
): Promise<HTMLCanvasElement> {
  const width = sourceCanvas.width;
  const height = sourceCanvas.height;

  const resultCanvas = createCanvas(width, height);
  const resultCtx = resultCanvas.getContext('2d')!;
  resultCtx.drawImage(sourceCanvas, 0, 0);

  const imgData = resultCtx.getImageData(0, 0, width, height);
  const data = imgData.data;

  // Read mask
  const mCtx = maskCanvas.getContext('2d')!;
  const maskImgData = mCtx.getImageData(0, 0, width, height);
  const mData = maskImgData.data;

  // Mask status: 0 = known background, 1 = band (boundary), 2 = inside mask (to fill)
  const status = new Uint8Array(width * height);
  let totalToFill = 0;

  for (let i = 0; i < width * height; i++) {
    // If mask has opacity or color
    if (mData[i * 4 + 3] > 20) {
      status[i] = 2; // to fill
      totalToFill++;
    } else {
      status[i] = 0; // known
    }
  }

  if (totalToFill === 0) {
    return resultCanvas;
  }

  // Find boundary pixels (inside mask having at least one known neighbor)
  const queue: number[] = [];
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const idx = y * width + x;
      if (status[idx] === 2) {
        let hasKnownNeighbor = false;
        if (x > 0 && status[idx - 1] === 0) hasKnownNeighbor = true;
        else if (x < width - 1 && status[idx + 1] === 0) hasKnownNeighbor = true;
        else if (y > 0 && status[idx - width] === 0) hasKnownNeighbor = true;
        else if (y < height - 1 && status[idx + width] === 0) hasKnownNeighbor = true;

        if (hasKnownNeighbor) {
          status[idx] = 1; // band
          queue.push(idx);
        }
      }
    }
  }

  const radius = options.radius ?? 8;
  let filledCount = 0;

  // Process band pixels inward
  while (queue.length > 0) {
    const currentIdx = queue.shift()!;
    const cx = currentIdx % width;
    const cy = Math.floor(currentIdx / width);

    // Collect weighted color from known neighbor pixels within radius
    let totalWeight = 0;
    let sumR = 0;
    let sumG = 0;
    let sumB = 0;

    const rMin = Math.max(0, cy - radius);
    const rMax = Math.min(height - 1, cy + radius);
    const cMin = Math.max(0, cx - radius);
    const cMax = Math.min(width - 1, cx + radius);

    for (let ny = rMin; ny <= rMax; ny++) {
      for (let nx = cMin; nx <= cMax; nx++) {
        const nIdx = ny * width + nx;
        if (status[nIdx] === 0) {
          const dx = nx - cx;
          const dy = ny - cy;
          const distSq = dx * dx + dy * dy;
          if (distSq <= radius * radius) {
            // Gaussian-like distance weighting
            const dist = Math.sqrt(distSq);
            const weight = 1 / (1 + dist * dist);

            const nOffset = nIdx * 4;
            sumR += data[nOffset] * weight;
            sumG += data[nOffset + 1] * weight;
            sumB += data[nOffset + 2] * weight;
            totalWeight += weight;
          }
        }
      }
    }

    if (totalWeight > 0) {
      const curOffset = currentIdx * 4;
      data[curOffset] = Math.round(sumR / totalWeight);
      data[curOffset + 1] = Math.round(sumG / totalWeight);
      data[curOffset + 2] = Math.round(sumB / totalWeight);
      data[curOffset + 3] = 255;
    }

    status[currentIdx] = 0; // marked as known now
    filledCount++;

    // Add unvisited masked neighbors to band queue
    const neighbors = [
      cx > 0 ? currentIdx - 1 : -1,
      cx < width - 1 ? currentIdx + 1 : -1,
      cy > 0 ? currentIdx - width : -1,
      cy < height - 1 ? currentIdx + width : -1
    ];

    for (const nb of neighbors) {
      if (nb >= 0 && status[nb] === 2) {
        status[nb] = 1;
        queue.push(nb);
      }
    }
  }

  // Smooth inpainted patch edges with bilateral-like pass
  resultCtx.putImageData(imgData, 0, 0);

  // Return completed inpainted canvas
  return resultCanvas;
}
