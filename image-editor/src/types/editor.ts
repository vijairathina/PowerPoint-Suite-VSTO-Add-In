export type ToolType = 
  | 'select'
  | 'move'
  | 'crop'
  | 'remove-bg'
  | 'eraser'
  | 'restore'
  | 'mark-remove'
  | 'background'
  | 'adjustments'
  | 'transform';

export type BackgroundType = 'transparent' | 'color' | 'gradient' | 'image';

export interface GradientConfig {
  type: 'linear' | 'radial';
  angle: number; // 0 - 360
  colors: string[]; // e.g. ['#6366f1', '#a855f7']
}

export interface BackgroundConfig {
  type: BackgroundType;
  color: string;
  gradient: GradientConfig;
  imageUrl?: string;
  imageScale: number; // 0.1 to 3.0
  imageX: number; // offset percent
  imageY: number; // offset percent
  imageBlur: number;
}

export interface AdjustmentConfig {
  brightness: number; // -100 to 100 (0 default)
  contrast: number; // -100 to 100 (0 default)
  saturation: number; // -100 to 100 (0 default)
  exposure: number; // -100 to 100 (0 default)
  blur: number; // 0 to 50 px (0 default)
  sharpen: number; // 0 to 100 (0 default)
  opacity: number; // 0 to 100 (100 default)
}

export interface CropRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface BrushSettings {
  size: number; // 1 to 200
  hardness: number; // 0 to 100 (blur/softness)
  opacity: number; // 0 to 100
  feather: number; // 0 to 50
}

export interface ImageMetadata {
  name: string;
  originalWidth: number;
  originalHeight: number;
  currentWidth: number;
  currentHeight: number;
  sizeBytes: number;
  format: string;
}

export interface HistoryStep {
  description: string;
  timestamp: number;
  // Alpha mask data (null if fully opaque)
  maskDataUrl?: string | null;
  // Inpainted or edited base image data
  baseImageDataUrl?: string;
  adjustments: AdjustmentConfig;
  background: BackgroundConfig;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
  cropRect?: CropRect | null;
}
