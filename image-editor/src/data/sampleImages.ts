/**
 * Generates sample images dynamically on canvas so no external network call or asset missing occurs!
 */

export interface SampleImageItem {
  id: string;
  name: string;
  category: string;
  description: string;
  dataUrl: string;
}

export function generateSampleImages(): SampleImageItem[] {
  // 1. Product: Sleek Headphones
  const canvas1 = document.createElement('canvas');
  canvas1.width = 800;
  canvas1.height = 800;
  const ctx1 = canvas1.getContext('2d')!;

  // Background: Studio soft pastel gradient
  const bgGrad1 = ctx1.createLinearGradient(0, 0, 800, 800);
  bgGrad1.addColorStop(0, '#f1f5f9');
  bgGrad1.addColorStop(1, '#cbd5e1');
  ctx1.fillStyle = bgGrad1;
  ctx1.fillRect(0, 0, 800, 800);

  // Soft studio shadow
  ctx1.save();
  ctx1.beginPath();
  ctx1.ellipse(400, 680, 220, 35, 0, 0, Math.PI * 2);
  ctx1.fillStyle = 'rgba(15, 23, 42, 0.18)';
  ctx1.filter = 'blur(16px)';
  ctx1.fill();
  ctx1.restore();

  // Headphone Arc
  ctx1.lineWidth = 28;
  ctx1.strokeStyle = '#0f172a';
  ctx1.lineCap = 'round';
  ctx1.beginPath();
  ctx1.arc(400, 380, 180, Math.PI, 0);
  ctx1.stroke();

  // Headband cushion
  ctx1.lineWidth = 14;
  ctx1.strokeStyle = '#3b82f6';
  ctx1.beginPath();
  ctx1.arc(400, 375, 175, Math.PI * 1.15, Math.PI * 1.85);
  ctx1.stroke();

  // Left & Right Earcups
  const drawEarCup = (x: number, y: number, color: string) => {
    ctx1.save();
    ctx1.translate(x, y);
    ctx1.fillStyle = '#1e293b';
    ctx1.beginPath();
    ctx1.roundRect(-40, -80, 80, 160, 40);
    ctx1.fill();
    // Inner accent
    ctx1.fillStyle = color;
    ctx1.beginPath();
    ctx1.roundRect(-25, -60, 50, 120, 25);
    ctx1.fill();
    // Core silver ring
    ctx1.strokeStyle = '#94a3b8';
    ctx1.lineWidth = 4;
    ctx1.beginPath();
    ctx1.arc(0, 0, 22, 0, Math.PI * 2);
    ctx1.stroke();
    ctx1.restore();
  };

  drawEarCup(220, 460, '#2563eb');
  drawEarCup(580, 460, '#2563eb');

  // 2. Portrait / Avatar Graphic
  const canvas2 = document.createElement('canvas');
  canvas2.width = 800;
  canvas2.height = 800;
  const ctx2 = canvas2.getContext('2d')!;

  // Background: Warm sunset
  const bgGrad2 = ctx2.createLinearGradient(0, 0, 0, 800);
  bgGrad2.addColorStop(0, '#fef3c7');
  bgGrad2.addColorStop(1, '#fed7aa');
  ctx2.fillStyle = bgGrad2;
  ctx2.fillRect(0, 0, 800, 800);

  // Decorative background circle
  ctx2.fillStyle = '#fde68a';
  ctx2.beginPath();
  ctx2.arc(400, 400, 280, 0, Math.PI * 2);
  ctx2.fill();

  // Person Silhouette / Character
  // Shoulders
  ctx2.fillStyle = '#4338ca';
  ctx2.beginPath();
  ctx2.ellipse(400, 720, 220, 140, 0, 0, Math.PI * 2);
  ctx2.fill();

  // Neck
  ctx2.fillStyle = '#f87171';
  ctx2.fillRect(360, 450, 80, 100);

  // Face
  ctx2.beginPath();
  ctx2.arc(400, 420, 110, 0, Math.PI * 2);
  ctx2.fill();

  // Hair
  ctx2.fillStyle = '#1e1b4b';
  ctx2.beginPath();
  ctx2.arc(400, 390, 125, Math.PI * 0.9, Math.PI * 2.1);
  ctx2.fill();

  // Sunglasses
  ctx2.fillStyle = '#0f172a';
  ctx2.beginPath();
  ctx2.roundRect(320, 400, 70, 40, 12);
  ctx2.roundRect(410, 400, 70, 40, 12);
  ctx2.fill();
  ctx2.fillRect(385, 412, 30, 8);

  // 3. Landscape with an unwanted object to test "Mark & Remove"
  const canvas3 = document.createElement('canvas');
  canvas3.width = 900;
  canvas3.height = 600;
  const ctx3 = canvas3.getContext('2d')!;

  // Sky
  const skyGrad = ctx3.createLinearGradient(0, 0, 0, 380);
  skyGrad.addColorStop(0, '#38bdf8');
  skyGrad.addColorStop(1, '#bae6fd');
  ctx3.fillStyle = skyGrad;
  ctx3.fillRect(0, 0, 900, 380);

  // Sun
  ctx3.fillStyle = '#fef08a';
  ctx3.beginPath();
  ctx3.arc(750, 120, 60, 0, Math.PI * 2);
  ctx3.fill();

  // Rolling green hills
  ctx3.fillStyle = '#16a34a';
  ctx3.beginPath();
  ctx3.moveTo(0, 380);
  ctx3.bezierCurveTo(250, 300, 450, 420, 900, 340);
  ctx3.lineTo(900, 600);
  ctx3.lineTo(0, 600);
  ctx3.closePath();
  ctx3.fill();

  // Foreground hill
  ctx3.fillStyle = '#15803d';
  ctx3.beginPath();
  ctx3.moveTo(0, 460);
  ctx3.bezierCurveTo(300, 400, 600, 520, 900, 440);
  ctx3.lineTo(900, 600);
  ctx3.lineTo(0, 600);
  ctx3.closePath();
  ctx3.fill();

  // Unwanted object: Watermark text & red banner
  ctx3.fillStyle = 'rgba(239, 68, 68, 0.9)';
  ctx3.beginPath();
  ctx3.roundRect(320, 240, 260, 70, 12);
  ctx3.fill();

  ctx3.fillStyle = '#ffffff';
  ctx3.font = 'bold 24px Inter, sans-serif';
  ctx3.textAlign = 'center';
  ctx3.textBaseline = 'middle';
  ctx3.fillText('WATERMARK', 450, 275);

  return [
    {
      id: 'headphones',
      name: 'Studio Headphones',
      category: 'Product Background Removal',
      description: 'Test removing solid background & replacing with custom gradients',
      dataUrl: canvas1.toDataURL('image/png'),
    },
    {
      id: 'portrait',
      name: 'Stylized Portrait',
      category: 'Portrait Editing',
      description: 'Test background extraction, custom colors, and edge refinement',
      dataUrl: canvas2.toDataURL('image/png'),
    },
    {
      id: 'landscape',
      name: 'Landscape with Watermark',
      category: 'Mark & Remove / Inpaint',
      description: 'Highlight the red watermark banner with Mark & Remove to erase it seamlessly',
      dataUrl: canvas3.toDataURL('image/png'),
    }
  ];
}
