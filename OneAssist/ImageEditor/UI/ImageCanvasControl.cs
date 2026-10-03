using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using OneAssist.ImageEditor.Core;

namespace OneAssist.ImageEditor.UI
{
    /// <summary>
    /// Interactive image canvas supporting zoom, pan, brush painting, crop box, and comparison.
    /// </summary>
    public class ImageCanvasControl : Control
    {
        public static readonly DependencyProperty ViewModelProperty =
            DependencyProperty.Register(nameof(ViewModel), typeof(ImageEditorViewModel), typeof(ImageCanvasControl),
                new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsRender, OnViewModelChanged));

        public ImageEditorViewModel ViewModel
        {
            get => (ImageEditorViewModel)GetValue(ViewModelProperty);
            set => SetValue(ViewModelProperty, value);
        }

        private static void OnViewModelChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var control = (ImageCanvasControl)d;
            if (e.OldValue is ImageEditorViewModel oldVm)
            {
                oldVm.PropertyChanged -= control.OnVmPropertyChanged;
            }
            if (e.NewValue is ImageEditorViewModel newVm)
            {
                newVm.PropertyChanged += control.OnVmPropertyChanged;
            }
            control.InvalidateVisual();
        }

        private void OnVmPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(ImageEditorViewModel.CurrentPreviewSource) ||
                e.PropertyName == nameof(ImageEditorViewModel.OriginalPreviewSource) ||
                e.PropertyName == nameof(ImageEditorViewModel.ShowBeforeAfterComparison) ||
                e.PropertyName == nameof(ImageEditorViewModel.ComparisonSplitPercent) ||
                e.PropertyName == nameof(ImageEditorViewModel.Zoom) ||
                e.PropertyName == nameof(ImageEditorViewModel.ActiveTool) ||
                e.PropertyName == nameof(ImageEditorViewModel.BrushSize) ||
                e.PropertyName == nameof(ImageEditorViewModel.CropRatio))
            {
                if (e.PropertyName == nameof(ImageEditorViewModel.CropRatio) && ViewModel?.Document?.HasImage == true)
                {
                    ResetCropToRatio();
                }
                InvalidateVisual();
            }
        }

        private Point _panOffset = new Point(0, 0);
        private Point _lastMousePos;
        private Point _currentMousePos;
        private bool _isPanning;
        private bool _isPainting;
        private bool _isDraggingSplit;
        private bool _showBrushCursor;

        // Crop state
        private bool _isDraggingCrop;
        private int _cropHandle = -1; // -1 = body move, 0=NW, 1=N, 2=NE, 3=E, 4=SE, 5=S, 6=SW, 7=W
        private Rect _cropBoxNorm = new Rect(0.1, 0.1, 0.8, 0.8); // Normalized 0..1 relative to image

        public ImageCanvasControl()
        {
            ClipToBounds = true;
            Focusable = true;
            Cursor = Cursors.Cross;

            Loaded += (s, e) => ResetView();
            SizeChanged += (s, e) => InvalidateVisual();
        }

        public void ResetView()
        {
            _panOffset = new Point(0, 0);
            InvalidateVisual();
        }

        private void ResetCropToRatio()
        {
            if (ViewModel?.Document == null || !ViewModel.Document.HasImage) return;

            double imgW = ViewModel.Document.Width;
            double imgH = ViewModel.Document.Height;
            if (imgW <= 0 || imgH <= 0) return;

            double targetRatio = imgW / imgH;
            switch (ViewModel.CropRatio)
            {
                case CropAspectRatio.Ratio1x1: targetRatio = 1.0; break;
                case CropAspectRatio.Ratio4x3: targetRatio = 4.0 / 3.0; break;
                case CropAspectRatio.Ratio16x9: targetRatio = 16.0 / 9.0; break;
                case CropAspectRatio.OriginalRatio: targetRatio = imgW / imgH; break;
            }

            double cropW = 0.8;
            double cropH = cropW / (targetRatio * (imgH / imgW));
            if (cropH > 0.8)
            {
                cropH = 0.8;
                cropW = cropH * (targetRatio * (imgH / imgW));
            }

            double cropX = (1.0 - cropW) / 2.0;
            double cropY = (1.0 - cropH) / 2.0;
            _cropBoxNorm = new Rect(cropX, cropY, cropW, cropH);
            SyncCropToViewModel();
        }

        private void SyncCropToViewModel()
        {
            if (ViewModel?.Document == null || !ViewModel.Document.HasImage) return;

            double imgW = ViewModel.Document.Width;
            double imgH = ViewModel.Document.Height;

            double pxX = Math.Max(0, Math.Min(imgW - 1, _cropBoxNorm.X * imgW));
            double pxY = Math.Max(0, Math.Min(imgH - 1, _cropBoxNorm.Y * imgH));
            double pxW = Math.Max(10, Math.Min(imgW - pxX, _cropBoxNorm.Width * imgW));
            double pxH = Math.Max(10, Math.Min(imgH - pxY, _cropBoxNorm.Height * imgH));

            ViewModel.CropRectangle = new Rect(pxX, pxY, pxW, pxH);
        }

        protected override void OnMouseEnter(MouseEventArgs e)
        {
            base.OnMouseEnter(e);
            _showBrushCursor = true;
            InvalidateVisual();
        }

        protected override void OnMouseLeave(MouseEventArgs e)
        {
            base.OnMouseLeave(e);
            _showBrushCursor = false;
            _isPainting = false;
            _isPanning = false;
            _isDraggingSplit = false;
            _isDraggingCrop = false;
            InvalidateVisual();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            _currentMousePos = e.GetPosition(this);

            if (_isPanning)
            {
                Vector delta = _currentMousePos - _lastMousePos;
                _panOffset = new Point(_panOffset.X + delta.X, _panOffset.Y + delta.Y);
                _lastMousePos = _currentMousePos;
                InvalidateVisual();
                return;
            }

            if (_isDraggingSplit && ViewModel != null && ActualWidth > 0)
            {
                double split = Math.Max(5.0, Math.Min(95.0, (_currentMousePos.X / ActualWidth) * 100.0));
                ViewModel.ComparisonSplitPercent = split;
                InvalidateVisual();
                return;
            }

            if (_isDraggingCrop && ViewModel?.ActiveTool == ToolType.Crop)
            {
                var imgRect = GetImageDisplayRect();
                if (imgRect.Width > 0 && imgRect.Height > 0)
                {
                    double dx = (_currentMousePos.X - _lastMousePos.X) / imgRect.Width;
                    double dy = (_currentMousePos.Y - _lastMousePos.Y) / imgRect.Height;

                    if (_cropHandle == -1) // Move crop box
                    {
                        double newX = Math.Max(0, Math.Min(1.0 - _cropBoxNorm.Width, _cropBoxNorm.X + dx));
                        double newY = Math.Max(0, Math.Min(1.0 - _cropBoxNorm.Height, _cropBoxNorm.Y + dy));
                        _cropBoxNorm = new Rect(newX, newY, _cropBoxNorm.Width, _cropBoxNorm.Height);
                    }
                    else if (_cropHandle == 4) // SE Corner
                    {
                        double newW = Math.Max(0.05, Math.Min(1.0 - _cropBoxNorm.X, _cropBoxNorm.Width + dx));
                        double newH = Math.Max(0.05, Math.Min(1.0 - _cropBoxNorm.Y, _cropBoxNorm.Height + dy));
                        _cropBoxNorm = new Rect(_cropBoxNorm.X, _cropBoxNorm.Y, newW, newH);
                    }
                    else if (_cropHandle == 0) // NW Corner
                    {
                        double newX = Math.Max(0, Math.Min(_cropBoxNorm.Right - 0.05, _cropBoxNorm.X + dx));
                        double newY = Math.Max(0, Math.Min(_cropBoxNorm.Bottom - 0.05, _cropBoxNorm.Y + dy));
                        _cropBoxNorm = new Rect(newX, newY, _cropBoxNorm.Right - newX, _cropBoxNorm.Bottom - newY);
                    }

                    SyncCropToViewModel();
                    _lastMousePos = _currentMousePos;
                    InvalidateVisual();
                    return;
                }
            }

            if (_isPainting && ViewModel?.Document?.AlphaMask != null && ViewModel.CurrentPreviewSource != null)
            {
                PaintAtPosition(_currentMousePos, _lastMousePos);
                _lastMousePos = _currentMousePos;
            }

            // Update cursor based on tool
            if (ViewModel?.ActiveTool == ToolType.Crop)
            {
                Cursor = Cursors.SizeAll;
            }
            else if (ViewModel?.ActiveTool == ToolType.Select)
            {
                Cursor = Cursors.Hand;
            }
            else
            {
                Cursor = Cursors.Cross;
            }

            _lastMousePos = _currentMousePos;
            InvalidateVisual();
        }

        protected override void OnMouseDown(MouseButtonEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();
            _lastMousePos = e.GetPosition(this);
            _currentMousePos = _lastMousePos;

            // Pan with middle mouse button or Alt+LeftClick or Select tool
            if (e.MiddleButton == MouseButtonState.Pressed ||
                (e.LeftButton == MouseButtonState.Pressed && (Keyboard.IsKeyDown(Key.LeftAlt) || ViewModel?.ActiveTool == ToolType.Select)))
            {
                _isPanning = true;
                CaptureMouse();
                Cursor = Cursors.SizeAll;
                return;
            }

            // Before/After split drag
            if (ViewModel != null && ViewModel.ShowBeforeAfterComparison)
            {
                double splitX = (ViewModel.ComparisonSplitPercent / 100.0) * ActualWidth;
                if (Math.Abs(_currentMousePos.X - splitX) < 18)
                {
                    _isDraggingSplit = true;
                    CaptureMouse();
                    return;
                }
            }

            // Crop Tool Drag
            if (ViewModel?.ActiveTool == ToolType.Crop && e.LeftButton == MouseButtonState.Pressed)
            {
                var imgRect = GetImageDisplayRect();
                var screenCropRect = new Rect(
                    imgRect.X + _cropBoxNorm.X * imgRect.Width,
                    imgRect.Y + _cropBoxNorm.Y * imgRect.Height,
                    _cropBoxNorm.Width * imgRect.Width,
                    _cropBoxNorm.Height * imgRect.Height);

                Point p = _currentMousePos;
                // Check SE corner handle
                if (Math.Abs(p.X - screenCropRect.Right) < 16 && Math.Abs(p.Y - screenCropRect.Bottom) < 16)
                {
                    _cropHandle = 4; // SE
                }
                // Check NW corner handle
                else if (Math.Abs(p.X - screenCropRect.Left) < 16 && Math.Abs(p.Y - screenCropRect.Top) < 16)
                {
                    _cropHandle = 0; // NW
                }
                else if (screenCropRect.Contains(p))
                {
                    _cropHandle = -1; // move body
                }
                else
                {
                    // Start new crop from click position
                    double normX = Math.Max(0, Math.Min(1, (p.X - imgRect.X) / imgRect.Width));
                    double normY = Math.Max(0, Math.Min(1, (p.Y - imgRect.Y) / imgRect.Height));
                    _cropBoxNorm = new Rect(normX, normY, 0.1, 0.1);
                    _cropHandle = 4;
                }

                _isDraggingCrop = true;
                CaptureMouse();
                return;
            }

            // Painting tools: Eraser, Restore, Mark & Remove
            if (e.LeftButton == MouseButtonState.Pressed && ViewModel != null)
            {
                if (ViewModel.ActiveTool == ToolType.Eraser ||
                    ViewModel.ActiveTool == ToolType.Restore ||
                    ViewModel.ActiveTool == ToolType.MarkRemove)
                {
                    _isPainting = true;
                    CaptureMouse();
                    PaintAtPosition(_currentMousePos, _currentMousePos);
                }
            }
        }

        protected override void OnMouseUp(MouseButtonEventArgs e)
        {
            base.OnMouseUp(e);

            if (_isPainting)
            {
                _isPainting = false;
                ReleaseMouseCapture();

                if (ViewModel?.ActiveTool == ToolType.MarkRemove)
                {
                    // Invalidate visual to show painted mark
                    InvalidateVisual();
                }
                else if (ViewModel?.ActiveTool == ToolType.Eraser)
                {
                    ViewModel.Document.UpdateMask(ViewModel.Document.AlphaMask, "Erase");
                    ViewModel.UpdatePreview();
                }
                else if (ViewModel?.ActiveTool == ToolType.Restore)
                {
                    ViewModel.Document.UpdateWorkingImage(ViewModel.Document.WorkingImage, "Restore");
                    ViewModel.UpdatePreview();
                }
            }

            if (_isPanning)
            {
                _isPanning = false;
                ReleaseMouseCapture();
            }

            if (_isDraggingSplit)
            {
                _isDraggingSplit = false;
                ReleaseMouseCapture();
            }

            if (_isDraggingCrop)
            {
                _isDraggingCrop = false;
                ReleaseMouseCapture();
                SyncCropToViewModel();
            }

            InvalidateVisual();
        }

        protected override void OnMouseWheel(MouseWheelEventArgs e)
        {
            base.OnMouseWheel(e);
            if (ViewModel == null) return;

            double zoomFactor = e.Delta > 0 ? 1.15 : 0.85;
            ViewModel.Zoom = Math.Max(0.2, Math.Min(4.0, ViewModel.Zoom * zoomFactor));
            InvalidateVisual();
        }

        private void PaintAtPosition(Point p0, Point p1)
        {
            if (ViewModel?.Document == null || ViewModel.CurrentPreviewSource == null) return;

            var imgRect = GetImageDisplayRect();
            if (imgRect.Width <= 0 || imgRect.Height <= 0) return;

            // Map canvas screen coordinates to image pixel coordinates
            double scaleX = ViewModel.Document.Width / imgRect.Width;
            double scaleY = ViewModel.Document.Height / imgRect.Height;

            int imgX0 = (int)Math.Round((p0.X - imgRect.X) * scaleX);
            int imgY0 = (int)Math.Round((p0.Y - imgRect.Y) * scaleY);
            int imgX1 = (int)Math.Round((p1.X - imgRect.X) * scaleX);
            int imgY1 = (int)Math.Round((p1.Y - imgRect.Y) * scaleY);

            int radius = (int)Math.Max(1, Math.Round((ViewModel.BrushSize / 2.0) * scaleX));
            float hardness = ViewModel.BrushHardness / 100.0f;
            float opacity = ViewModel.BrushOpacity / 100.0f;

            if (ViewModel.ActiveTool == ToolType.MarkRemove)
            {
                // Paint into RemovalMask (255 = marked to remove)
                ViewModel.Document.RemovalMask?.PaintLine(imgX1, imgY1, imgX0, imgY0, radius, 255, hardness, opacity);
                InvalidateVisual();
            }
            else if (ViewModel.ActiveTool == ToolType.Eraser)
            {
                // Erase alpha in AlphaMask (0 = transparent)
                ViewModel.Document.AlphaMask?.PaintLine(imgX1, imgY1, imgX0, imgY0, radius, 0, hardness, opacity);
                ViewModel.UpdatePreview();
            }
            else if (ViewModel.ActiveTool == ToolType.Restore)
            {
                // Restore actual RGB and alpha pixels from OriginalImage!
                ViewModel.Document.RestoreLineFromOriginal(imgX1, imgY1, imgX0, imgY0, radius, hardness, opacity);
                ViewModel.UpdatePreview();
            }
        }

        private Rect GetImageDisplayRect()
        {
            if (ViewModel?.CurrentPreviewSource == null || ActualWidth <= 0 || ActualHeight <= 0)
                return Rect.Empty;

            double imgW = ViewModel.CurrentPreviewSource.PixelWidth;
            double imgH = ViewModel.CurrentPreviewSource.PixelHeight;
            if (imgW <= 0 || imgH <= 0) return Rect.Empty;

            double zoom = ViewModel.Zoom;
            double availW = Math.Max(20, ActualWidth - 24);
            double availH = Math.Max(20, ActualHeight - 24);

            double fitScale = Math.Min(availW / imgW, availH / imgH);
            fitScale = Math.Min(1.0, fitScale); // Do not artificially upscale unless zoom > 1

            double displayW = imgW * fitScale * zoom;
            double displayH = imgH * fitScale * zoom;

            double displayX = (ActualWidth - displayW) / 2.0 + _panOffset.X;
            double displayY = (ActualHeight - displayH) / 2.0 + _panOffset.Y;

            return new Rect(displayX, displayY, displayW, displayH);
        }

        protected override void OnRender(DrawingContext dc)
        {
            base.OnRender(dc);

            // Dark neutral canvas background
            dc.DrawRectangle(new SolidColorBrush(Color.FromRgb(30, 30, 30)), null, new Rect(0, 0, ActualWidth, ActualHeight));

            if (ViewModel == null || !ViewModel.HasImage || ViewModel.CurrentPreviewSource == null)
            {
                return;
            }

            var imgRect = GetImageDisplayRect();
            if (imgRect.IsEmpty) return;

            // Draw drop shadow / border around image
            var shadowRect = new Rect(imgRect.X - 1, imgRect.Y - 1, imgRect.Width + 2, imgRect.Height + 2);
            dc.DrawRectangle(null, new Pen(Brushes.Black, 2), shadowRect);

            // 1. Draw image
            if (ViewModel.ShowBeforeAfterComparison && ViewModel.OriginalPreviewSource != null)
            {
                double splitX = (ViewModel.ComparisonSplitPercent / 100.0) * ActualWidth;

                // Left side: Original Image clipped
                dc.PushClip(new RectangleGeometry(new Rect(0, 0, splitX, ActualHeight)));
                dc.DrawImage(ViewModel.OriginalPreviewSource, imgRect);
                dc.Pop();

                // Right side: Edited Image clipped
                dc.PushClip(new RectangleGeometry(new Rect(splitX, 0, ActualWidth - splitX, ActualHeight)));
                dc.DrawImage(ViewModel.CurrentPreviewSource, imgRect);
                dc.Pop();

                // Draw vertical divider bar
                var splitPen = new Pen(Brushes.White, 2);
                dc.DrawLine(splitPen, new Point(splitX, 0), new Point(splitX, ActualHeight));

                // Draw slider handle circle
                dc.DrawEllipse(Brushes.White, new Pen(new SolidColorBrush(Color.FromRgb(210, 71, 38)), 2), new Point(splitX, ActualHeight / 2), 14, 14);
            }
            else
            {
                // Normal Preview Image
                dc.DrawImage(ViewModel.CurrentPreviewSource, imgRect);
            }

            // 2. Mark & Remove painted overlay
            if (ViewModel.ActiveTool == ToolType.MarkRemove && ViewModel.Document?.RemovalMask != null)
            {
                var rMask = ViewModel.Document.RemovalMask;
                int mw = rMask.Width;
                int mh = rMask.Height;

                // Create lightweight visual overlay for marked areas
                var overlayBitmap = new WriteableBitmap(mw, mh, 96, 96, PixelFormats.Bgra32, null);
                byte[] pixels = new byte[mw * mh * 4];

                bool hasMarkedPixels = false;
                for (int i = 0; i < rMask.Data.Length; i++)
                {
                    byte v = rMask.Data[i];
                    if (v > 20)
                    {
                        hasMarkedPixels = true;
                        int px = i * 4;
                        pixels[px + 0] = 68;  // B
                        pixels[px + 1] = 68;  // G
                        pixels[px + 2] = 239; // R
                        pixels[px + 3] = (byte)(v * 0.65); // Alpha
                    }
                }

                if (hasMarkedPixels)
                {
                    overlayBitmap.WritePixels(new Int32Rect(0, 0, mw, mh), pixels, mw * 4, 0);
                    overlayBitmap.Freeze();
                    dc.DrawImage(overlayBitmap, imgRect);
                }
            }

            // 3. Interactive Crop Rectangle Overlay
            if (ViewModel.ActiveTool == ToolType.Crop)
            {
                var screenCropRect = new Rect(
                    imgRect.X + _cropBoxNorm.X * imgRect.Width,
                    imgRect.Y + _cropBoxNorm.Y * imgRect.Height,
                    _cropBoxNorm.Width * imgRect.Width,
                    _cropBoxNorm.Height * imgRect.Height);

                // Dim areas outside crop box
                var combinedGeom = new CombinedGeometry(
                    GeometryCombineMode.Exclude,
                    new RectangleGeometry(new Rect(0, 0, ActualWidth, ActualHeight)),
                    new RectangleGeometry(screenCropRect));
                dc.DrawGeometry(new SolidColorBrush(Color.FromArgb(140, 0, 0, 0)), null, combinedGeom);

                // White border around crop box
                var cropPen = new Pen(Brushes.White, 1.5)
                {
                    DashStyle = DashStyles.Dash
                };
                dc.DrawRectangle(null, cropPen, screenCropRect);

                // Rule-of-thirds grid
                var gridPen = new Pen(new SolidColorBrush(Color.FromArgb(80, 255, 255, 255)), 1);
                double thirdW = screenCropRect.Width / 3.0;
                double thirdH = screenCropRect.Height / 3.0;
                dc.DrawLine(gridPen, new Point(screenCropRect.Left + thirdW, screenCropRect.Top), new Point(screenCropRect.Left + thirdW, screenCropRect.Bottom));
                dc.DrawLine(gridPen, new Point(screenCropRect.Left + 2 * thirdW, screenCropRect.Top), new Point(screenCropRect.Left + 2 * thirdW, screenCropRect.Bottom));
                dc.DrawLine(gridPen, new Point(screenCropRect.Left, screenCropRect.Top + thirdH), new Point(screenCropRect.Right, screenCropRect.Top + thirdH));
                dc.DrawLine(gridPen, new Point(screenCropRect.Left, screenCropRect.Top + 2 * thirdH), new Point(screenCropRect.Right, screenCropRect.Top + 2 * thirdH));

                // 4 Corner Grab Handles
                double handleSize = 8;
                var handleBrush = Brushes.White;
                var handleBorder = new Pen(new SolidColorBrush(Color.FromRgb(210, 71, 38)), 1.5);

                dc.DrawRectangle(handleBrush, handleBorder, new Rect(screenCropRect.Left - handleSize / 2, screenCropRect.Top - handleSize / 2, handleSize, handleSize));
                dc.DrawRectangle(handleBrush, handleBorder, new Rect(screenCropRect.Right - handleSize / 2, screenCropRect.Top - handleSize / 2, handleSize, handleSize));
                dc.DrawRectangle(handleBrush, handleBorder, new Rect(screenCropRect.Left - handleSize / 2, screenCropRect.Bottom - handleSize / 2, handleSize, handleSize));
                dc.DrawRectangle(handleBrush, handleBorder, new Rect(screenCropRect.Right - handleSize / 2, screenCropRect.Bottom - handleSize / 2, handleSize, handleSize));
            }

            // 4. Draw Brush Cursor
            if (_showBrushCursor && (ViewModel.ActiveTool == ToolType.Eraser ||
                                    ViewModel.ActiveTool == ToolType.Restore ||
                                    ViewModel.ActiveTool == ToolType.MarkRemove))
            {
                double brushR = Math.Max(2, ViewModel.BrushSize / 2.0);
                Brush fill = ViewModel.ActiveTool == ToolType.MarkRemove ?
                    new SolidColorBrush(Color.FromArgb(90, 239, 68, 68)) : // translucent red
                    (ViewModel.ActiveTool == ToolType.Restore ?
                        new SolidColorBrush(Color.FromArgb(80, 59, 130, 246)) : // translucent blue
                        new SolidColorBrush(Color.FromArgb(60, 255, 255, 255))); // translucent white

                Pen pen = new Pen(Brushes.White, 1.5);
                pen.DashStyle = DashStyles.Dash;

                dc.DrawEllipse(fill, pen, _currentMousePos, brushR, brushR);
            }
        }
    }
}
