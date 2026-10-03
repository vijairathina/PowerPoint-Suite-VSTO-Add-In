using System;
using System.Runtime.InteropServices;
using OneAssist.ImageEditor.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.ImageEditor.PowerPointServices
{
    public class ShapeSelectionEventArgs : EventArgs
    {
        public bool IsPictureSelected { get; set; }
        public string Message { get; set; }
    }

    /// <summary>
    /// Listens for PowerPoint selection changes and notifies the Image Editor.
    /// </summary>
    public class ShapeSelectionService : IDisposable
    {
        private PowerPoint.Application _powerPointApp;
        private bool _isHooked;

        public event EventHandler<ShapeSelectionEventArgs> SelectionChanged;

        public void Initialize(PowerPoint.Application app)
        {
            _powerPointApp = app ?? throw new ArgumentNullException(nameof(app));
            if (!_isHooked)
            {
                try
                {
                    _powerPointApp.WindowSelectionChange += OnWindowSelectionChange;
                    _isHooked = true;
                }
                catch
                {
                    // Ignore hook failure in case COM events are restricted
                }
            }
        }

        private void OnWindowSelectionChange(PowerPoint.Selection sel)
        {
            try
            {
                bool isPic = PowerPointImageService.IsPictureSelected(_powerPointApp, out string reason);
                SelectionChanged?.Invoke(this, new ShapeSelectionEventArgs
                {
                    IsPictureSelected = isPic,
                    Message = reason
                });
            }
            catch
            {
                // Silently handle any COM exceptions during background selection events
            }
        }

        public void CheckCurrentSelection()
        {
            try
            {
                bool isPic = PowerPointImageService.IsPictureSelected(_powerPointApp, out string reason);
                SelectionChanged?.Invoke(this, new ShapeSelectionEventArgs
                {
                    IsPictureSelected = isPic,
                    Message = reason
                });
            }
            catch { }
        }

        public void Dispose()
        {
            if (_isHooked && _powerPointApp != null)
            {
                try
                {
                    _powerPointApp.WindowSelectionChange -= OnWindowSelectionChange;
                }
                catch { }
                _isHooked = false;
            }
        }
    }
}
