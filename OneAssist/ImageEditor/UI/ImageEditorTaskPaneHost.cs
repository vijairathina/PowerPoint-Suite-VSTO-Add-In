using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.ImageEditor.UI
{
    /// <summary>
    /// WinForms UserControl container that hosts the WPF ImageEditorControl inside VSTO CustomTaskPane.
    /// </summary>
    public class ImageEditorTaskPaneHost : UserControl
    {
        private readonly ElementHost _elementHost;
        private readonly ImageEditorControl _wpfControl;
        private readonly ImageEditorViewModel _viewModel;

        public ImageEditorViewModel ViewModel => _viewModel;

        public ImageEditorTaskPaneHost()
        {
            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill
            };

            _viewModel = new ImageEditorViewModel();
            _wpfControl = new ImageEditorControl
            {
                DataContext = _viewModel
            };

            _elementHost.Child = _wpfControl;
            Controls.Add(_elementHost);
        }

        public void Initialize(PowerPoint.Application app)
        {
            _viewModel.Initialize(app);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _viewModel?.Dispose();
                _elementHost?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
