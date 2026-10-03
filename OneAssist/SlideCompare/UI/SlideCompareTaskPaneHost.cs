using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.SlideCompare.UI
{
    /// <summary>
    /// WinForms UserControl container hosting the WPF SlideCompareControl inside a VSTO CustomTaskPane.
    /// </summary>
    public class SlideCompareTaskPaneHost : UserControl
    {
        private readonly ElementHost _elementHost;
        private readonly SlideCompareControl _wpfControl;
        private readonly SlideCompareViewModel _viewModel;

        public SlideCompareViewModel ViewModel => _viewModel;

        public SlideCompareTaskPaneHost()
        {
            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill
            };

            _viewModel = new SlideCompareViewModel();
            _wpfControl = new SlideCompareControl
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
                _elementHost?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
