using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.UI
{
    /// <summary>
    /// WinForms UserControl container that hosts the WPF DesignAssistantControl inside VSTO CustomTaskPane.
    /// </summary>
    public class DesignAssistantTaskPaneHost : UserControl
    {
        private readonly ElementHost _elementHost;
        private readonly DesignAssistantControl _wpfControl;
        private readonly DesignAssistantViewModel _viewModel;

        public DesignAssistantViewModel ViewModel => _viewModel;

        public DesignAssistantTaskPaneHost()
        {
            _elementHost = new ElementHost
            {
                Dock = DockStyle.Fill
            };

            _viewModel = new DesignAssistantViewModel();
            _wpfControl = new DesignAssistantControl
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
