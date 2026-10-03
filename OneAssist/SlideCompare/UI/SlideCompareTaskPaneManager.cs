using System;
using Microsoft.Office.Tools;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.SlideCompare.UI
{
    /// <summary>
    /// Manages the reusable CustomTaskPane instance for the Slide Compare module.
    /// </summary>
    public class SlideCompareTaskPaneManager : IDisposable
    {
        private static SlideCompareTaskPaneManager _instance;
        public static SlideCompareTaskPaneManager Instance => _instance ?? (_instance = new SlideCompareTaskPaneManager());

        private CustomTaskPane _taskPane;
        private SlideCompareTaskPaneHost _hostControl;

        public CustomTaskPane TaskPane => _taskPane;
        public SlideCompareViewModel ViewModel => _hostControl?.ViewModel;

        public void EnsureTaskPaneCreated()
        {
            if (_taskPane == null)
            {
                _hostControl = new SlideCompareTaskPaneHost();
                _hostControl.Initialize(Globals.ThisAddIn.Application);

                _taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(_hostControl, "Slide Compare & Difference");
                _taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Width = 460;
            }
        }

        public void ShowOrActivate()
        {
            EnsureTaskPaneCreated();

            if (!_taskPane.Visible)
            {
                _taskPane.Visible = true;
            }

            _hostControl.ViewModel?.RefreshDestination();
        }

        public void Toggle()
        {
            EnsureTaskPaneCreated();
            _taskPane.Visible = !_taskPane.Visible;
            if (_taskPane.Visible)
            {
                _hostControl.ViewModel?.RefreshDestination();
            }
        }

        public void Dispose()
        {
            if (_taskPane != null)
            {
                try
                {
                    Globals.ThisAddIn.CustomTaskPanes.Remove(_taskPane);
                }
                catch { }
                _taskPane = null;
            }

            _hostControl?.Dispose();
            _hostControl = null;
        }
    }
}
