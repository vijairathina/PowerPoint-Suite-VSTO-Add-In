using System;
using Microsoft.Office.Tools;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.Assistant.UI
{
    /// <summary>
    /// Manages the reusable CustomTaskPane instance for the Design & Quality Assistant.
    /// </summary>
    public class DesignAssistantTaskPaneManager : IDisposable
    {
        private static DesignAssistantTaskPaneManager _instance;
        public static DesignAssistantTaskPaneManager Instance => _instance ?? (_instance = new DesignAssistantTaskPaneManager());

        private CustomTaskPane _taskPane;
        private DesignAssistantTaskPaneHost _hostControl;

        public CustomTaskPane TaskPane => _taskPane;
        public DesignAssistantViewModel ViewModel => _hostControl?.ViewModel;

        public void EnsureTaskPaneCreated()
        {
            if (_taskPane == null)
            {
                _hostControl = new DesignAssistantTaskPaneHost();
                _hostControl.Initialize(Globals.ThisAddIn.Application);

                _taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(_hostControl, "Design & Quality Assistant");
                _taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Width = 440;
            }
        }

        public void ShowOrActivate()
        {
            EnsureTaskPaneCreated();

            if (!_taskPane.Visible)
            {
                _taskPane.Visible = true;
            }

            // Trigger a slide color refresh and quick status check
            _hostControl.ViewModel?.RefreshColors();
        }

        public void Toggle()
        {
            EnsureTaskPaneCreated();
            _taskPane.Visible = !_taskPane.Visible;
            if (_taskPane.Visible)
            {
                _hostControl.ViewModel?.RefreshColors();
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
