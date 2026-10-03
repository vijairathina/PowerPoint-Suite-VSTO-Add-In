using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using OneAssist.ImageEditor.UI;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.ImageEditor.PowerPointServices
{
    /// <summary>
    /// Manages the single reusable CustomTaskPane instance for the Image Editor.
    /// </summary>
    public class TaskPaneManager : IDisposable
    {
        private static TaskPaneManager _instance;
        public static TaskPaneManager Instance => _instance ?? (_instance = new TaskPaneManager());

        private CustomTaskPane _taskPane;
        private ImageEditorTaskPaneHost _hostControl;

        public CustomTaskPane TaskPane => _taskPane;
        public ImageEditorViewModel ViewModel => _hostControl?.ViewModel;

        public void EnsureTaskPaneCreated()
        {
            if (_taskPane == null)
            {
                _hostControl = new ImageEditorTaskPaneHost();
                _hostControl.Initialize(Globals.ThisAddIn.Application);

                _taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(_hostControl, "Image Editor");
                _taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Width = 420;
                _taskPane.VisibleChanged += (s, e) =>
                {
                    if (_taskPane.Visible)
                    {
                        _hostControl.ViewModel?.CheckSelection();
                    }
                };
            }
        }

        public void ShowOrActivate()
        {
            EnsureTaskPaneCreated();

            if (!_taskPane.Visible)
            {
                _taskPane.Visible = true;
            }

            // Check selection and load picture if available
            _hostControl.ViewModel?.CheckSelection();
        }

        public void Toggle()
        {
            EnsureTaskPaneCreated();
            _taskPane.Visible = !_taskPane.Visible;
            if (_taskPane.Visible)
            {
                _hostControl.ViewModel?.CheckSelection();
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
