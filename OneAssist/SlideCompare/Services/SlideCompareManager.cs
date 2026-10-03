using System;
using System.IO;
using System.Windows.Forms;
using OneAssist.SlideCompare.UI;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneAssist.SlideCompare.Services
{
    /// <summary>
    /// Public entry point for launching Slide Compare within any PowerPoint VSTO Add-in.
    /// </summary>
    public static class SlideCompareManager
    {
        public static void Start(PowerPoint.Presentation destinationPresentation = null, string sourceFilePath = null)
        {
            var app = Globals.ThisAddIn?.Application;
            if (destinationPresentation == null && app != null)
            {
                try
                {
                    destinationPresentation = app.ActivePresentation;
                }
                catch { }
            }

            // If source file path not provided, prompt user with file dialog
            if (string.IsNullOrEmpty(sourceFilePath))
            {
                using (var ofd = new OpenFileDialog())
                {
                    ofd.Title = "Select Source PowerPoint Presentation to Compare Against";
                    ofd.Filter = "PowerPoint Presentations (*.pptx;*.pptm;*.ppt)|*.pptx;*.pptm;*.ppt|All Files (*.*)|*.*";
                    ofd.CheckFileExists = true;

                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        sourceFilePath = ofd.FileName;
                    }
                    else
                    {
                        // User cancelled file selection
                        return;
                    }
                }
            }

            // Ensure task pane is visible
            SlideCompareTaskPaneManager.Instance.ShowOrActivate();

            // Populate and trigger comparison in the view model
            var vm = SlideCompareTaskPaneManager.Instance.ViewModel;
            if (vm != null)
            {
                vm.SetPresentations(destinationPresentation, sourceFilePath);
                vm.StartCompareCommand.Execute(null);
            }
        }

        public static void ShowPanel()
        {
            SlideCompareTaskPaneManager.Instance.ShowOrActivate();
        }
    }
}
