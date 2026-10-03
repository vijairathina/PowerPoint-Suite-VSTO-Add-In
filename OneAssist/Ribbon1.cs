using Microsoft.Office.Tools.Ribbon;
using OneAssist.ImageEditor.PowerPointServices;
using System;

namespace OneAssist
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            TaskPaneManager.Instance.ShowOrActivate();
        }

        private void btnDesignAssistant_Click(object sender, RibbonControlEventArgs e)
        {
            OneAssist.Assistant.UI.DesignAssistantTaskPaneManager.Instance.ShowOrActivate();
        }

        private async void btnQuickScan_Click(object sender, RibbonControlEventArgs e)
        {
            var mgr = OneAssist.Assistant.UI.DesignAssistantTaskPaneManager.Instance;
            mgr.ShowOrActivate();
            if (mgr.ViewModel != null)
            {
                await mgr.ViewModel.RunAnalysisAsync();
            }
        }

        private void btnSlideCompare_Click(object sender, RibbonControlEventArgs e)
        {
            OneAssist.SlideCompare.Services.SlideCompareManager.Start();
        }
    }
}
