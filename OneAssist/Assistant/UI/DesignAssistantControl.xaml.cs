using System;
using System.Windows;
using System.Windows.Controls;

namespace OneAssist.Assistant.UI
{
    public partial class DesignAssistantControl : UserControl
    {
        public DesignAssistantControl()
        {
            InitializeComponent();
        }

        private void TabClick(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton rb && rb.Tag != null && int.TryParse(rb.Tag.ToString(), out int tabIdx))
            {
                if (DataContext is DesignAssistantViewModel vm)
                {
                    vm.SelectedTabIndex = tabIdx;
                }
            }
        }

        private void ModeClick(object sender, RoutedEventArgs e)
        {
            if (sender is RadioButton rb && rb.Tag != null && Enum.TryParse(rb.Tag.ToString(), out OneAssist.Assistant.Core.AppMode mode))
            {
                if (DataContext is DesignAssistantViewModel vm)
                {
                    vm.CurrentMode = mode;
                }
            }
        }
    }
}
