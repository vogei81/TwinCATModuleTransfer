using System.Windows;
using System.Windows.Controls;
using TwinCATModuleTransfer.ViewModels;

namespace TwinCATModuleTransfer.Views
{
    public partial class ImportView : UserControl
    {
        public ImportView() { InitializeComponent(); }

        private void TreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var vm = this.DataContext as ImportViewModel;
            if (vm != null) vm.SelectedTargetParent = e.NewValue as SelectionTreeItem;
        }
    }
}