using System.Windows;
using System.Windows.Controls;
using TwinCATModuleTransfer.ViewModels;

namespace TwinCATModuleTransfer.Views
{
    public partial class ExportPage : Page
    {
        public ExportPage() => InitializeComponent();

        private void TreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (DataContext is ExportViewModel vm)
                vm.SelectedNode = e.NewValue as SelectionTreeItem;
        }
    }
}