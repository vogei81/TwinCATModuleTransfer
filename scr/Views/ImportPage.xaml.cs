using System.Windows;
using System.Windows.Controls;
using TwinCATModuleTransfer.ViewModels;

namespace TwinCATModuleTransfer.Views
{
    public partial class ImportPage : Page
    {
        public ImportPage() => InitializeComponent();

        private void TreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (DataContext is ImportViewModel vm)
                vm.SelectedTargetParent = e.NewValue as SelectionTreeItem;
        }
    }
}