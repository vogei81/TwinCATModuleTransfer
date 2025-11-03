using System;
using System.Windows;
using System.Windows.Controls;
using TwinCATModuleTransfer.ViewModels;

namespace TwinCATModuleTransfer.Views
{
    public partial class ExportView : UserControl
    {
        public ExportView()
        {
            InitializeComponent();
        }

        private void TreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var vm = this.DataContext as ExportViewModel;
            if (vm != null)
            {
                vm.SelectedNode = e.NewValue as SelectionTreeItem;
            }
        }

        private void TextBox_PreviewDragOver(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                    if (files != null && files.Length == 1 &&
                        files[0].EndsWith(".sln", StringComparison.OrdinalIgnoreCase))
                    {
                        e.Effects = DragDropEffects.Copy;
                    }
                    else
                    {
                        e.Effects = DragDropEffects.None;
                    }
                }
                else
                {
                    e.Effects = DragDropEffects.None;
                }
            }
            catch
            {
                e.Effects = DragDropEffects.None;
            }
            finally
            {
                e.Handled = true;
            }
        }

        private void TextBox_Drop(object sender, DragEventArgs e)
        {
            try
            {
                var vm = this.DataContext as ExportViewModel;
                if (vm == null) return;

                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                    if (files != null && files.Length == 1 &&
                        files[0].EndsWith(".sln", StringComparison.OrdinalIgnoreCase))
                    {
                        vm.SolutionPath = files[0];
                        vm.AddRecentSolution(files[0]);
                    }
                }
            }
            catch
            {
            }
            finally
            {
                e.Handled = true;
            }
        }
    }
}