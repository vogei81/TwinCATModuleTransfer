using System.Collections.ObjectModel;

namespace TwinCATModuleTransfer.ViewModels
{
    public class SelectionTreeItem : BaseViewModel
    {
        public string Path { get; set; }
        public string DisplayName { get; set; }
        public bool IsSelectable { get; set; } = true;

        private bool _isChecked;
        public bool IsChecked
        {
            get { return _isChecked; }
            set { _isChecked = value; Raise(); }
        }

        public ObservableCollection<SelectionTreeItem> Children { get; private set; }

        public SelectionTreeItem()
        {
            Children = new ObservableCollection<SelectionTreeItem>();
        }
    }
}
