using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace TwinCATModuleTransfer.ViewModels
{
    public abstract class BaseViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void Raise([CallerMemberName] string prop = null)
        {
            var h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(prop));
        }
    }
}
