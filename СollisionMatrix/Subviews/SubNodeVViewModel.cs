using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix.Subviews
{
    public class SubNodeVViewModel : ObservableObject
    {
        private string _header;
        public string Header
        {
            get { return _header; }
            set
            {
                if (_header != value)
                {
                    _header = value;
                    OnPropertyChanged();
                }
            }
        }
    }
}
