using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix.Mainviews
{
    public class MainCellViewModel : ObservableObject
    {
        private XMLCT.Clashtest _clashtest;
        public XMLCT.Clashtest Clashtest
        {
            get { return _clashtest; }
            set
            {
                if (_clashtest != value)
                {
                    _clashtest = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _presenter;
        public string Presenter
        {
            get { return _presenter; }
            set
            {
                if (_presenter != value)
                {
                    _presenter = value;
                    OnPropertyChanged();
                }
            }
        }
    }
}
