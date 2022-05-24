using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix.Mainviews
{
    public class MainCellViewModel : ObservableObject
    {
        private Models.Summary _summary;
        public Models.Summary Summary
        {
            get { return _summary; }
            set
            {
                if (_summary != value)
                {
                    _summary = value;
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
