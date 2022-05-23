using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace СollisionMatrix
{
    public class MatrixSelectionCellViewModel : ObservableObject
    {
        private Models.Clashtest _clashtest;
        public Models.Clashtest Clashtest
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

        private string _tolerance;
        public string Tolerance
        {
            get { return _tolerance; }
            set
            {
                if (_tolerance != value)
                {
                    _tolerance = value;
                    OnPropertyChanged();
                }
            }
        }
    }
}
