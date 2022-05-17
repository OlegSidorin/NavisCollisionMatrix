using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace СollisionMatrix
{
    public class MatrixSelectionCellVewModel : ObservableObject
    {
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
