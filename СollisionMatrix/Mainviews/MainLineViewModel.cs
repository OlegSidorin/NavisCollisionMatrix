using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace СollisionMatrix.Mainviews
{
    public class MainLineViewModel : ObservableObject
    {
        private string _nameOfSelection;
        public string NameOfSelection
        {
            get { return _nameOfSelection; }
            set
            {
                if (_nameOfSelection != value)
                {
                    _nameOfSelection = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _rowNum;
        public string RowNum
        {
            get { return _rowNum; }
            set
            {
                if (_rowNum != value)
                {
                    _rowNum = value;
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<UserControl> CellViews { get; set; }
    }
}
