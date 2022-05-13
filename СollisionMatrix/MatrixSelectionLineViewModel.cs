using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace СollisionMatrix
{
    public class MatrixSelectionLineViewModel : ObservableObject
    {
        private MatrixSelectionLineModel _matrixSelectionLineModel;
        public MatrixSelectionLineModel MatrixSelectionLineModel
        {
            get { return _matrixSelectionLineModel; }
            set
            {
                if (_matrixSelectionLineModel != value)
                {
                    _matrixSelectionLineModel = value;
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<UserControl> ToleranceViews { get; set; }

        // as proxy property for binding
        public string NameOfSelection
        {
            get { return MatrixSelectionLineModel.NameOfSelection; }
            set
            {
                if (MatrixSelectionLineModel.NameOfSelection != value)
                {
                    MatrixSelectionLineModel.NameOfSelection = value;
                    OnPropertyChanged();
                }
            }
        }

        public MatrixSelectionLineViewModel()
        {
            MatrixSelectionLineModel = new MatrixSelectionLineModel()
            {
                NameOfSelection = "Example",
                SelectionIntersectionTolerance = new List<string>()
            };

            ToleranceViews = new ObservableCollection<UserControl>();

            
        }
    }
}
