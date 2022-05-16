using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

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

        private Visibility _buttonsVisibility;
        public Visibility ButtonsVisibility
        {
            get { return _buttonsVisibility; }
            set
            {
                if (_buttonsVisibility != value)
                {
                    _buttonsVisibility = value;
                    OnPropertyChanged();
                }
            }
        }

        private int _rowNum;
        public int RowNum
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

        public ICommand DoIfButtonDeletePushCommand { get; set; }
        private void OnDoIfButtonDeletePushCommandExecuted(object p)
        {
            //MessageBox.Show($"{RowNum} from {Selections.Count} and {Selection.NameOfSelection}");
            // Удалить ячейку с номером строки из всех строк из модели Пересечений Selections
            //foreach (var ss in Selection.SelectionIntersectionTolerance)
            //{
            //    ss.Remove(ss.ElementAt(RowNum));
            //}
            // Удалить ячейку с номером строки из всех строк из списка вью-строчек на окне
            foreach (var uc in UserControlsInAllMatrixWithLineUserControls)
            {
                MatrixSelectionLineUserControl ucl = (MatrixSelectionLineUserControl)uc;
                MatrixSelectionLineViewModel uclvm = (MatrixSelectionLineViewModel)ucl.DataContext;
                uclvm.ToleranceViews.Remove(uclvm.ToleranceViews.ElementAt(RowNum));
                if (uclvm.RowNum > RowNum) uclvm.RowNum -= 1; 
            }
            Selections.Remove(Selection); // удалить из Общего списка Пересечений Selections из модели, которая будет использована для экспорта в xml в итоге
            UserControlsInAllMatrixWithLineUserControls.Remove(UserControl_MatrixSelectionLineUserControl); // Удалить из списка всех вью-строчек на окне

            

        }
        private bool CanDoIfButtonDeletePushCommandExecute(object p) => true;

        public ObservableCollection<MatrixSelectionLineModel> Selections { get; set; }
        public MatrixSelectionLineModel Selection { get; set; }
        public ObservableCollection<UserControl> UserControlsInAllMatrixWithLineUserControls { get; set; }
        public MatrixSelectionLineUserControl UserControl_MatrixSelectionLineUserControl { get; set; }

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

            ButtonsVisibility = Visibility.Hidden;

            DoIfButtonDeletePushCommand = new RelayCommand(OnDoIfButtonDeletePushCommandExecuted, CanDoIfButtonDeletePushCommandExecute); 
        }
    }
}
