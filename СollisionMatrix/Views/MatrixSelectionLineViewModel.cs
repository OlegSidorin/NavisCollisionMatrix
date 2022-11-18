using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace СollisionMatrix
{
    public class MatrixSelectionLineViewModel : ObservableObject
    {
        public double headerWidth;
        public double HeaderWidth { get { return headerWidth; } set { headerWidth = value; OnPropertyChanged(); } }

        private SS.Selectionset jselectionset;
        public SS.Selectionset JSelectionset { get { return jselectionset; } set { jselectionset = value; OnPropertyChanged(); } }

        private MatrixSelectionLineModel _matrixSelectionLineModel;
        public MatrixSelectionLineModel MatrixSelectionLineModel { get { return _matrixSelectionLineModel; } set { _matrixSelectionLineModel = value; OnPropertyChanged(); } }

        private Visibility _buttonsVisibility;
        public Visibility ButtonsVisibility { get { return _buttonsVisibility; } set { _buttonsVisibility = value; OnPropertyChanged(); } }

        private int _rowNum;
        public int RowNum { get { return _rowNum; } set{ _rowNum = value; OnPropertyChanged(); } }

        public ICommand DoIfIClickDownButton { get; set; }
        private void OnDoIfIClickDownButtonExecuted(object p)
        {
            var numOfRow = RowNum; // запишем исходное значение, вдруг поменяется еще во время расстановок 

            UserControl userControl = UserControlsInAllMatrixWithLineUserControls.ElementAt(numOfRow);
            UserControl userControl2 = UserControlsInSelectionNameUserControls.ElementAt(numOfRow);

            if (numOfRow < UserControlsInAllMatrixWithLineUserControls.Count() - 1)
            {
                // Поменять местами сверху вниз строки в матрице
                UserControlsInAllMatrixWithLineUserControls.RemoveAt(numOfRow);
                UserControlsInAllMatrixWithLineUserControls.Insert(numOfRow + 1, userControl);

                // Меняем вверху заголовки наименования между собой слева направо
                UserControlsInSelectionNameUserControls.RemoveAt(numOfRow);
                UserControlsInSelectionNameUserControls.Insert(numOfRow + 1, userControl2);

                // Меняем нумерацию строк в матрице 
                int i = 0;
                foreach (var uc in UserControlsInAllMatrixWithLineUserControls)
                {
                    MatrixSelectionLineUserControl msluc = (MatrixSelectionLineUserControl)uc;
                    MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;
                    mslvm.RowNum = i;
                    i++;
                }

                for (int k = 0; k < UserControlsInAllMatrixWithLineUserControls.Count; k++)
                {
                    // меняем на строчке k местами два квадратика
                    userControl = UserControlsInAllMatrixWithLineUserControls.ElementAt(k);

                    MatrixSelectionLineUserControl uc = (MatrixSelectionLineUserControl)userControl;
                    MatrixSelectionLineViewModel vm = (MatrixSelectionLineViewModel)uc.DataContext;
                    UserControl tvAtNunOfRow2 = vm.ToleranceViews.ElementAt(numOfRow);
                    vm.ToleranceViews.RemoveAt(numOfRow);
                    vm.ToleranceViews.Insert(numOfRow + 1, tvAtNunOfRow2);
                }
            }
        }
        private bool CanDoIfIClickDownButtonExecute(object p) => true;

        public ICommand DoIfIClickUpButton { get; set; }
        private void OnDoIfIClickUpButtonExecuted(object p)
        {
            var numOfRow = RowNum; // запишем исходное значение, вдруг поменяется еще во время расстановок 

            UserControl userControl = UserControlsInAllMatrixWithLineUserControls.ElementAt(numOfRow);
            UserControl userControl2 = UserControlsInSelectionNameUserControls.ElementAt(numOfRow);

            // Поменять местами снизу вверх строки в матрице
            UserControlsInAllMatrixWithLineUserControls.RemoveAt(numOfRow);
            UserControlsInAllMatrixWithLineUserControls.Insert(numOfRow - 1, userControl);

            // Меняем вверху заголовки наименования между собой справа налево
            UserControlsInSelectionNameUserControls.RemoveAt(numOfRow);
            UserControlsInSelectionNameUserControls.Insert(numOfRow - 1, userControl2);

            // Меняем нумерацию строк в матрице 
            int i = 0;
            foreach (var uc in UserControlsInAllMatrixWithLineUserControls)
            {
                MatrixSelectionLineUserControl msluc = (MatrixSelectionLineUserControl)uc;
                MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;
                mslvm.RowNum = i;
                i++;
            }

            for (int k = 0; k < UserControlsInAllMatrixWithLineUserControls.Count; k++)
            {
                // меняем на строчке k местами два квадратика
                userControl = UserControlsInAllMatrixWithLineUserControls.ElementAt(k);

                MatrixSelectionLineUserControl uc = (MatrixSelectionLineUserControl)userControl;
                MatrixSelectionLineViewModel vm = (MatrixSelectionLineViewModel)uc.DataContext;
                UserControl tvAtNunOfRow2 = vm.ToleranceViews.ElementAt(numOfRow);
                vm.ToleranceViews.RemoveAt(numOfRow);
                vm.ToleranceViews.Insert(numOfRow - 1, tvAtNunOfRow2);
            }

        }
        private bool CanDoIfIClickUpButtonExecute(object p)
        {
            if (RowNum < 1) return false;
            else return true;
        }

        public ICommand DoIfIClickDeleteButton { get; set; }
        private void OnDoIfIClickDeleteButtonExecuted(object p)
        {
            //MessageBox.Show($"{RowNum} from {Selections.Count} and {Selection.NameOfSelection}");
            // Удалить ячейку с номером строки из всех строк из модели Пересечений Selections
            //foreach (var SS in Selection.SelectionIntersectionTolerance)
            //{
            //    SS.Remove(SS.ElementAt(RowNum));
            //}
            // Удалить ячейку с номером строки из всех строк из списка вью-строчек на окне
            foreach (var uc in UserControlsInAllMatrixWithLineUserControls)
            {
                MatrixSelectionLineUserControl ucl = (MatrixSelectionLineUserControl)uc;
                MatrixSelectionLineViewModel uclvm = (MatrixSelectionLineViewModel)ucl.DataContext;
                if (RowNum < uclvm.ToleranceViews.Count()) uclvm.ToleranceViews.Remove(uclvm.ToleranceViews.ElementAt(RowNum));
                if (uclvm.RowNum > RowNum) uclvm.RowNum -= 1; 
            }
            //Selections.Remove(Selection); // удалить из Общего списка Пересечений Selections из модели, которая будет использована для экспорта в xml в итоге
            UserControlsInAllMatrixWithLineUserControls.Remove(UserControl_MatrixSelectionLineUserControl); // Удалить из списка всех вью-строчек на окне
            MatrixSelectionNameUserControl userControlForDeleting = null;
            foreach (var uc in UserControlsInSelectionNameUserControls)
            {
                MatrixSelectionNameUserControl msnuc = (MatrixSelectionNameUserControl)uc;
                MatrixSelectionLineViewModel uclvm = (MatrixSelectionLineViewModel)msnuc.DataContext;
                if (uclvm.NameOfSelection == NameOfSelection) userControlForDeleting = msnuc;
            }
            if (userControlForDeleting != null) UserControlsInSelectionNameUserControls.Remove(userControlForDeleting);
        }
        private bool CanDoIfIClickDeleteButtonExecute(object p) => true;

        public ICommand DoIfIClickBottomAddButton { get; set; }
        private void OnDoIfIClickBottomAddButtonExecuted(object p)
        {
            // creating new view user control line
            int indexOfNewRow = RowNum + 1;
            MatrixSelectionLineViewModel lineViewModel_new = new MatrixSelectionLineViewModel();
            lineViewModel_new.RowNum = indexOfNewRow;
            lineViewModel_new.NameOfSelection = "_АР_ | Стены, Перекрытия";
            lineViewModel_new.HeaderWidth = HeaderWidth;
            
            lineViewModel_new.ToleranceViews = new ObservableCollection<UserControl>();
            for (int i = 0; i < UserControlsInAllMatrixWithLineUserControls.Count(); i++)
            {
                MatrixSelectionCellViewModel cell_vm = new MatrixSelectionCellViewModel();
                cell_vm.Tolerance = string.Empty;
                MatrixSelectionCellUserControl cell_uc = new MatrixSelectionCellUserControl();
                cell_uc.DataContext = cell_vm;
                lineViewModel_new.ToleranceViews.Add(cell_uc);
            }
            lineViewModel_new.ButtonsVisibility = Visibility.Hidden;
            lineViewModel_new.UserControlsInAllMatrixWithLineUserControls = UserControlsInAllMatrixWithLineUserControls;
            lineViewModel_new.UserControlsInSelectionNameUserControls = UserControlsInSelectionNameUserControls;

            MatrixSelectionLineUserControl lineUserControl_new = new MatrixSelectionLineUserControl();
            lineViewModel_new.UserControl_MatrixSelectionLineUserControl = lineUserControl_new;
            lineUserControl_new.DataContext = lineViewModel_new;

            MatrixSelectionNameUserControl selectionNameUserControl_new = new MatrixSelectionNameUserControl();
            selectionNameUserControl_new.DataContext = lineViewModel_new;

            UserControlsInAllMatrixWithLineUserControls.Insert(indexOfNewRow, lineUserControl_new);
            UserControlsInSelectionNameUserControls.Insert(indexOfNewRow, selectionNameUserControl_new);

            int iterator = 0;
            foreach (UserControl usercontrolline in UserControlsInAllMatrixWithLineUserControls)
            {
                MatrixSelectionLineUserControl msluc = (MatrixSelectionLineUserControl)usercontrolline;
                MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext; // line

                mslvm.RowNum = iterator;

                MatrixSelectionCellViewModel cellViewModel_new = new MatrixSelectionCellViewModel();
                cellViewModel_new.Tolerance = string.Empty;
                MatrixSelectionCellUserControl cellUserControl_new = new MatrixSelectionCellUserControl();
                cellUserControl_new.DataContext = cellViewModel_new;

                mslvm.ToleranceViews.Insert(RowNum + 1, cellUserControl_new);

                foreach (UserControl usercontrolcell in mslvm.ToleranceViews)
                {
                    MatrixSelectionCellUserControl mscuc = (MatrixSelectionCellUserControl)usercontrolcell;
                    MatrixSelectionCellViewModel mscvm = (MatrixSelectionCellViewModel)mscuc.DataContext; // cell
                }
                iterator += 1;
            }

        }
        private bool CanDoIfIClickBottomAddButtonExecute(object p) => true;

        public ObservableCollection<MatrixSelectionLineModel> Selections { get; set; }
        public MatrixSelectionLineModel Selection { get; set; }
        public ObservableCollection<UserControl> UserControlsInAllMatrixWithLineUserControls { get; set; }
        public ObservableCollection<UserControl> UserControlsInSelectionNameUserControls { get; set; }
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
            HeaderWidth = 120;
            MatrixSelectionLineModel = new MatrixSelectionLineModel()
            {
                NameOfSelection = "Example",
                SelectionIntersectionTolerance = new List<string>()
            };

            ToleranceViews = new ObservableCollection<UserControl>();

            ButtonsVisibility = Visibility.Hidden;

            DoIfIClickDeleteButton = new RelayCommand(OnDoIfIClickDeleteButtonExecuted, CanDoIfIClickDeleteButtonExecute);
            DoIfIClickBottomAddButton = new RelayCommand(OnDoIfIClickBottomAddButtonExecuted, CanDoIfIClickBottomAddButtonExecute);
            DoIfIClickUpButton = new RelayCommand(OnDoIfIClickUpButtonExecuted, CanDoIfIClickUpButtonExecute);
            DoIfIClickDownButton = new RelayCommand(OnDoIfIClickDownButtonExecuted, CanDoIfIClickDownButtonExecute);


        }
    }
}
