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
        private Models.Selectionset _selectionset;
        public Models.Selectionset Selectionset
        {
            get { return _selectionset; }
            set
            {
                if (_selectionset != value)
                {
                    _selectionset = value;
                    OnPropertyChanged();
                }
            }
        }

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

        public ICommand DoIfIClickDeleteButton { get; set; }
        private void OnDoIfIClickDeleteButtonExecuted(object p)
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
            lineViewModel_new.NameOfSelection = "XX_New";
            
            lineViewModel_new.ToleranceViews = new ObservableCollection<UserControl>();
            for (int i = 0; i < UserControlsInAllMatrixWithLineUserControls.Count(); i++)
            {
                MatrixSelectionCellVewModel cell_vm = new MatrixSelectionCellVewModel();
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

                MatrixSelectionCellVewModel cellViewModel_new = new MatrixSelectionCellVewModel();
                cellViewModel_new.Tolerance = string.Empty;
                MatrixSelectionCellUserControl cellUserControl_new = new MatrixSelectionCellUserControl();
                cellUserControl_new.DataContext = cellViewModel_new;

                mslvm.ToleranceViews.Insert(RowNum + 1, cellUserControl_new);

                foreach (UserControl usercontrolcell in mslvm.ToleranceViews)
                {
                    MatrixSelectionCellUserControl mscuc = (MatrixSelectionCellUserControl)usercontrolcell;
                    MatrixSelectionCellVewModel mscvm = (MatrixSelectionCellVewModel)mscuc.DataContext; // cell
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
            MatrixSelectionLineModel = new MatrixSelectionLineModel()
            {
                NameOfSelection = "Example",
                SelectionIntersectionTolerance = new List<string>()
            };

            ToleranceViews = new ObservableCollection<UserControl>();

            ButtonsVisibility = Visibility.Hidden;

            DoIfIClickDeleteButton = new RelayCommand(OnDoIfIClickDeleteButtonExecuted, CanDoIfIClickDeleteButtonExecute);
            DoIfIClickBottomAddButton = new RelayCommand(OnDoIfIClickBottomAddButtonExecuted, CanDoIfIClickBottomAddButtonExecute); 
        }
    }
}
