using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace СollisionMatrix
{
    public class MatrixCreatingViewModel : ObservableObject
    {
        public ObservableCollection<MatrixSelectionLineModel> Selections { get; set; }
        public ObservableCollection<UserControl> UserControlsInWholeMatrix { get; set; }
        public MatrixCreatingViewModel()
        {
            Selections = new ObservableCollection<MatrixSelectionLineModel>();
            var new1 = new MatrixSelectionLineModel()
            {
                NameOfSelection = "AR_Walls",
                SelectionIntersectionTolerance = new List<string>() 
                {
                    "23", "", ""
                }
            };
            var new2 = new MatrixSelectionLineModel()
            {
                NameOfSelection = "AR_Floors",
                SelectionIntersectionTolerance = new List<string>()
                {
                    "", "12", ""
                }
            };
            var new3 = new MatrixSelectionLineModel()
            {
                NameOfSelection = "AR_Doors",
                SelectionIntersectionTolerance = new List<string>()
                {
                    "", "", "5"
                }
            };

            Selections.Add(new1);
            Selections.Add(new2);
            Selections.Add(new3);

            UserControlsInWholeMatrix = new ObservableCollection<UserControl>();
            foreach (MatrixSelectionLineModel model in Selections)
            {
                MatrixSelectionLineViewModel userControlvm = new MatrixSelectionLineViewModel();
                userControlvm.MatrixSelectionLineModel = model;
                userControlvm.RowNum = Selections.IndexOf(model);
                userControlvm.Selections = Selections;
                userControlvm.Selection = model;
                
                userControlvm.ToleranceViews = new ObservableCollection<UserControl>();
                foreach (string tolerance in model.SelectionIntersectionTolerance)
                {

                    MatrixSelectionCellVewModel cellViewModel = new MatrixSelectionCellVewModel()
                    {
                        Tolerance = tolerance
                    };
                    MatrixSelectionCellUserControl cellView = new MatrixSelectionCellUserControl();
                    cellView.DataContext = cellViewModel;

                    userControlvm.ToleranceViews.Add(cellView);
                }


                MatrixSelectionLineUserControl userControl = new MatrixSelectionLineUserControl();
                userControlvm.UserControl_MatrixSelectionLineUserControl = userControl;
                userControlvm.UserControlsInAllMatrixWithLineUserControls = UserControlsInWholeMatrix;
                userControl.DataContext = userControlvm;
                UserControlsInWholeMatrix.Add(userControl);
            };
            

        }
    }
}
