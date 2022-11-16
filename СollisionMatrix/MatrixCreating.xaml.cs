using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace СollisionMatrix
{
    /// <summary>
    /// Логика взаимодействия для MatrixCreating.xaml
    /// </summary>
    public partial class MatrixCreating : Window
    {
        public MatrixCreating()
        {
            InitializeComponent();
        }

        private void GridSplitter2_DragCompleted(object sender, System.Windows.Controls.Primitives.DragCompletedEventArgs e)
        {
            string str = "";
            MatrixCreatingViewModel VM = (MatrixCreatingViewModel)this.DataContext;
            
            var c = new ObservableCollection<MatrixSelectionLineUserControl>();
            foreach (UserControl uc in VM.UserControlsInWholeMatrix)
            {
                MatrixSelectionLineUserControl uc1 = (MatrixSelectionLineUserControl)uc;
                MatrixSelectionLineViewModel vm1 = (MatrixSelectionLineViewModel)uc1.DataContext;
                str += vm1.HeaderWidth + ", ";
                vm1.HeaderWidth = VM.WidthColumn - 65;
            }

            //MessageBox.Show($"{VM.WidthColumn}" + "\n" + str);
        }

        private void Grid1_MouseMove(object sender, MouseEventArgs e)
        {
            GridSplitter1.Visibility = Visibility.Visible;
            GridSplitter2.Visibility = Visibility.Visible;
        }

        private void Grid1_MouseLeave(object sender, MouseEventArgs e)
        {
            GridSplitter1.Visibility = Visibility.Collapsed;
            GridSplitter2.Visibility = Visibility.Collapsed;
        }

        //private void ButtonReadXML_Click(object sender, RoutedEventArgs e)
        //{
        //    MatrixCreatingViewModel VM = (MatrixCreatingViewModel)DataContext;
        //    VM.PerformReadXML();
        //}

        //private void ButtonSaveXML_Click(object sender, RoutedEventArgs e)
        //{
        //    MatrixCreatingViewModel VM = (MatrixCreatingViewModel)DataContext;
        //    VM.PerformSaveXML();
        //}
    }
}
