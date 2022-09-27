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
using System.Windows.Navigation;
using System.Windows.Shapes;
using СollisionMatrix.Mainviews;

namespace СollisionMatrix
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            MainWindowModel mainWindowModel = (MainWindowModel)DataContext;
            mainWindowModel.ThisView = this;
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

        private void GridSplitter2_DragCompleted(object sender, System.Windows.Controls.Primitives.DragCompletedEventArgs e)
        {
            string str = "";
            MainWindowModel VM = (MainWindowModel)this.DataContext;

            var c = new ObservableCollection<MatrixSelectionLineUserControl>();

            foreach (UserControl uc in VM.LineUserControls)
            {
                MainLineUserControl uc1 = (MainLineUserControl)uc;
                MainLineViewModel vm1 = (MainLineViewModel)uc1.DataContext;
                str += vm1.HeaderWidth + ", ";
                vm1.HeaderWidth = VM.WidthColumn - 35;
            }

            //MessageBox.Show($"{VM.WidthColumn}" + "\n" + str);
        }
    }
}
