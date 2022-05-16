using System;
using System.Collections.Generic;
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

namespace СollisionMatrix
{
    /// <summary>
    /// Логика взаимодействия для MatrixSelectionLineUserControl.xaml
    /// </summary>
    public partial class MatrixSelectionLineUserControl : UserControl
    {
        public MatrixSelectionLineUserControl()
        {
            InitializeComponent();
        }

        private void UserControl_MouseMove(object sender, MouseEventArgs e)
        {
            MatrixSelectionLineViewModel vm = (MatrixSelectionLineViewModel)this.DataContext;
            vm.ButtonsVisibility = Visibility.Visible;
        }

        private void UserControl_MouseLeave(object sender, MouseEventArgs e)
        {
            MatrixSelectionLineViewModel vm = (MatrixSelectionLineViewModel)this.DataContext;
            vm.ButtonsVisibility = Visibility.Hidden;
        }
    }
}
