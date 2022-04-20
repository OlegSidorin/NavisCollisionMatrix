using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Xml;
using СollisionMatrix.Subviews;
using Xceed.Wpf;
using System.Windows;
using System.Reflection;

namespace СollisionMatrix
{
    public class MainWindowModel : ObservableObject
    {
        private MainWindow _thisView;
        public MainWindow ThisView
        {
            get { return _thisView; }
            set
            {
                if (_thisView != value)
                {
                    _thisView = value;
                    OnPropertyChanged();
                }
            }
        }

        public ObservableCollection<SelectionSet> SelectionSets { get; set; }

        public MainWindowModel()
        {
            SelectionSets = new ObservableCollection<SelectionSet>();
            Commanda = new RelayCommand(OnCommandaExecuted, CanCommandaExecute);
            Com1 = new RelayCommand(OnCom1Executed, CanCom1Execute);
            Com2 = new RelayCommand(OnCom2Executed, CanCom2Execute);
            Com3 = new RelayCommand(OnCom3Executed, CanCom3Execute);

        }

        private string _example;
        public string Example
        {
            get { return _example; }
            set
            {
                if (_example != value)
                {
                    _example = value;
                    OnPropertyChanged();
                }
            }
        }
        public ICommand Commanda { get; set; }
        private void OnCommandaExecuted(object p)
        {

        }
        private bool CanCommandaExecute(object p) => true;

        public ICommand Com1 { get; set; }
        private void OnCom1Executed(object p)
        {
            string pathtoexe = Assembly.GetExecutingAssembly().Location;
            string folderpath = pathtoexe.Replace("СollisionMatrix.exe", "");
            string pathtoxml = folderpath + "selectionsets.xml";
            //Xceed.Wpf.Toolkit.MessageBox.Show(pathtoxml, "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(pathtoxml);
            XmlElement xRoot = xDoc.DocumentElement;
            foreach (XmlElement xnode in xRoot)
            {
                SelectionSet selectionSet = new SelectionSet();
                selectionSet.Selections = new List<string>();

                XmlNode setName = xnode.Attributes.GetNamedItem("name");

                selectionSet.Name = setName.Value;

                foreach (XmlNode childnode in xnode.ChildNodes)
                {
                    selectionSet.Selections.Add(childnode.InnerText);
                }

                SelectionSets.Add(selectionSet);
            }

            /*
            string output = string.Empty;
            foreach (SelectionSet en in SelectionSets)
            {
                output += en.Name + "\n";
                foreach(var ej in en.Selections)
                {
                    output += ej + "\n";
                }
            }

            Xceed.Wpf.Toolkit.MessageBox.Show("Вот такая матрица\n" + output, "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);

            */

            foreach (SelectionSet selSet in SelectionSets)
            {
                NodeHView nodeHView = new NodeHView();
                NodeHViewModel nodeHViewModel = (NodeHViewModel)nodeHView.DataContext;
                nodeHViewModel.Header = selSet.Name;
                foreach (string selName in selSet.Selections)
                {
                    SubNodeVView subView = new SubNodeVView();
                    subView.Width = 20;
                    SubNodeVViewModel subViewModel = (SubNodeVViewModel)subView.DataContext;
                    subViewModel.Header = selName;

                    nodeHView.spSelectionNames.Children.Add(subView);
                }
                ThisView.spSelectionSetsH.Children.Add(nodeHView);
            }
            foreach (SelectionSet selSet in SelectionSets)
            {
                NodeVView nodeVView = new NodeVView();
                NodeVViewModel nodeVViewModel = (NodeVViewModel)nodeVView.DataContext;
                nodeVViewModel.Header = selSet.Name;
                foreach (string selName in selSet.Selections)
                {
                    SubNodeHView subView = new SubNodeHView();
                    subView.Height = 20;
                    SubNodeHViewModel subViewModel = (SubNodeHViewModel)subView.DataContext;
                    subViewModel.Header = selName;

                    nodeVView.spSelectionNames.Children.Add(subView);
                }
                ThisView.spSelectionSetsV.Children.Add(nodeVView);
            }

        }
        private bool CanCom1Execute(object p) => true;

        public ICommand Com2 { get; set; }
        private void OnCom2Executed(object p)
        {
            
            foreach (SelectionSet selVerSet in SelectionSets)
            {
                foreach (string subH in selVerSet.Selections)
                {
                    DataLineView dataLineView = new DataLineView();
                    foreach (SelectionSet selHorSet in SelectionSets)
                    {
                        foreach (string subV in selHorSet.Selections)
                        {
                            DataCellView dataCellView = new DataCellView();
                            dataCellView.Width = 20;
                            dataCellView.Height = 20;
                            dataLineView.spLine.Children.Add(dataCellView);
                        }

                    }
                    ThisView.spDataVLine.Children.Add(dataLineView);
                }
            }

        }
        private bool CanCom2Execute(object p) => true;

        public ICommand Com3 { get; set; }
        private void OnCom3Executed(object p)
        {
            


        }
        private bool CanCom3Execute(object p) => true;
    }
}
