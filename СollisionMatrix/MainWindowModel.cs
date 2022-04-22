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
using Forms = System.Windows.Forms;
using OfficeOpenXml;
using System.IO;

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

        private double _matrixWidth;
        public double MatrixWidth
        {
            get { return _matrixWidth; }
            set
            {
                if (_matrixWidth != value)
                {
                    _matrixWidth = value;
                    OnPropertyChanged();
                }
            }
        }

        private double _matrixHeight;
        public double MatrixHeight
        {
            get { return _matrixHeight; }
            set
            {
                if (_matrixHeight != value)
                {
                    _matrixHeight = value;
                    OnPropertyChanged();
                }
            }
        }

        private double _matrixHeaderHeight;
        public double MatrixHeaderHeight
        {
            get { return _matrixHeaderHeight; }
            set
            {
                if (_matrixHeaderHeight != value)
                {
                    _matrixHeaderHeight = value;
                    OnPropertyChanged();
                }
            }
        }

        private double _matrixHeaderWidth;
        public double MatrixHeaderWidth
        {
            get { return _matrixHeaderWidth; }
            set
            {
                if (_matrixHeaderWidth != value)
                {
                    _matrixHeaderWidth = value;
                    OnPropertyChanged();
                }
            }
        }

        private double _matrixCellHorizontalDimension;
        public double MatrixCellHorizontalDimension
        {
            get { return _matrixCellHorizontalDimension; }
            set
            {
                if (_matrixCellHorizontalDimension != value)
                {
                    _matrixCellHorizontalDimension = value;
                    OnPropertyChanged();
                }
            }
        }

        private double _matrixCellVerticalDimension;
        public double MatrixCellVerticalDimension
        {
            get { return _matrixCellVerticalDimension; }
            set
            {
                if (_matrixCellVerticalDimension != value)
                {
                    _matrixCellVerticalDimension = value;
                    OnPropertyChanged();
                }
            }
        }

        public System.Windows.Media.SolidColorBrush Color15 { get; set; } = System.Windows.Media.Brushes.OrangeRed;
        private System.Windows.Media.SolidColorBrush Color30 { get; set; } = System.Windows.Media.Brushes.PaleVioletRed;
        private System.Windows.Media.SolidColorBrush Color50 { get; set; } = System.Windows.Media.Brushes.LightBlue;
        private System.Windows.Media.SolidColorBrush Color80 { get; set; } = System.Windows.Media.Brushes.LightGreen;

        private List<List<string>> ExlsCells { get; set; }

        public ObservableCollection<SelectionSet> SelectionSets { get; set; }
        public ObservableCollection<SelectionSetNode> SelectionSetNodes { get; set; }
        public ObservableCollection<ClashTest> ClashTests { get; set; }

        public MainWindowModel()
        {
            MatrixHeight = 800;
            MatrixWidth = 800;
            MatrixHeaderHeight = 150;
            MatrixHeaderWidth = 150;
            MatrixCellHorizontalDimension = 36;
            MatrixCellVerticalDimension = 26;

            SelectionSets = new ObservableCollection<SelectionSet>();
            SelectionSetNodes = new ObservableCollection<SelectionSetNode>();
            ClashTests = new ObservableCollection<ClashTest>();

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
            int selectionsNumber = 0;

            #region reading xml file
            string pathtoexe = Assembly.GetExecutingAssembly().Location;
            string folderpath = pathtoexe.Replace("СollisionMatrix.exe", "");
            string pathtoxml = folderpath + "selectionsets.xml";

            Forms.OpenFileDialog openFileDialog = new Forms.OpenFileDialog();
            openFileDialog.ShowDialog();
            pathtoxml = openFileDialog.FileName;

            //Xceed.Wpf.Toolkit.MessageBox.Show(pathtoxml, "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);
            XmlDocument xDoc = new XmlDocument();
            try
            {
                xDoc.Load(pathtoxml);
            }
            catch (Exception ex)
            {
                Xceed.Wpf.Toolkit.MessageBox.Show(ex.ToString(), "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            #endregion

            #region creating array SelectionSet from xml file

            XmlElement _exchange_ = xDoc.DocumentElement;
            foreach (XmlNode _selectionsets_ in _exchange_)
            {
                foreach (XmlNode _selectionset_ in _selectionsets_.ChildNodes)
                {
                    SelectionSet selectionSet = new SelectionSet();
                    selectionSet.Name = _selectionset_.Attributes.GetNamedItem("name").InnerText;
                    int divider = 0;
                    divider = selectionSet.Name.IndexOf("_");
                    if (divider != 0)
                    {
                        selectionSet.DraftName = selectionSet.Name.Substring(0, divider);
                    }
                    else
                    {
                        selectionSet.DraftName = "?";
                    }
                    SelectionSets.Add(selectionSet);
                }
                
            }

            /*
            string output = string.Empty;
            foreach (SelectionSet en in SelectionSets)
            {
                output += en.Name + "(" + en.DraftName + ")" + "\n";
            }

            Xceed.Wpf.Toolkit.MessageBox.Show("SelectionSets\n" + output, "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);
            */

            List<string> drafts = new List<string>();
            foreach (SelectionSet ss in SelectionSets)
            {
                bool included = false;
                foreach (string draft in drafts)
                {
                    if (ss.DraftName == draft) included = true;
                }
                if (!included) drafts.Add(ss.DraftName);
            }

            List<string> sortedDrafts = drafts.OrderBy(o => o).ToList();

            List<string> reversedDrafts = new List<string>();
            foreach (string draft in sortedDrafts)
            {
                reversedDrafts.Add(draft);
            }

            Stack<string> stackDrafts = new Stack<string>();

            bool IncludedInList(string input, List<string> inputlist)
            {
                string draft = input.ToLower();
                foreach (string d in inputlist)
                {
                    if (d.ToLower().Contains(draft)) return true;
                }
                return false;
            }

            string GetDraftWith(string input, List<string> inputlist)
            {
                foreach (string d in inputlist)
                {
                    if (IncludedInList(input, inputlist)) return d;
                }
                return "";
            }
            
            //sortedDrafts = stackDrafts.ToList();


            //List<string> sortedDrafts = drafts;
            /*
            string output = string.Empty;
            foreach (string en in sortedDrafts)
            {
                output += en + "\n";
            }

            Xceed.Wpf.Toolkit.MessageBox.Show("Drafts\n" + output, "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);
            */
            List<SelectionSetNode> selectionSetNodes = new List<SelectionSetNode>();
            foreach (string draft in sortedDrafts)
            {
                SelectionSetNode selectionSetNode = new SelectionSetNode();
                selectionSetNode.DraftName = draft;
                selectionSetNode.SelectionSets = new List<SelectionSet>();
                foreach (SelectionSet selectionSet in SelectionSets)
                {
                    if (selectionSet.DraftName == draft) selectionSetNode.SelectionSets.Add(selectionSet);
                }
                selectionSetNodes.Add(selectionSetNode);
            }
            

            foreach (SelectionSetNode ssn in selectionSetNodes)
            {
                SelectionSetNode newSSN = new SelectionSetNode();
                newSSN.DraftName = ssn.DraftName;
                newSSN.SelectionSets = ssn.SelectionSets.OrderBy(o => o.Name).ToList();
                SelectionSetNodes.Add(newSSN);
            }

            /*
            string output = string.Empty;
            foreach (SelectionSetNode ssn in SelectionSetNodes)
            {
                output += ssn.DraftName + "\n";
                foreach (SelectionSet ss in ssn.SelectionSets)
                {
                    output += "   " + ss.Name + "\n";
                }
            }

            Xceed.Wpf.Toolkit.MessageBox.Show("SelectionSets\n" + output, "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);
            */


            #endregion

            selectionsNumber = SelectionSets.Count;
            MatrixHeight = selectionsNumber * MatrixCellVerticalDimension + MatrixHeaderHeight;
            MatrixWidth = selectionsNumber * MatrixCellHorizontalDimension + MatrixHeaderWidth;

            #region see SelectionSet
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
            #endregion

            #region creating matrix headers horizontal
            foreach (SelectionSetNode selSetNode in SelectionSetNodes)
            {
                NodeHView nodeHView = new NodeHView();
                NodeHViewModel nodeHViewModel = (NodeHViewModel)nodeHView.DataContext;
                nodeHViewModel.Header = selSetNode.DraftName;
                foreach (SelectionSet selSet in selSetNode.SelectionSets)
                {
                    SubNodeVView subView = new SubNodeVView();
                    subView.Width = MatrixCellHorizontalDimension;
                    SubNodeVViewModel subViewModel = (SubNodeVViewModel)subView.DataContext;
                    subViewModel.Header = selSet.Name;

                    nodeHView.spSelectionNames.Children.Add(subView);
                }
                ThisView.spSelectionSetsH.Children.Add(nodeHView);
            }
            #endregion

            #region creating matrix headers vertical left
            foreach (SelectionSetNode selSetNode in SelectionSetNodes)
            {
                NodeVView nodeVView = new NodeVView();
                NodeVViewModel nodeVViewModel = (NodeVViewModel)nodeVView.DataContext;
                nodeVViewModel.Header = selSetNode.DraftName;
                foreach (SelectionSet selSet in selSetNode.SelectionSets)
                {
                    SubNodeHView subView = new SubNodeHView();
                    subView.Height = MatrixCellVerticalDimension;
                    SubNodeHViewModel subViewModel = (SubNodeHViewModel)subView.DataContext;
                    subViewModel.Header = selSet.Name;

                    nodeVView.spSelectionNames.Children.Add(subView);
                }
                ThisView.spSelectionSetsV.Children.Add(nodeVView);
            }
            #endregion

            ThisView.borderUpLeft.BorderThickness = new Thickness(0, 0, 1, 1);
            ThisView.borderUpLeft.BorderBrush = System.Windows.Media.Brushes.DarkRed;

        }
        private bool CanCom1Execute(object p) => true;

        public ICommand Com2 { get; set; }
        private void OnCom2Executed(object p)
        {
            ClashTests.Clear();

            #region reading xml file

            string pathtoexe = Assembly.GetExecutingAssembly().Location;
            string folderpath = pathtoexe.Replace("СollisionMatrix.exe", "");
            string pathtoxml = folderpath + "batchtest.xml";

            Forms.OpenFileDialog openFileDialog = new Forms.OpenFileDialog();
            openFileDialog.ShowDialog();
            pathtoxml = openFileDialog.FileName;

            //Xceed.Wpf.Toolkit.MessageBox.Show(pathtoxml, "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);

            XmlDocument xDoc = new XmlDocument();

            try
            {
                xDoc.Load(pathtoxml);
            }
            catch (Exception ex)
            {
                Xceed.Wpf.Toolkit.MessageBox.Show(ex.ToString(), "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
            #endregion

            #region creating array ClashTests from xml file
            
            try
            {
                XmlElement _exchange_ = xDoc.DocumentElement; // <exchange xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" units="m" filename="" filepath="">

                foreach (XmlNode _batchtest_ in _exchange_) // <batchtest name="Report" internal_name="Report" units="m">
                {
                    foreach (XmlNode _clashtests_ in _batchtest_.ChildNodes) // <clashtests>
                    {
                        foreach (XmlNode _clashtest_ in _clashtests_.ChildNodes) // <clashtest name="АР_Вертикальные конструкции - АР_Вертикальные конструкции" test_type="duplicate" status="ok" tolerance="0.000" merge_composites="0">
                        {
                            ClashTest clashTest = new ClashTest();
                            clashTest.Name = _clashtest_.Attributes.GetNamedItem("name").InnerText;
                            clashTest.Tolerance = _clashtest_.Attributes.GetNamedItem("tolerance").InnerText;
                            foreach (XmlNode _elementInClashTest_ in _clashtest_.ChildNodes)
                            {
                                // <left><clashselection selfintersect="0" primtypes="1"><locator>lcop_selection_set_tree/АР_Вертикальные конструкции</locator></clashselection></left>
                                if (_elementInClashTest_.Name == "left")
                                {
                                    foreach (XmlNode _left_ in _elementInClashTest_.ChildNodes)
                                    {
                                        foreach (XmlNode _clashselection_ in _left_.ChildNodes)
                                        {
                                            foreach (XmlNode _locator_ in _clashselection_.ChildNodes)
                                            {
                                                clashTest.SelectionLeftName = _locator_.InnerText.Replace("lcop_selection_set_tree/", "");
                                            }
                                        }
                                    }
                                }
                                if (_elementInClashTest_.Name == "right")
                                {
                                    foreach (XmlNode _right_ in _elementInClashTest_.ChildNodes)
                                    {
                                        foreach (XmlNode _clashselection_ in _right_.ChildNodes)
                                        {
                                            foreach (XmlNode _locator_ in _clashselection_.ChildNodes)
                                            {
                                                clashTest.SelectionRightName = _locator_.InnerText.Replace("lcop_selection_set_tree/", "");
                                            }
                                        }
                                    }
                                }
                                if (_elementInClashTest_.Name == "summary")
                                {
                                    clashTest.SummaryTotal = _elementInClashTest_.Attributes.GetNamedItem("total").InnerText;
                                    clashTest.SummaryActive = _elementInClashTest_.Attributes.GetNamedItem("active").InnerText;
                                    clashTest.SummaryReviewed = _elementInClashTest_.Attributes.GetNamedItem("reviewed").InnerText;
                                    clashTest.SummaryResolved = _elementInClashTest_.Attributes.GetNamedItem("resolved").InnerText;
                                }
                            }
                            ClashTests.Add(clashTest);
                        }
                    }
                }
            } 
            catch (Exception ex)
            {
                Xceed.Wpf.Toolkit.MessageBox.Show(ex.ToString(), "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            #endregion

            #region see ClashTests
            /*
            string output = string.Empty;
            foreach (ClashTest ct in ClashTests)
            {
                output += ct.Name + " -> " + ct.SummaryTotal + "\n";
            }

            Xceed.Wpf.Toolkit.MessageBox.Show("Вот такие тесты\n" + output, "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);
            */
            #endregion

            ExlsCells = new List<List<string>>();

            foreach (SelectionSetNode selectionSetNodeVerticalSet in SelectionSetNodes)
            {
                foreach (SelectionSet selectionSetHName in selectionSetNodeVerticalSet.SelectionSets)
                {
                    DataLineView dataLineView = new DataLineView();

                    List<string> cellsLine = new List<string>(); // create new line in exls

                    foreach (SelectionSetNode selectionSetNodeHorizontalSet in SelectionSetNodes)
                    {
                        foreach (SelectionSet selectionSetVName in selectionSetNodeHorizontalSet.SelectionSets)
                        {
                            DataCellView dataCellView = new DataCellView();
                            dataCellView.Width = MatrixCellHorizontalDimension;
                            dataCellView.Height = MatrixCellVerticalDimension;
                            DataCellViewModel dataCellViewModel = (DataCellViewModel)dataCellView.DataContext;
                            dataCellViewModel.Selection1Name = selectionSetHName.Name;
                            dataCellViewModel.Selection2Name = selectionSetVName.Name;
                            dataCellViewModel.CollisionsTotalNumber = "";
                            dataCellViewModel.CollisionsActiveNumber = "";
                            dataCellViewModel.CollisionsResolvedNumber = "";
                            dataCellViewModel.CollisionsReviewedNumber = "";
                            
                            foreach (ClashTest clashTest in ClashTests)
                            {
                                if (dataCellViewModel.Selection1Name == clashTest.SelectionLeftName)
                                {
                                    if (dataCellViewModel.Selection2Name == clashTest.SelectionRightName)
                                    {
                                        dataCellViewModel.CollisionsTotalNumber = clashTest.SummaryTotal;
                                        dataCellViewModel.CollisionsActiveNumber = clashTest.SummaryActive;
                                        dataCellViewModel.CollisionsReviewedNumber = clashTest.SummaryReviewed;
                                        dataCellViewModel.CollisionsResolvedNumber = clashTest.SummaryResolved;

                                        bool result = double.TryParse(clashTest.Tolerance.Replace(".", ","), out double tolerance);
                                        
                                        if (result)
                                        {
                                            if (tolerance <= 0.08) dataCellView.border.Background = Color80;
                                            if (tolerance <= 0.05) dataCellView.border.Background = Color50;
                                            if (tolerance <= 0.03) dataCellView.border.Background = Color30;
                                            if (tolerance <= 0.015) dataCellView.border.Background = Color15;

                                            dataCellView.tbLeftBracket.Visibility = Visibility.Visible;
                                            dataCellView.tbRightBracket.Visibility = Visibility.Visible;
                                            dataCellView.divider.Visibility = Visibility.Visible;

                                            dataCellView.ToolTip = $"{selectionSetHName.Name} / {selectionSetVName.Name}\nАктивно: {clashTest.SummaryActive}\nИсправлено: {clashTest.SummaryResolved}\nВсего: {clashTest.SummaryTotal}";
                                            
                                        }
                                    }
                                }
                            }
                            
                            dataLineView.spLine.Children.Add(dataCellView);

                            if (dataCellViewModel.CollisionsTotalNumber != "")
                            {
                                bool s1 = int.TryParse(dataCellViewModel.CollisionsTotalNumber, out int total);
                                bool s2 = int.TryParse(dataCellViewModel.CollisionsResolvedNumber, out int resolved);

                                int diff = total - resolved;

                                cellsLine.Add(diff.ToString() +
                                    "(" + dataCellViewModel.CollisionsResolvedNumber + ")" + 
                                    "\n" + dataCellViewModel.CollisionsTotalNumber);  // create cell 
                            }
                            else cellsLine.Add("");
                        }

                    }
                    ThisView.spDataVLine.Children.Add(dataLineView);
                    ExlsCells.Add(cellsLine);
                }
            }

            /*
            string output = string.Empty;
            foreach (List<string> ls in ExlsCells)
            {
                foreach (string cell in ls)
                {
                    output += cell + " ";
                }
                output += "\n";
            }

            Xceed.Wpf.Toolkit.MessageBox.Show("Вот такие тесты\n" + output, "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);
            */
        }
        private bool CanCom2Execute(object p) => true;

        public ICommand Com3 { get; set; }
        private void OnCom3Executed(object p)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelPackage excel = new ExcelPackage();

            Forms.OpenFileDialog openFileDialog = new Forms.OpenFileDialog();
            openFileDialog.ShowDialog();
            string pathtoxls = openFileDialog.FileName;

            if (!File.Exists(pathtoxls))
            {
                try
                {
                    File.Create(pathtoxls).Close();
                }
                catch (Exception ex)
                {
                    Xceed.Wpf.Toolkit.MessageBox.Show("Возможно, нет доступа к папке\n" + ex.ToString(), "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }

            if (File.Exists(pathtoxls))
            {
                try
                {
                    File.Delete(pathtoxls);
                    excel.Workbook.Worksheets.Add("Матрица");
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets["Матрица"];
                    int rowCount = 3;
                    int colCount = 1;
                    foreach (SelectionSetNode ssn in SelectionSetNodes)
                    {
                        foreach (SelectionSet ss in ssn.SelectionSets)
                        {
                            worksheet.Cells[rowCount, colCount].Value = ssn.DraftName;
                            worksheet.Cells[rowCount, colCount + 1].Value = ss.Name;

                            rowCount += 1;
                        }

                    }

                    rowCount = 1;
                    colCount = 3;
                    foreach(SelectionSetNode ssn in SelectionSetNodes)
                    {
                        foreach(SelectionSet ss in ssn.SelectionSets)
                        {
                            worksheet.Cells[rowCount, colCount].Value = ssn.DraftName;
                            worksheet.Cells[rowCount + 1, colCount].Value = ss.Name;
                            worksheet.Cells[rowCount + 1, colCount].Style.TextRotation = 90;

                            colCount += 1;
                        }
                        
                    }
                    rowCount = 3;
                    colCount = 3;
                    foreach (List<string> ls in ExlsCells)
                    {
                        foreach (string cell in ls)
                        {
                            worksheet.Cells[rowCount, colCount].Value = cell;
                            colCount++;
                        }
                        rowCount++;
                        colCount = 3;
                    }

                    for (int i = 1; i <= ExlsCells.Count + 2; i++)
                    {
                        worksheet.Row(i).Style.WrapText = true;
                        worksheet.Row(i).Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Row(i).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }

                    FileInfo excelFile = new FileInfo(pathtoxls);
                    excel.SaveAs(excelFile);
                }
                catch (Exception ex)
                {
                    Xceed.Wpf.Toolkit.MessageBox.Show("Что-то произошло\n" + ex.ToString(), "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

        }
        private bool CanCom3Execute(object p) => true;
    }
}
