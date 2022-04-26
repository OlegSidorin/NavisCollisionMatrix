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

        public ObservableCollection<SelectionSet> SelectionSets { get; set; }
        public ObservableCollection<SelectionSetNode> SelectionSetNodes { get; set; }
        public ObservableCollection<ClashTest> ClashTests { get; set; }
        public ObservableCollection<ClashTest> ClashTestsTotal { get; set; }
        public ObservableCollection<string> Drafts { get; set; }

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
            ClashTestsTotal = new ObservableCollection<ClashTest>();
            Drafts = new ObservableCollection<string>();

            Commanda = new RelayCommand(OnCommandaExecuted, CanCommandaExecute); // example
            ImportXMLSelectionSets = new RelayCommand(OnImportXMLSelectionSetsExecuted, CanImportXMLSelectionSetsExecute);
            ImportXMLClashTests = new RelayCommand(OnImportXMLClashTestsExecuted, CanImportXMLClashTestsExecute);
            ExcelExport = new RelayCommand(OnExcelExportExecuted, CanExcelExportExecute);

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

        public ICommand ImportXMLSelectionSets { get; set; }
        private void OnImportXMLSelectionSetsExecuted(object p)
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

            foreach (string draft in sortedDrafts)
            {
                Drafts.Add(draft);
            };

            List<SelectionSetNode> selectionSetNodes = new List<SelectionSetNode>();
            foreach (string draft in Drafts)
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
        private bool CanImportXMLSelectionSetsExecute(object p) => true;

        public ICommand ImportXMLClashTests { get; set; }
        private void OnImportXMLClashTestsExecuted(object p)
        {
            ClashTests.Clear();
            ClashTestsTotal.Clear();

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

                                                int divider = 0;
                                                divider = clashTest.SelectionLeftName.IndexOf("_");
                                                if (divider != 0)
                                                {
                                                    clashTest.DraftLeftName = clashTest.SelectionLeftName.Substring(0, divider);
                                                }
                                                else
                                                {
                                                    clashTest.DraftLeftName = "?";
                                                }
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

                                                int divider = 0;
                                                divider = clashTest.SelectionRightName.IndexOf("_");
                                                if (divider != 0)
                                                {
                                                    clashTest.DraftRightName = clashTest.SelectionRightName.Substring(0, divider);
                                                }
                                                else
                                                {
                                                    clashTest.DraftRightName = "?";
                                                }
                                            }
                                        }
                                    }
                                }
                                if (_elementInClashTest_.Name == "summary")
                                {
                                    clashTest.SummaryTotal = _elementInClashTest_.Attributes.GetNamedItem("total").InnerText;
                                    clashTest.SummaryNovy = _elementInClashTest_.Attributes.GetNamedItem("new").InnerText;
                                    clashTest.SummaryActive = _elementInClashTest_.Attributes.GetNamedItem("active").InnerText;
                                    clashTest.SummaryReviewed = _elementInClashTest_.Attributes.GetNamedItem("reviewed").InnerText;
                                    clashTest.SummaryResolved = _elementInClashTest_.Attributes.GetNamedItem("resolved").InnerText;
                                    clashTest.SummaryApproved = _elementInClashTest_.Attributes.GetNamedItem("approved").InnerText;
                                }
                            }
                            ClashTests.Add(clashTest);
                        }
                    }
                }
                ThisView.notes.Visibility = Visibility.Visible;
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
                output += ct.Name + "(" + ct.SelectionLeftName + " / " + ct.SelectionRightName + ")" + " -> " + ct.SummaryTotal + "\n";
            }

            Xceed.Wpf.Toolkit.MessageBox.Show("Вот такие тесты\n" + output, "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);
            */
            #endregion

            #region creating array ClashTestsTotal from ClashTests

            foreach (ClashTest ct in ClashTests)
            {
                foreach (string draftV in Drafts)
                {
                    foreach (string draftH in Drafts)
                    {
                        int totalTotal = 0;
                        int totalResolved = 0;
                        int totalApproved = 0;
                        int totalReviewed = 0;
                        int totalActive = 0;
                        int totalNovy = 0;

                        int clashTestsNumber = 0;

                        foreach (ClashTest clashTest in ClashTests)
                        {
                            if (clashTest.DraftLeftName == draftV && clashTest.DraftRightName == draftH)
                            {
                                bool resultTotal = int.TryParse(clashTest.SummaryTotal, out int total);
                                bool resultResolved = int.TryParse(clashTest.SummaryResolved, out int resolved);
                                bool resultApproved = int.TryParse(clashTest.SummaryApproved, out int approved);
                                bool resultReviewed = int.TryParse(clashTest.SummaryReviewed, out int reviewed);
                                bool resultActive = int.TryParse(clashTest.SummaryActive, out int active);
                                bool resultNovy = int.TryParse(clashTest.SummaryNovy, out int novy);

                                totalTotal += total;
                                totalResolved += resolved;
                                totalApproved += approved;
                                totalReviewed += reviewed;
                                totalActive += active;
                                totalNovy += novy;

                                clashTestsNumber++;
                            }
                        }

                        if (clashTestsNumber > 0)
                        {
                            ClashTest clashTestOnDrafts = new ClashTest()
                            {
                                Name = $"{draftV} / {draftH}",
                                DraftLeftName = draftV,
                                DraftRightName = draftH,
                                SummaryNovy = totalNovy.ToString(),
                                SummaryTotal = totalTotal.ToString(),
                                SummaryActive = totalActive.ToString(),
                                SummaryApproved = totalApproved.ToString(),
                                SummaryResolved = totalResolved.ToString(),
                                SummaryReviewed = totalReviewed.ToString()
                            };
                            ClashTestsTotal.Add(clashTestOnDrafts);
                        }
                    }
                }
            }

            #endregion

            #region see ClashTestsTotal
            /*
            string output = string.Empty;
            foreach (ClashTest ct in ClashTestsTotal)
            {
                output += ct.Name + "(" + ct.DraftLeftName + " : " + ct.DraftRightName + ")" + " -> " + ct.SummaryTotal + "\n";
            }

            Xceed.Wpf.Toolkit.MessageBox.Show("Вот такие тесты tottal\n" + output, "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);
            */
            #endregion


            foreach (SelectionSetNode selectionSetNodeVerticalSet in SelectionSetNodes)
            {
                foreach (SelectionSet selectionSetHName in selectionSetNodeVerticalSet.SelectionSets)
                {
                    DataLineView dataLineView = new DataLineView();

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
                        }
                        
                    }
                    ThisView.spDataVLine.Children.Add(dataLineView);
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
        private bool CanImportXMLClashTestsExecute(object p) => true;

        public ICommand ExcelExport { get; set; }
        private void OnExcelExportExecuted(object p)
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

                    #region creating Matrix worksheet 
                    List<ExcelWorksheet> worksheets = new List<ExcelWorksheet>();

                    excel.Workbook.Worksheets.Add("Matrix");
                    excel.Workbook.Worksheets.Add("Matrix-New");
                    excel.Workbook.Worksheets.Add("Matrix-Active");
                    excel.Workbook.Worksheets.Add("Matrix-Reviewed");
                    excel.Workbook.Worksheets.Add("Matrix-Approved");
                    excel.Workbook.Worksheets.Add("Matrix-Resolved");
                    excel.Workbook.Worksheets.Add("Matrix-Total");


                    worksheets.Add(excel.Workbook.Worksheets["Matrix"]);
                    worksheets.Add(excel.Workbook.Worksheets["Matrix-New"]);
                    worksheets.Add(excel.Workbook.Worksheets["Matrix-Active"]);
                    worksheets.Add(excel.Workbook.Worksheets["Matrix-Reviewed"]);
                    worksheets.Add(excel.Workbook.Worksheets["Matrix-Approved"]);
                    worksheets.Add(excel.Workbook.Worksheets["Matrix-Resolved"]);
                    worksheets.Add(excel.Workbook.Worksheets["Matrix-Total"]);

                    #region headers
                    // вертикальные заголовки слева
                    int rowCount = 3;
                    int colCount = 1;
                    foreach (SelectionSetNode ssn in SelectionSetNodes)
                    {
                        foreach (SelectionSet ss in ssn.SelectionSets)
                        {
                            foreach (ExcelWorksheet worksheet in worksheets)
                            {
                                worksheet.Cells[rowCount, colCount].Value = ssn.DraftName;
                                worksheet.Cells[rowCount, colCount + 1].Value = ss.Name;

                                worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                worksheet.Cells[rowCount, colCount + 1].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[rowCount, colCount + 1].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[rowCount, colCount + 1].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[rowCount, colCount + 1].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            }

                            rowCount += 1;
                        }
                    }

                    // горизонтальные заголовки сверху
                    rowCount = 1;
                    colCount = 3;
                    foreach(SelectionSetNode ssn in SelectionSetNodes)
                    {
                        foreach(SelectionSet ss in ssn.SelectionSets)
                        {
                            foreach (ExcelWorksheet worksheet in worksheets)
                            {
                                worksheet.Cells[rowCount, colCount].Value = ssn.DraftName;
                                worksheet.Cells[rowCount + 1, colCount].Value = ss.Name;
                                worksheet.Cells[rowCount + 1, colCount].Style.TextRotation = 90;

                                worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                worksheet.Cells[rowCount + 1, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[rowCount + 1, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[rowCount + 1, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheet.Cells[rowCount + 1, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            }
                                

                            colCount += 1;
                        }
                        
                    }

                    // объединение ячеек разделов проекта
                    int startcolumn = 3;
                    int startrow = 3;
                    int column = 0;
                    int row = 0;
                    foreach (SelectionSetNode ssn in SelectionSetNodes) // worksheet.Cells[FromRow, FromColumn, ToRow, ToColumn].Merge = true;
                    {
                        foreach(ExcelWorksheet worksheet in worksheets)
                        {
                            worksheet.Cells[1, startcolumn + column, 1, startcolumn + column + ssn.SelectionSets.Count - 1].Merge = true;
                            worksheet.Cells[1, startcolumn + column, 1, startcolumn + column + ssn.SelectionSets.Count - 1].Style.Font.Bold = true;
                            worksheet.Cells[startrow + row, 1, startrow + row + ssn.SelectionSets.Count - 1, 1].Merge = true;
                            worksheet.Cells[startrow + row, 1, startrow + row + ssn.SelectionSets.Count - 1, 1].Style.Font.Bold = true;

                            
                        }
                        for (int i2 = 0; i2 < ssn.SelectionSets.Count; i2++)
                        {
                            foreach (ExcelWorksheet worksheet in worksheets)
                            {
                                worksheet.Row(startrow + row + i2).CustomHeight = true;
                            }
                                
                        }

                        for (int i2 = 0; i2 < ssn.SelectionSets.Count; i2++)
                        {
                            foreach (ExcelWorksheet worksheet in worksheets)
                            {
                                worksheet.Column(startcolumn + column + i2).AutoFit();
                            }
                        }

                        column += ssn.SelectionSets.Count;
                        row += ssn.SelectionSets.Count;
                    }
                    foreach (ExcelWorksheet worksheet in worksheets)
                    {
                        worksheet.Cells[1, 1, 2, 2].Merge = true;
                        worksheet.Cells[1, 1, 2, 2].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[1, 1, 2, 2].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[1, 1, 2, 2].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[1, 1, 2, 2].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    }
                       

                    #endregion

                    rowCount = 3;
                    colCount = 3;
                    foreach (SelectionSet selSetV in SelectionSets)
                    {
                        foreach (SelectionSet selSetH in SelectionSets)
                        {
                            foreach (ClashTest clashTest in ClashTests)
                            {
                                if (clashTest.SelectionLeftName != null && clashTest.SelectionRightName != null && 
                                    clashTest.SelectionLeftName.Equals(selSetV.Name) && clashTest.SelectionRightName.Equals(selSetH.Name))
                                {
                                    //worksheet.Cells[rowCount, colCount].Value = clashTest.Name;

                                    bool resultTotal = int.TryParse(clashTest.SummaryTotal, out int total);
                                    bool resultResolved = int.TryParse(clashTest.SummaryResolved, out int resolved);
                                    bool resultApproved = int.TryParse(clashTest.SummaryApproved, out int approved);
                                    bool resultReviewed = int.TryParse(clashTest.SummaryReviewed, out int reviewed);
                                    bool resultActive = int.TryParse(clashTest.SummaryActive, out int active);
                                    bool resultNovy = int.TryParse(clashTest.SummaryNovy, out int novy);

                                    int summary = novy + active + reviewed + approved;

                                    foreach (ExcelWorksheet worksheet in worksheets)
                                    {
                                        if (worksheet.Name.Equals("Matrix"))
                                            worksheet.Cells[rowCount, colCount].Value = summary + "(" + clashTest.SummaryResolved + ")" + "\n" + clashTest.SummaryTotal;
                                        else if (worksheet.Name.Equals("Matrix-New"))
                                            worksheet.Cells[rowCount, colCount].Value = novy;
                                        else if (worksheet.Name.Equals("Matrix-Active"))
                                            worksheet.Cells[rowCount, colCount].Value = active;
                                        else if (worksheet.Name.Equals("Matrix-Reviewed"))
                                            worksheet.Cells[rowCount, colCount].Value = reviewed;
                                        else if (worksheet.Name.Equals("Matrix-Approved"))
                                            worksheet.Cells[rowCount, colCount].Value = approved;
                                        else if (worksheet.Name.Equals("Matrix-Resolved"))
                                            worksheet.Cells[rowCount, colCount].Value = resolved;
                                        else if (worksheet.Name.Equals("Matrix-Total"))
                                            worksheet.Cells[rowCount, colCount].Value = total;

                                        if (total == 0)
                                        {
                                            worksheet.Cells[rowCount, colCount].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                                        }
                                        else if (resolved == total)
                                        {
                                            worksheet.Cells[rowCount, colCount].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                                        }
                                        else if (resolved == 0)
                                        {
                                            worksheet.Cells[rowCount, colCount].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                                        }
                                        else
                                        {
                                            worksheet.Cells[rowCount, colCount].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                                        };
                                    }

                                }
                                foreach (ExcelWorksheet worksheet in worksheets)
                                {
                                    worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                }
                            }
                            colCount++;
                        }
                        rowCount++;
                        colCount = 3;
                    }

                    
                    // выравнивание содержимого ячеек результатов проверок
                    for (int i = 1; i <= ClashTests.Count + 2; i++)
                    {
                        foreach (ExcelWorksheet worksheet in worksheets)
                        {
                            worksheet.Row(i).Style.WrapText = true;
                            worksheet.Row(i).Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheet.Row(i).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        }   
                    }

                    // выравнивание наименований поисковых запросов в заголовках
                    foreach (ExcelWorksheet worksheet in worksheets)
                    {
                        worksheet.Row(2).Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Bottom;
                        worksheet.Row(2).Height = 150;
                        worksheet.Column(2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                        worksheet.Column(2).Width = 25;
                    }


                    // пояснение 
                    foreach (ExcelWorksheet worksheet in worksheets)
                    {
                        rowCount = SelectionSets.Count + 4;
                        colCount = SelectionSets.Count;
                        worksheet.Row(rowCount).Height = 30;
                        worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Merge = true;
                        
                        worksheet.Row(rowCount).Style.WrapText = true;
                        worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        //worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        //worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                        worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        if (worksheet.Name.Equals("Matrix"))
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Value = "В работе(Исправлено)\nВсего";
                        else if (worksheet.Name.Equals("Matrix-New"))
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Value = "Новые";
                        else if (worksheet.Name.Equals("Matrix-Active"))
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Value = "Активные";
                        else if (worksheet.Name.Equals("Matrix-Reviewed"))
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Value = "Проверенные";
                        else if (worksheet.Name.Equals("Matrix-Approved"))
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Value = "Подтвержденные";
                        else if (worksheet.Name.Equals("Matrix-Resolved"))
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Value = "Исправленные";
                        else if (worksheet.Name.Equals("Matrix-Total"))
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 2].Value = "Всего";
                    }
                        

                    #endregion

                    #region creating MatrixTotal worksheet

                    excel.Workbook.Worksheets.Add("Draft Matrix");
                    ExcelWorksheet worksheetDraft = excel.Workbook.Worksheets["Draft Matrix"];

                    #region headers

                    rowCount = 2;
                    colCount = 1;
                    foreach (SelectionSetNode ssn in SelectionSetNodes)
                    {
                        worksheetDraft.Cells[rowCount, colCount].Value = ssn.DraftName;

                        worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                        rowCount += 1;
                    }

                    rowCount = 1;
                    colCount = 2;
                    foreach (SelectionSetNode ssn in SelectionSetNodes)
                    {
                        worksheetDraft.Cells[rowCount, colCount].Value = ssn.DraftName;

                        worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                        colCount += 1;
                    }

                    #endregion

                    #region filling MatrixTotal with clash tests

                    rowCount = 2;
                    colCount = 2;
                    foreach(string draftV in Drafts)
                    {
                        foreach(string draftH in Drafts)
                        {
                            foreach(ClashTest clashTest in ClashTestsTotal)
                            {
                                if (clashTest.DraftLeftName.Equals(draftV) && clashTest.DraftRightName.Equals(draftH))
                                {
                                    //worksheet.Cells[rowCount, colCount].Value = clashTest.Name;

                                    bool resultTotal = int.TryParse(clashTest.SummaryTotal, out int total);
                                    bool resultResolved = int.TryParse(clashTest.SummaryResolved, out int resolved);
                                    bool resultApproved = int.TryParse(clashTest.SummaryApproved, out int approved);
                                    bool resultReviewed = int.TryParse(clashTest.SummaryReviewed, out int reviewed);
                                    bool resultActive = int.TryParse(clashTest.SummaryActive, out int active);
                                    bool resultNovy = int.TryParse(clashTest.SummaryNovy, out int novy);

                                    int summary = novy + active + reviewed + approved;

                                    worksheetDraft.Cells[rowCount, colCount].Value = summary + "(" + clashTest.SummaryResolved + ")" + "\n" + clashTest.SummaryTotal;

                                    if (total == 0)
                                    {
                                        worksheetDraft.Cells[rowCount, colCount].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                                    }
                                    else if (resolved == total)
                                    {
                                        worksheetDraft.Cells[rowCount, colCount].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                                    }
                                    else if (resolved == 0)
                                    {
                                        worksheetDraft.Cells[rowCount, colCount].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                                    }
                                    else
                                    {
                                        worksheetDraft.Cells[rowCount, colCount].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                                    };
                                }
                                
                            }
                            colCount++;
                        }
                        rowCount++;
                        colCount = 2;
                    }

                    for (int i = 1; i < Drafts.Count + 2; i++)
                    {
                        for (int j = 1; j < Drafts.Count + 2; j++)
                        {
                            worksheetDraft.Cells[i, j].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[i, j].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[i, j].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[i, j].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        }
                    }

                    // выравнивание содержимого ячеек результатов проверок
                    for (int i = 1; i <= Drafts.Count + 1; i++)
                    {
                        worksheetDraft.Row(i).Style.WrapText = true;
                        worksheetDraft.Row(i).Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheetDraft.Row(i).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheetDraft.Column(i).AutoFit();
                    }

                    worksheetDraft.Row(1).Style.Font.Bold = true;
                    worksheetDraft.Row(1).Height = 25;
                    worksheetDraft.Column(1).Style.Font.Bold = true;

                    // пояснение 
                    rowCount = Drafts.Count + 3;
                    colCount = Drafts.Count;
                    worksheetDraft.Row(rowCount).Height = 35;
                    worksheetDraft.Cells[rowCount, colCount, rowCount, colCount + 1].Merge = true;
                    worksheetDraft.Cells[rowCount, colCount, rowCount, colCount + 1].Value = "В работе(Исправлено)\nВсего";
                    worksheetDraft.Row(rowCount).Style.WrapText = true;
                    worksheetDraft.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheetDraft.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheetDraft.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheetDraft.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheetDraft.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheetDraft.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                    worksheetDraft.Cells[rowCount, colCount, rowCount, colCount + 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    worksheetDraft.Cells[rowCount, colCount, rowCount, colCount + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    #endregion


                    #endregion


                    FileInfo excelFile = new FileInfo(pathtoxls);
                    excel.SaveAs(excelFile);
                }
                catch (Exception ex)
                {
                    Xceed.Wpf.Toolkit.MessageBox.Show("Что-то произошло\n" + ex.ToString(), "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }

        }
        private bool CanExcelExportExecute(object p) => true;
    }
}
