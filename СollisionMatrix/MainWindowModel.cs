using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Xml;
using СollisionMatrix;
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

        private int _totalTotalCollisions;
        public int TotalTotalCollisions
        {
            get { return _totalTotalCollisions; }
            set
            {
                if (_totalTotalCollisions != value)
                {
                    _totalTotalCollisions = value;
                    OnPropertyChanged();
                }
            }
        }

        private int _totalNovyCollisions;
        public int TotalNovyCollisions
        {
            get { return _totalNovyCollisions; }
            set
            {
                if (_totalNovyCollisions != value)
                {
                    _totalNovyCollisions = value;
                    OnPropertyChanged();
                }
            }
        }
        
        private int _totalActiveCollisions;
        public int TotalActiveCollisions
        {
            get { return _totalActiveCollisions; }
            set
            {
                if (_totalActiveCollisions != value)
                {
                    _totalActiveCollisions = value;
                    OnPropertyChanged();
                }
            }
        }

        private int _totalReviewedCollisions;
        public int TotalReviewedCollisions
        {
            get { return _totalReviewedCollisions; }
            set
            {
                if (_totalReviewedCollisions != value)
                {
                    _totalReviewedCollisions = value;
                    OnPropertyChanged();
                }
            }
        }

        private int _totalApprovedCollisions;
        public int TotalApprovedCollisions
        {
            get { return _totalApprovedCollisions; }
            set
            {
                if (_totalApprovedCollisions != value)
                {
                    _totalApprovedCollisions = value;
                    OnPropertyChanged();
                }
            }
        }

        private int _totalResolvedCollisions;
        public int TotalResolvedCollisions
        {
            get { return _totalResolvedCollisions; }
            set
            {
                if (_totalResolvedCollisions != value)
                {
                    _totalResolvedCollisions = value;
                    OnPropertyChanged();
                }
            }
        }

        public System.Windows.Media.SolidColorBrush Color15 { get; set; } = new System.Windows.Media.SolidColorBrush()
        {
            Color = System.Windows.Media.Color.FromArgb(255, 255, 199, 199)
        };
        public System.Windows.Media.SolidColorBrush Color30 { get; set; } = new System.Windows.Media.SolidColorBrush()
        {
            Color = System.Windows.Media.Color.FromArgb(255, 248, 212, 255)
        };
        public System.Windows.Media.SolidColorBrush Color50 { get; set; } = new System.Windows.Media.SolidColorBrush()
        {
            Color = System.Windows.Media.Color.FromArgb(255, 204, 222, 255)
        };
        public System.Windows.Media.SolidColorBrush Color80 { get; set; } = new System.Windows.Media.SolidColorBrush()
        {
            Color = System.Windows.Media.Color.FromArgb(204, 204, 255, 220)
        };

        public ObservableCollection<SelectionSet> SelectionSets { get; set; }
        public ObservableCollection<SelectionSetNode> SelectionSetNodes { get; set; }
        public ObservableCollection<ClashTest> ClashTests { get; set; }
        public ObservableCollection<ClashTest> ClashTestsTotal { get; set; }
        public ObservableCollection<string> Drafts { get; set; }

        public DataCellViewModel[,] DataCellViewModels { get; set; }

        public MainWindowModel()
        {
            MatrixHeight = 800;
            MatrixWidth = 800;
            MatrixHeaderHeight = 150;
            MatrixHeaderWidth = 150;
            MatrixCellHorizontalDimension = 36;
            MatrixCellVerticalDimension = 24;

            TotalTotalCollisions = 0;
            TotalNovyCollisions = 0;
            TotalActiveCollisions = 0;
            TotalReviewedCollisions = 0;
            TotalApprovedCollisions = 0;
            TotalResolvedCollisions = 0;

            SelectionSets = new ObservableCollection<SelectionSet>();
            SelectionSetNodes = new ObservableCollection<SelectionSetNode>();
            ClashTests = new ObservableCollection<ClashTest>();
            ClashTestsTotal = new ObservableCollection<ClashTest>();
            Drafts = new ObservableCollection<string>();

            Commanda = new RelayCommand(OnCommandaExecuted, CanCommandaExecute); // example
            MatrixCreatingCommand = new RelayCommand(OnMatrixCreatingCommandExecuted, CanMatrixCreatingCommandExecute);
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

        public ICommand MatrixCreatingCommand { get; set; }
        private void OnMatrixCreatingCommandExecuted(object p)
        {
            MatrixCreating matrixCreatingWindow = new MatrixCreating();
            
            matrixCreatingWindow.Show();
        }
        private bool CanMatrixCreatingCommandExecute(object p) => true;

        public ICommand ImportXMLSelectionSets { get; set; }
        private void OnImportXMLSelectionSetsExecuted(object p)
        {
            bool success = true;

            int selectionsNumber = 0;

            SelectionSets = new ObservableCollection<SelectionSet>();
            SelectionSetNodes = new ObservableCollection<SelectionSetNode>();
            Drafts = new ObservableCollection<string>();
            ThisView.spSelectionSetsH.Children.Clear();
            ThisView.spSelectionSetsV.Children.Clear();
            ThisView.spDataVLine.Children.Clear();
            ThisView.borderUpLeft.Visibility = Visibility.Hidden;
            MatrixHeight = 800;
            MatrixWidth = 800;

            #region reading xml file
            string pathtoexe = Assembly.GetExecutingAssembly().Location;
            string dllname = Assembly.GetExecutingAssembly().GetName().Name.ToString();
            string folderpath = pathtoexe.Replace(dllname + ".exe", "");
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
                success = false;
                Xceed.Wpf.Toolkit.MessageBox.Show("Возможно, не выбран файл ... \n", "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            #endregion

            #region creating array SelectionSet from xml file

            try
            {
                XmlElement _exchange_ = xDoc.DocumentElement;
                foreach (XmlNode _selectionsets_ in _exchange_)
                {
                    foreach (XmlNode _child_ in _selectionsets_.ChildNodes)
                    {
                        if (_child_.Name.Equals("selectionset"))
                        {
                            SelectionSet selectionSet = new SelectionSet();
                            selectionSet.Name = _child_.Attributes.GetNamedItem("name").InnerText;

                            char lastchar = selectionSet.Name.Last();
                            do
                            {
                                if (lastchar.Equals(' ')) selectionSet.Name = selectionSet.Name.Remove(selectionSet.Name.Length - 1);
                                lastchar = selectionSet.Name.Last();
                            }
                            while (lastchar.Equals(' '));

                            var names = selectionSet.Name.Split('_');
                            if (names.Count() > 1)
                            {
                                selectionSet.DraftName = names.First();
                            }
                            else
                            {
                                selectionSet.DraftName = "?";
                            }

                            //int divUnderscore = 0;
                            //divUnderscore = selectionSet.Name.IndexOf("_");
                            //if (divUnderscore > 0)
                            //{
                            //    selectionSet.DraftName = selectionSet.Name.Substring(0, divUnderscore);
                            //}
                            //else
                            //{
                            //    selectionSet.DraftName = "?";
                            //}
                            SelectionSets.Add(selectionSet);
                        }
                        if (_child_.Name.Equals("viewfolder"))
                        {
                            foreach (XmlNode _child2_ in _child_.ChildNodes)
                            {
                                if (_child2_.Name.Equals("selectionset"))
                                {
                                    SelectionSet selectionSet = new SelectionSet();
                                    selectionSet.Name = _child2_.Attributes.GetNamedItem("name").InnerText;

                                    char lastchar = selectionSet.Name.Last();
                                    do
                                    {
                                        if (lastchar.Equals(' ')) selectionSet.Name = selectionSet.Name.Remove(selectionSet.Name.Length - 1);
                                        lastchar = selectionSet.Name.Last();
                                    }
                                    while (lastchar.Equals(' '));

                                    var names = selectionSet.Name.Split('_');
                                    if (names.Count() > 1)
                                    {
                                        selectionSet.DraftName = names.First();
                                    }
                                    else
                                    {
                                        selectionSet.DraftName = "?";
                                    }

                                    //int divUnderscore = 0;
                                    //divUnderscore = selectionSet.Name.IndexOf("_");
                                    //if (divUnderscore > 0)
                                    //{
                                    //    selectionSet.DraftName = selectionSet.Name.Substring(0, divUnderscore);
                                    //}
                                    //else
                                    //{
                                    //    selectionSet.DraftName = "?";
                                    //}
                                    SelectionSets.Add(selectionSet);
                                }  
                            }  
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                success = false;
                Xceed.Wpf.Toolkit.MessageBox.Show("Возможно, файл не содержит поисковых наборов ... \n", "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            /*
            string output = string.Empty;
            foreach (SelectionSet en in SelectionSets)
            {
                output += en.Name + "(" + en.DraftName + ")" + "\n";
            }

            Xceed.Wpf.Toolkit.MessageBox.Show("SelectionSets\n" + output, "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);
            */

            if (!success) return;

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
            ThisView.borderUpLeft.Visibility = Visibility.Visible;

        }
        private bool CanImportXMLSelectionSetsExecute(object p) => true;

        public ICommand ImportXMLClashTests { get; set; }
        private void OnImportXMLClashTestsExecuted(object p)
        {
            bool success = true;

            ClashTests = new ObservableCollection<ClashTest>();
            ClashTestsTotal = new ObservableCollection<ClashTest>();
            ThisView.spDataVLine.Children.Clear();


            #region reading xml file

            string pathtoexe = Assembly.GetExecutingAssembly().Location;
            string dllname = Assembly.GetExecutingAssembly().GetName().Name.ToString();
            string folderpath = pathtoexe.Replace(dllname + ".exe", "");
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
                success = false;
                Xceed.Wpf.Toolkit.MessageBox.Show("Возможно, не выбран файл ... \n", "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            if (!success) return;

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

                                                var names = clashTest.SelectionLeftName.Split('/');
                                                clashTest.SelectionLeftName = names.LastOrDefault();

                                                var names2 = clashTest.SelectionLeftName.Split('_');
                                                if (names2.Count() > 1)
                                                {
                                                    clashTest.DraftLeftName = names2.First();
                                                }
                                                else
                                                {
                                                    clashTest.DraftLeftName = "?";
                                                }

                                                //int divUnderscore = 0;
                                                //divUnderscore = clashTest.SelectionLeftName.IndexOf("_");
                                                //if (divUnderscore != 0 && divUnderscore > 0)
                                                //{
                                                //    clashTest.DraftLeftName = clashTest.SelectionLeftName.Substring(0, divUnderscore);
                                                //}
                                                //else
                                                //{
                                                //    clashTest.DraftLeftName = "?";
                                                //}
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

                                                var names = clashTest.SelectionRightName.Split('/');
                                                clashTest.SelectionRightName = names.LastOrDefault();

                                                var names2 = clashTest.SelectionRightName.Split('_');
                                                if (names2.Count() > 1)
                                                {
                                                    clashTest.DraftRightName = names2.First();
                                                }
                                                else
                                                {
                                                    clashTest.DraftRightName = "?";
                                                }

                                                //int divUnderscore = 0;
                                                //divUnderscore = clashTest.SelectionRightName.IndexOf("_");
                                                //if (divUnderscore > 0)
                                                //{
                                                //    clashTest.DraftRightName = clashTest.SelectionRightName.Substring(0, divUnderscore);
                                                //}
                                                //else
                                                //{
                                                //    clashTest.DraftRightName = "?";
                                                //}
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
                success = false;
                Xceed.Wpf.Toolkit.MessageBox.Show("Возможно, файл не содержит результатов проверки... \n" + ex.ToString(), "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            if (!success) return;

            #endregion

            #region see ClashTests
            /*
            string output = string.Empty;
            foreach (ClashTest ct in ClashTests)
            {
                output += ct.Name + "\n" + "  (" + ct.SelectionLeftName + " / " + ct.SelectionRightName + ")" + "\n" + " -> " + ct.SummaryTotal + "\n";
            }

            Xceed.Wpf.Toolkit.MessageBox.Show("Вот такие тесты\n" + output, "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);
            */
            #endregion

            #region creating array ClashTestsTotal from ClashTests

            foreach (ClashTest ct in ClashTests)
            {
                bool resTotal = int.TryParse(ct.SummaryTotal, out int ct_total);
                bool resResolved = int.TryParse(ct.SummaryResolved, out int ct_resolved);
                bool resApproved = int.TryParse(ct.SummaryApproved, out int ct_approved);
                bool resReviewed = int.TryParse(ct.SummaryReviewed, out int ct_reviewed);
                bool resActive = int.TryParse(ct.SummaryActive, out int ct_active);
                bool resNovy = int.TryParse(ct.SummaryNovy, out int ct_novy);

                TotalTotalCollisions += ct_total;
                TotalResolvedCollisions += ct_resolved;
                TotalApprovedCollisions += ct_approved;
                TotalReviewedCollisions += ct_reviewed;
                TotalActiveCollisions += ct_active;
                TotalNovyCollisions += ct_novy;


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

            //string filename = @"C:\Users\" + Environment.UserName + @"\Downloads\reportcollisions.txt";

            int row_models = 0;
            int col_models = 0;
            DataCellViewModels = new DataCellViewModel[SelectionSets.Count, SelectionSets.Count];

            foreach (SelectionSetNode selectionSetNodeVerticalSet in SelectionSetNodes)
            {
                foreach (SelectionSet selectionSetVName in selectionSetNodeVerticalSet.SelectionSets)
                {
                    DataLineView dataLineView = new DataLineView();

                    foreach (SelectionSetNode selectionSetNodeHorizontalSet in SelectionSetNodes)
                    {
                        foreach (SelectionSet selectionSetHName in selectionSetNodeHorizontalSet.SelectionSets)
                        {
                            DataCellView dataCellView = new DataCellView();
                            dataCellView.Width = MatrixCellHorizontalDimension;
                            dataCellView.Height = MatrixCellVerticalDimension;
                            DataCellViewModel dataCellViewModel = (DataCellViewModel)dataCellView.DataContext;
                            dataCellViewModel.SelectionLeftName = selectionSetVName.Name;
                            dataCellViewModel.SelectionRightName = selectionSetHName.Name;
                            dataCellViewModel.CollisionsTotalNumber = "";
                            dataCellViewModel.CollisionsActiveNumber = "";
                            dataCellViewModel.CollisionsResolvedNumber = "";
                            dataCellViewModel.CollisionsReviewedNumber = "";
                            
                            foreach (ClashTest clashTest in ClashTests)
                            {
                                //bool result1 = dataCellViewModel.SelectionLeftName.Equals(clashTest.SelectionLeftName);
                                //bool result2 = dataCellViewModel.SelectionRightName.Equals(clashTest.SelectionRightName);

                                //WriteToFile(filename, "dataCellView.SelectionLeftName: " + dataCellViewModel.SelectionLeftName + ", clashTest.SelectionLeftName: " + clashTest.SelectionLeftName + result1);
                                //WriteToFile(filename, "dataCellView.SelectionRightName: " + dataCellViewModel.SelectionRightName + ", clashTest.SelectionRightName: " + clashTest.SelectionRightName + result2);

                                if (dataCellViewModel.SelectionLeftName.Equals(clashTest.SelectionLeftName) && dataCellViewModel.SelectionRightName.Equals(clashTest.SelectionRightName))
                                {
                                    bool resultNovy = int.TryParse(clashTest.SummaryNovy, out int novy);
                                    bool resultActive = int.TryParse(clashTest.SummaryActive, out int active);
                                    bool resultReviewed = int.TryParse(clashTest.SummaryReviewed, out int reviewed);
                                    bool resultResolved = int.TryParse(clashTest.SummaryResolved, out int resolved);
                                    bool resultApproved = int.TryParse(clashTest.SummaryApproved, out int approved);
                                    bool resultTotal = int.TryParse(clashTest.SummaryTotal, out int total);

                                    int summary = novy + active + reviewed; // + approved;

                                    dataCellViewModel.CollisionsSumma = summary.ToString();
                                    dataCellViewModel.CollisionsTotalNumber = clashTest.SummaryTotal;
                                    dataCellViewModel.CollisionsActiveNumber = clashTest.SummaryActive;
                                    dataCellViewModel.CollisionsReviewedNumber = clashTest.SummaryReviewed;
                                    dataCellViewModel.CollisionsResolvedNumber = clashTest.SummaryResolved;

                                    dataCellViewModel.ClashTest = clashTest;

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

                                        dataCellView.ToolTip = $"{selectionSetVName.Name} / {selectionSetHName.Name}\nАктивно: {clashTest.SummaryActive}\nИсправлено: {clashTest.SummaryResolved}\nВсего: {clashTest.SummaryTotal}";
                                    }
                                }
                            }
                            
                            dataLineView.spLine.Children.Add(dataCellView);
                            DataCellViewModels[row_models, col_models] = (DataCellViewModel)dataCellView.DataContext;
                            col_models += 1;
                        }
                        
                    }
                    ThisView.spDataVLine.Children.Add(dataLineView);
                    col_models = 0;
                    row_models += 1;
                }
            }

            /*
            string output = string.Empty;
            for (int r = 0; r < SelectionSets.Count; r++)
            {
                for (int c = 0; c < SelectionSets.Count; c++)
                {
                    output += DataCellViewModels[r, c].CollisionsTotalNumber + ",";
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
                    Xceed.Wpf.Toolkit.MessageBox.Show("Возможно, нет доступа к папке\n", "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
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
                    for (int r = 0; r < SelectionSets.Count; r++)
                    {
                        rowCount = 3 + r;

                        for (int c = 0; c < SelectionSets.Count; c++)
                        {
                            colCount = 3 + c;

                            if (DataCellViewModels[r, c].ClashTest != null)
                            {
                                ClashTest clashTest = DataCellViewModels[r, c].ClashTest;

                                bool resultNovy = int.TryParse(clashTest.SummaryNovy, out int novy);
                                bool resultActive = int.TryParse(clashTest.SummaryActive, out int active);
                                bool resultReviewed = int.TryParse(clashTest.SummaryReviewed, out int reviewed);
                                bool resultResolved = int.TryParse(clashTest.SummaryResolved, out int resolved);
                                bool resultApproved = int.TryParse(clashTest.SummaryApproved, out int approved);
                                bool resultTotal = int.TryParse(clashTest.SummaryTotal, out int total);

                                int summary = novy + active + reviewed; // + approved;

                                foreach (ExcelWorksheet worksheet in worksheets)
                                {
                                    if (worksheet.Name.Equals("Matrix"))
                                        worksheet.Cells[rowCount, colCount].Value = summary + "(" + resolved + ")" + "\n" + total;
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
                    }

                    #region other way
                    /*
                    rowCount = 3;
                    colCount = 3;
                    foreach (SelectionSetNode ssnV in SelectionSetNodes)
                    {
                        foreach (SelectionSet selSetV in ssnV.SelectionSets)
                        {
                            foreach(SelectionSetNode ssnH in SelectionSetNodes)
                            {
                                foreach (SelectionSet selSetH in ssnH.SelectionSets)
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
                                                    //worksheet.Cells[rowCount, colCount].Value = clashTest.SelectionLeftName + "\n" + clashTest.SelectionRightName;
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
                            }
                            rowCount++;
                            colCount = 3;
                        }

                    }
                    */
                    #endregion

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
                        int colstart = 6;
                        int coldelta = 4;
                        int colmerge = 3;
                        rowCount = SelectionSets.Count + 4;

                        if (worksheet.Name.Equals("Matrix"))
                        {
                            colCount = colstart;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            worksheet.Cells[rowCount, colCount].Value = "0(0)\n0";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии отсуствуют";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            worksheet.Cells[rowCount, colCount].Value = "0(6)\n6";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Все коллизии устранены";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 2;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                            worksheet.Cells[rowCount, colCount].Value = "6(0)\n6";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Устранение коллизий не ведется";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 3;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                            worksheet.Cells[rowCount, colCount].Value = "12(4)\n16";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии устранены частично";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 4 + 1;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Merge = true;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Value = $"{TotalNovyCollisions + TotalActiveCollisions + TotalReviewedCollisions}({TotalResolvedCollisions})\n{TotalTotalCollisions}";
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Merge = true;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Value = "В работе(Устранено)\nВсего";
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                        }
                        else if (worksheet.Name.Equals("Matrix-New"))
                        {
                            colCount = colstart;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            worksheet.Cells[rowCount, colCount].Value = 0;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии отсуствуют";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            worksheet.Cells[rowCount, colCount].Value = 0;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Все коллизии устранены";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 2;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                            worksheet.Cells[rowCount, colCount].Value = 6;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Устранение коллизий не ведется";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 3;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                            worksheet.Cells[rowCount, colCount].Value = 6;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии устранены частично";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 4 + 1;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Merge = true;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Value = TotalNovyCollisions;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Merge = true;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Value = "Всего новых созданных коллизий";
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                        }
                        else if (worksheet.Name.Equals("Matrix-Active"))
                        {
                            colCount = colstart;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            worksheet.Cells[rowCount, colCount].Value = 0;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии отсуствуют";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            worksheet.Cells[rowCount, colCount].Value = 0;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Все коллизии устранены";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 2;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                            worksheet.Cells[rowCount, colCount].Value = 6;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Устранение коллизий не ведется";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 3;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                            worksheet.Cells[rowCount, colCount].Value = 6;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии устранены частично";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 4 + 1;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Merge = true;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Value = TotalActiveCollisions;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Merge = true;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Value = "Всего активных коллизий";
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;
                        }
                        else if (worksheet.Name.Equals("Matrix-Reviewed"))
                        {
                            colCount = colstart;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            worksheet.Cells[rowCount, colCount].Value = 0;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии отсуствуют";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            worksheet.Cells[rowCount, colCount].Value = 0;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Все коллизии устранены";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 2;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                            worksheet.Cells[rowCount, colCount].Value = 6;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Устранение коллизий не ведется";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 3;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                            worksheet.Cells[rowCount, colCount].Value = 6;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии устранены частично";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 4 + 1;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Merge = true;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Value = TotalReviewedCollisions;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Merge = true;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Value = "Всего проверенных коллизий";
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;
                        }
                        else if (worksheet.Name.Equals("Matrix-Approved"))
                        {
                            colCount = colstart;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            worksheet.Cells[rowCount, colCount].Value = 0;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии отсуствуют";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            worksheet.Cells[rowCount, colCount].Value = 0;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Все коллизии устранены";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 2;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                            worksheet.Cells[rowCount, colCount].Value = 6;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Устранение коллизий не ведется";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 3;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                            worksheet.Cells[rowCount, colCount].Value = 6;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии устранены частично";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 4 + 1;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Merge = true;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Value = TotalApprovedCollisions;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Merge = true;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Value = "Всего подтвержденных коллизий (не коллизии)";
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;
                        }
                        else if (worksheet.Name.Equals("Matrix-Resolved"))
                        {
                            colCount = colstart;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            worksheet.Cells[rowCount, colCount].Value = 0;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии отсуствуют";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            worksheet.Cells[rowCount, colCount].Value = 0;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Все коллизии устранены";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 2;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                            worksheet.Cells[rowCount, colCount].Value = 6;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Устранение коллизий не ведется";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 3;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                            worksheet.Cells[rowCount, colCount].Value = 6;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии устранены частично";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 4 + 1;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Merge = true;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Value = TotalResolvedCollisions;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Merge = true;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Value = "Всего исправленных коллизий";
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;
                        }
                        else if (worksheet.Name.Equals("Matrix-Total"))
                        {
                            colCount = colstart;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            worksheet.Cells[rowCount, colCount].Value = 0;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии отсуствуют";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            worksheet.Cells[rowCount, colCount].Value = 0;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Все коллизии устранены";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 2;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                            worksheet.Cells[rowCount, colCount].Value = 6;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Устранение коллизий не ведется";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 3;
                            worksheet.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                            worksheet.Cells[rowCount, colCount].Value = 6;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии устранены частично";
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;

                            colCount = colstart + coldelta * 4 + 1;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Merge = true;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheet.Cells[rowCount, colCount, rowCount, colCount + 1].Value = TotalTotalCollisions;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Merge = true;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Value = "Всего коллизий";
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheet.Cells[rowCount, colCount + 2, rowCount, colCount + colmerge + 1].Style.Indent = 1;
                            worksheet.Row(rowCount).Height = 35;
                        }
                    }
                        

                    #endregion

                    #region creating MatrixTotal worksheet

                    excel.Workbook.Worksheets.Add("DraftMatrix");
                    excel.Workbook.Worksheets.Add("DraftMatrix-New");
                    excel.Workbook.Worksheets.Add("DraftMatrix-Active");
                    excel.Workbook.Worksheets.Add("DraftMatrix-Reviewed");
                    excel.Workbook.Worksheets.Add("DraftMatrix-Approved");
                    excel.Workbook.Worksheets.Add("DraftMatrix-Resolved");
                    excel.Workbook.Worksheets.Add("DraftMatrix-Total");

                    List<ExcelWorksheet> worksheetsDraft = new List<ExcelWorksheet>();

                    worksheetsDraft.Add(excel.Workbook.Worksheets["DraftMatrix"]);
                    worksheetsDraft.Add(excel.Workbook.Worksheets["DraftMatrix-New"]);
                    worksheetsDraft.Add(excel.Workbook.Worksheets["DraftMatrix-Active"]);
                    worksheetsDraft.Add(excel.Workbook.Worksheets["DraftMatrix-Reviewed"]);
                    worksheetsDraft.Add(excel.Workbook.Worksheets["DraftMatrix-Approved"]);
                    worksheetsDraft.Add(excel.Workbook.Worksheets["DraftMatrix-Resolved"]);
                    worksheetsDraft.Add(excel.Workbook.Worksheets["DraftMatrix-Total"]);

                    #region headers

                    rowCount = 2;
                    colCount = 1;
                    foreach (SelectionSetNode ssn in SelectionSetNodes)
                    {
                        foreach (ExcelWorksheet worksheetDraft in worksheetsDraft) 
                        {
                            worksheetDraft.Cells[rowCount, colCount].Value = ssn.DraftName;

                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        }

                        rowCount += 1;
                    }

                    rowCount = 1;
                    colCount = 2;
                    foreach (SelectionSetNode ssn in SelectionSetNodes)
                    {
                        foreach (ExcelWorksheet worksheetDraft in worksheetsDraft)
                        {
                            worksheetDraft.Cells[rowCount, colCount].Value = ssn.DraftName;

                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        }

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
                                    bool resultNovy = int.TryParse(clashTest.SummaryNovy, out int novy);
                                    bool resultActive = int.TryParse(clashTest.SummaryActive, out int active);
                                    bool resultReviewed = int.TryParse(clashTest.SummaryReviewed, out int reviewed);
                                    bool resultApproved = int.TryParse(clashTest.SummaryApproved, out int approved);
                                    bool resultResolved = int.TryParse(clashTest.SummaryResolved, out int resolved);
                                    bool resultTotal = int.TryParse(clashTest.SummaryTotal, out int total);

                                    int summary = novy + active + reviewed; // + approved;

                                    foreach (ExcelWorksheet worksheetDraft in worksheetsDraft)
                                    {
                                        if (worksheetDraft.Name.Equals("DraftMatrix"))
                                            worksheetDraft.Cells[rowCount, colCount].Value = summary + "(" + resolved + ")" + "\n" + total;
                                        else if (worksheetDraft.Name.Equals("DraftMatrix-New"))
                                            worksheetDraft.Cells[rowCount, colCount].Value = novy;
                                        else if (worksheetDraft.Name.Equals("DraftMatrix-Active"))
                                            worksheetDraft.Cells[rowCount, colCount].Value = active;
                                        else if (worksheetDraft.Name.Equals("DraftMatrix-Reviewed"))
                                            worksheetDraft.Cells[rowCount, colCount].Value = reviewed;
                                        else if (worksheetDraft.Name.Equals("DraftMatrix-Approved"))
                                            worksheetDraft.Cells[rowCount, colCount].Value = approved;
                                        else if (worksheetDraft.Name.Equals("DraftMatrix-Resolved"))
                                            worksheetDraft.Cells[rowCount, colCount].Value = resolved;
                                        else if (worksheetDraft.Name.Equals("DraftMatrix-Total"))
                                            worksheetDraft.Cells[rowCount, colCount].Value = total;
                                        
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
                            foreach (ExcelWorksheet worksheetDraft in worksheetsDraft)
                            {
                                worksheetDraft.Cells[i, j].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheetDraft.Cells[i, j].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheetDraft.Cells[i, j].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                worksheetDraft.Cells[i, j].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                if (i > 1 && j > 1) worksheetDraft.Cells[i, j].Style.Font.Size = 10;
                            }
                        }
                    }

                    // выравнивание содержимого ячеек результатов проверок
                    for (int i = 1; i <= Drafts.Count + 1; i++)
                    {
                        foreach (ExcelWorksheet worksheetDraft in worksheetsDraft)
                        {
                            worksheetDraft.Row(i).Style.WrapText = true;
                            worksheetDraft.Row(i).Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Row(i).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Column(i).AutoFit();
                            if (i > 1) worksheetDraft.Column(i).Width = 10;
                        }
                    }
                    foreach (ExcelWorksheet worksheetDraft in worksheetsDraft)
                    {
                        worksheetDraft.Row(1).Style.Font.Bold = true;
                        worksheetDraft.Row(1).Height = 25;
                        worksheetDraft.Column(1).Style.Font.Bold = true;
                    }
                        

                    // пояснение 
                    
                    foreach (ExcelWorksheet worksheetDraft in worksheetsDraft)
                    {
                        
                        int rowstart = Drafts.Count + 4;
                        int rowdelta = 2;
                        int colmerge = 3;

                        colCount = 2;
                        
                        if (worksheetDraft.Name.Equals("DraftMatrix"))
                        {
                            rowCount = rowstart;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            worksheetDraft.Cells[rowCount, colCount].Value = "0(0)\n0";
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии отсуствуют";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            worksheetDraft.Cells[rowCount, colCount].Value = "0(6)\n6";
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Все коллизии устранены";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 2;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                            worksheetDraft.Cells[rowCount, colCount].Value = "0(6)\n6";
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Устранение коллизий не ведется";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 3;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                            worksheetDraft.Cells[rowCount, colCount].Value = "12(4)\n16";
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии устранены частично";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 4;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Value = $"{TotalNovyCollisions + TotalActiveCollisions + TotalReviewedCollisions}({TotalResolvedCollisions})\n{TotalTotalCollisions}";
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "В работе(Устранено)\nВсего";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;
                        }
                        else if (worksheetDraft.Name.Equals("DraftMatrix-New"))
                        {
                            rowCount = rowstart;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            worksheetDraft.Cells[rowCount, colCount].Value = 0;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии отсуствуют";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            worksheetDraft.Cells[rowCount, colCount].Value = 0;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Все коллизии устранены";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 2;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                            worksheetDraft.Cells[rowCount, colCount].Value = 6;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Устранение коллизий не ведется";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 3;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                            worksheetDraft.Cells[rowCount, colCount].Value = 6;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии устранены частично";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 4;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Value = TotalNovyCollisions;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Всего новых созданных коллизий";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;
                        }
                        else if (worksheetDraft.Name.Equals("DraftMatrix-Active"))
                        {
                            rowCount = rowstart;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            worksheetDraft.Cells[rowCount, colCount].Value = 0;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии отсуствуют";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            worksheetDraft.Cells[rowCount, colCount].Value = 0;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Все коллизии устранены";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 2;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                            worksheetDraft.Cells[rowCount, colCount].Value = 6;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Устранение коллизий не ведется";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 3;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                            worksheetDraft.Cells[rowCount, colCount].Value = 6;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии устранены частично";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 4;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Value = TotalActiveCollisions;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Всего активных коллизий";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;
                        }
                        else if (worksheetDraft.Name.Equals("DraftMatrix-Reviewed"))
                        {
                            rowCount = rowstart;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            worksheetDraft.Cells[rowCount, colCount].Value = 0;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии отсуствуют";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            worksheetDraft.Cells[rowCount, colCount].Value = 0;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Все коллизии устранены";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 2;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                            worksheetDraft.Cells[rowCount, colCount].Value = 6;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Устранение коллизий не ведется";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 3;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                            worksheetDraft.Cells[rowCount, colCount].Value = 6;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии устранены частично";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 4;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Value = TotalReviewedCollisions;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Всего проверенных коллизий";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;
                        }
                        else if (worksheetDraft.Name.Equals("DraftMatrix-Approved"))
                        {
                            rowCount = rowstart;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            worksheetDraft.Cells[rowCount, colCount].Value = 0;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии отсуствуют";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            worksheetDraft.Cells[rowCount, colCount].Value = 0;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Все коллизии устранены";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 2;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                            worksheetDraft.Cells[rowCount, colCount].Value = 6;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Устранение коллизий не ведется";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 3;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                            worksheetDraft.Cells[rowCount, colCount].Value = 6;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии устранены частично";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 4;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Value = TotalApprovedCollisions;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Всего подтвержденных коллизий (не коллизии)";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;
                        }
                        else if (worksheetDraft.Name.Equals("DraftMatrix-Resolved"))
                        {
                            rowCount = rowstart;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            worksheetDraft.Cells[rowCount, colCount].Value = 0;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии отсуствуют";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            worksheetDraft.Cells[rowCount, colCount].Value = 0;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Все коллизии устранены";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 2;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                            worksheetDraft.Cells[rowCount, colCount].Value = 6;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Устранение коллизий не ведется";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 3;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                            worksheetDraft.Cells[rowCount, colCount].Value = 6;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии устранены частично";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 4;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Value = TotalResolvedCollisions;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Всего исправленных коллизий";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;
                        }
                        else if (worksheetDraft.Name.Equals("DraftMatrix-Total"))
                        {
                            rowCount = rowstart;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            worksheetDraft.Cells[rowCount, colCount].Value = 0;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии отсуствуют";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            worksheetDraft.Cells[rowCount, colCount].Value = 0;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Все коллизии устранены";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 2;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(255, 254, 191));
                            worksheetDraft.Cells[rowCount, colCount].Value = 6;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Устранение коллизий не ведется";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 3;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(191, 233, 255));
                            worksheetDraft.Cells[rowCount, colCount].Value = 6;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Коллизии устранены частично";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;

                            rowCount = rowstart + rowdelta * 4;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            worksheetDraft.Cells[rowCount, colCount].Value = TotalTotalCollisions;
                            worksheetDraft.Cells[rowCount, colCount].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount].Style.Font.Size = 10;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Merge = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Value = "Всего коллизий";
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.WrapText = true;
                            worksheetDraft.Cells[rowCount, colCount + 1, rowCount, colCount + colmerge].Style.Indent = 1;
                            worksheetDraft.Row(rowCount).Height = 30;
                        }
                    }

                    #endregion


                    #endregion


                    FileInfo excelFile = new FileInfo(pathtoxls);
                    excel.SaveAs(excelFile);

                    var choose = Xceed.Wpf.Toolkit.MessageBox.Show("Сохранен файл с отчетом:\n" + pathtoxls, "Вопросик", MessageBoxButton.YesNo, MessageBoxImage.Question);

                    if (choose == MessageBoxResult.Yes)
                    {
                        System.Diagnostics.Process.Start(pathtoxls);
                    }

                }
                catch (Exception ex)
                {
                    Xceed.Wpf.Toolkit.MessageBox.Show("Что-то произошло\n" + ex.ToString(), "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }

        }
        private bool CanExcelExportExecute(object p) => true;


        private static void WriteToFile(string fileName, string txt)
        {
            Encoding u16LE = Encoding.Unicode; // UTF-16 little endian

            if (!File.Exists(fileName))
            {
                try
                {
                    using (FileStream fs = File.Create(fileName))
                    {
                        byte[] byteString = new UnicodeEncoding().GetBytes("");
                        fs.Write(byteString, 0, byteString.Length);
                    }
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show(ex.ToString());
                }
            }
            using (StreamWriter writer = new StreamWriter(fileName, true))
            {
                writer.WriteLine(txt);
            }

        }

    }
}
