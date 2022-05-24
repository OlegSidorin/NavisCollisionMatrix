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
using СollisionMatrix.Models;
using System.Windows.Controls;

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

        List<string> SelectionsetNames { get; set; }
        public List<Selectionset> Selectionsets { get; set; }
        public List<Clashtest> Clashtests { get; set; }
        public Batchtest Batchtest { get; set; }
        public ObservableCollection<UserControl> LineUserControls { get; set; }

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

        public ObservableCollection<string> Drafts { get; set; }

        public DataCellViewModel[,] DataCellViewModels { get; set; }

        public MainWindowModel()
        {
            LineUserControls = new ObservableCollection<UserControl>();

            Batchtest = new Batchtest()
            {
                Tag_name = "Report",
                Tag_internal_name = "Report",
                Tag_units = "m"
            };

            SelectionsetNames = new List<string>();

            TotalTotalCollisions = 0;
            TotalNovyCollisions = 0;
            TotalActiveCollisions = 0;
            TotalReviewedCollisions = 0;
            TotalApprovedCollisions = 0;
            TotalResolvedCollisions = 0;

            Drafts = new ObservableCollection<string>();

            Commanda = new RelayCommand(OnCommandaExecuted, CanCommandaExecute); // example
            MatrixCreatingCommand = new RelayCommand(OnMatrixCreatingCommandExecuted, CanMatrixCreatingCommandExecute);
            ImportXMLClashtests = new RelayCommand(OnImportXMLClashtestsExecuted, CanImportXMLClashtestsExecute);
            //ExcelExport = new RelayCommand(OnExcelExportExecuted, CanExcelExportExecute);

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

        public ICommand ImportXMLClashtests { get; set; }
        private void OnImportXMLClashtestsExecuted(object p)
        {
            bool success = true;

            LineUserControls.Clear();
            SelectionsetNames.Clear();
            Selectionsets = new List<Selectionset>();
            Clashtests = new List<Clashtest>();

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

            #region creating array Clashtests from xml file
            
            try
            {
                XmlElement exchange_element = xDoc.DocumentElement; // <exchange xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd" units="ft" filename="" filepath="">

                foreach (XmlNode batchtest_node in exchange_element) // <batchtest name="Без имени" internal_name="Без имени" units="ft">
                {
                    if (batchtest_node.Name == "batchtest")
                    {
                        Batchtest.Tag_name = batchtest_node.Attributes.GetNamedItem("name").InnerText;
                        Batchtest.Tag_internal_name = batchtest_node.Attributes.GetNamedItem("internal_name").InnerText;
                        Batchtest.Tag_units = batchtest_node.Attributes.GetNamedItem("units").InnerText;
                        foreach (XmlNode node_in_batchtest_node in batchtest_node.ChildNodes) // <clashtests>
                        {
                            if (node_in_batchtest_node.Name == "clashtests")
                            {
                                foreach (XmlNode clashtest_node in node_in_batchtest_node.ChildNodes) // <clashtest name="АР_Вертикальные конструкции - АР_Вертикальные конструкции" test_type="duplicate" status="ok" tolerance="0.000" merge_composites="0">
                                {
                                    if (clashtest_node.Name == "clashtest")
                                    {
                                        Clashtest clashtest = new Clashtest();
                                        clashtest.Tag_name = clashtest_node.Attributes.GetNamedItem("name").InnerText;
                                        clashtest.Tag_test_type = clashtest_node.Attributes.GetNamedItem("test_type").InnerText;
                                        clashtest.Tag_status = clashtest_node.Attributes.GetNamedItem("status").InnerText;
                                        clashtest.Tag_merge_composites = clashtest_node.Attributes.GetNamedItem("merge_composites").InnerText;
                                        clashtest.Tag_tolerance = clashtest_node.Attributes.GetNamedItem("tolerance").InnerText;
                                        bool to_double = double.TryParse(clashtest.Tag_tolerance.Replace(".", ","), out double td);
                                        if (Batchtest.Tag_units == "ft")
                                        {
                                            td *= 304.8;
                                            clashtest.Tag_tolerance_in_mm = ((int)td).ToString();
                                        } 
                                        else if (Batchtest.Tag_units == "m")
                                        {
                                            td *= 1000;
                                            clashtest.Tag_tolerance_in_mm = ((int)td).ToString();
                                        }
                                        foreach (XmlNode node_in_clashtest_node in clashtest_node.ChildNodes)
                                        {
                                            if (node_in_clashtest_node.Name == "linkage")
                                            {
                                                clashtest.Linkage = new Linkage();
                                                clashtest.Linkage.Tag_mode = node_in_clashtest_node.Attributes.GetNamedItem("mode").InnerText;
                                            }
                                            else if (node_in_clashtest_node.Name == "left")
                                            {
                                                clashtest.Left = new Left();
                                                foreach (XmlNode clashselection_node in node_in_clashtest_node.ChildNodes)
                                                {
                                                    if (clashselection_node.Name == "clashselection")
                                                    {
                                                        clashtest.Left.Clashselection = new Clashselection();
                                                        clashtest.Left.Clashselection.Tag_primtypes = clashselection_node.Attributes.GetNamedItem("primtypes").InnerText;
                                                        clashtest.Left.Clashselection.Tag_selfintersect = clashselection_node.Attributes.GetNamedItem("selfintersect").InnerText;

                                                        foreach (XmlNode locator_node in clashselection_node.ChildNodes)
                                                        {
                                                            if (locator_node.Name == "locator")
                                                            {
                                                                clashtest.Left.Clashselection.Locator = new Locator();
                                                                clashtest.Left.Clashselection.Locator.Tag_inner_text = locator_node.InnerText;
                                                                //clashtest.Left.Clashselection.Locator.Tag_inner_text_selection_name = locator_node.InnerText.Replace("lcop_selection_set_tree/", "");

                                                                var folders_names = locator_node.InnerText.Split('/');
                                                                clashtest.Left.Clashselection.Locator.Tag_inner_text_selection_name = folders_names.LastOrDefault();
                                                                clashtest.Left.Clashselection.Locator.Tag_inner_text_folders = new List<string>();
                                                                foreach (string folder_name in folders_names)
                                                                {
                                                                    clashtest.Left.Clashselection.Locator.Tag_inner_text_folders.Add(folder_name);
                                                                }

                                                                var draft_names = clashtest.Left.Clashselection.Locator.Tag_inner_text_selection_name.Split('_');
                                                                if (draft_names.Count() > 1)
                                                                {
                                                                    clashtest.Left.Clashselection.Locator.Tag_inner_text_draft_name = draft_names.First();
                                                                }
                                                                else
                                                                {
                                                                    clashtest.Left.Clashselection.Locator.Tag_inner_text_draft_name = "?";
                                                                }
                                                                SelectionsetNames.Add(clashtest.Left.Clashselection.Locator.Tag_inner_text_selection_name);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            else if (node_in_clashtest_node.Name == "right")
                                            {
                                                clashtest.Right = new Right();
                                                foreach (XmlNode clashselection_node in node_in_clashtest_node.ChildNodes)
                                                {
                                                    if (clashselection_node.Name == "clashselection")
                                                    {
                                                        clashtest.Right.Clashselection = new Clashselection();
                                                        clashtest.Right.Clashselection.Tag_primtypes = clashselection_node.Attributes.GetNamedItem("primtypes").InnerText;
                                                        clashtest.Right.Clashselection.Tag_selfintersect = clashselection_node.Attributes.GetNamedItem("selfintersect").InnerText;

                                                        foreach (XmlNode locator_node in clashselection_node.ChildNodes)
                                                        {
                                                            if (locator_node.Name == "locator")
                                                            {
                                                                clashtest.Right.Clashselection.Locator = new Locator();
                                                                clashtest.Right.Clashselection.Locator.Tag_inner_text = locator_node.InnerText;
                                                                //clashtest.Right.Clashselection.Locator.Tag_inner_text_selection_name = locator_node.InnerText.Replace("lcop_selection_set_tree/", "");

                                                                var folders_names = locator_node.InnerText.Split('/');
                                                                clashtest.Right.Clashselection.Locator.Tag_inner_text_selection_name = folders_names.LastOrDefault();
                                                                clashtest.Right.Clashselection.Locator.Tag_inner_text_folders = new List<string>();
                                                                foreach (string folder_name in folders_names)
                                                                {
                                                                    clashtest.Right.Clashselection.Locator.Tag_inner_text_folders.Add(folder_name);
                                                                }

                                                                var draft_names = clashtest.Right.Clashselection.Locator.Tag_inner_text_selection_name.Split('_');
                                                                if (draft_names.Count() > 1)
                                                                {
                                                                    clashtest.Right.Clashselection.Locator.Tag_inner_text_draft_name = draft_names.First();
                                                                }
                                                                else
                                                                {
                                                                    clashtest.Right.Clashselection.Locator.Tag_inner_text_draft_name = "?";
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            else if (node_in_clashtest_node.Name == "rules")
                                            {
                                                clashtest.Rules = new Rules();
                                            }
                                            else if (node_in_clashtest_node.Name == "summary")
                                            {
                                                clashtest.Summary = new Summary()
                                                {
                                                    Tag_total = node_in_clashtest_node.Attributes.GetNamedItem("total").InnerText,
                                                    Tag_new = node_in_clashtest_node.Attributes.GetNamedItem("new").InnerText,
                                                    Tag_active = node_in_clashtest_node.Attributes.GetNamedItem("active").InnerText,
                                                    Tag_reviewed = node_in_clashtest_node.Attributes.GetNamedItem("reviewed").InnerText,
                                                    Tag_approved = node_in_clashtest_node.Attributes.GetNamedItem("approved").InnerText,
                                                    Tag_resolved = node_in_clashtest_node.Attributes.GetNamedItem("resolved").InnerText,
                                                    Testtype = new Testtype(),
                                                    Teststatus = new Teststatus()
                                                };
                                                foreach (XmlNode node_in_summary_node in node_in_clashtest_node.ChildNodes)
                                                {
                                                    if (node_in_summary_node.Name == "testtype")
                                                    {
                                                        clashtest.Summary.Testtype.Tag_inner_text = node_in_summary_node.InnerText;
                                                    }
                                                    else if (node_in_summary_node.Name == "teststatus")
                                                    {
                                                        clashtest.Summary.Teststatus.Tag_inner_text = node_in_summary_node.InnerText;
                                                    }
                                                }
                                                if (int.TryParse(clashtest.Summary.Tag_total, out int int_total)) clashtest.Summary.Tag_total_int = int_total;
                                                if (int.TryParse(clashtest.Summary.Tag_new, out int int_new)) clashtest.Summary.Tag_total_int = int_new;
                                                if (int.TryParse(clashtest.Summary.Tag_active, out int int_active)) clashtest.Summary.Tag_total_int = int_active;
                                                if (int.TryParse(clashtest.Summary.Tag_reviewed, out int int_reviewed)) clashtest.Summary.Tag_total_int = int_reviewed;
                                                if (int.TryParse(clashtest.Summary.Tag_approved, out int int_approved)) clashtest.Summary.Tag_total_int = int_approved;
                                                if (int.TryParse(clashtest.Summary.Tag_resolved, out int int_resolved)) clashtest.Summary.Tag_total_int = int_resolved;
                                            }
                                        }

                                        Clashtests.Add(clashtest);
                                    }
                                }
                            }
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

            List<string> SelectionsetsNamesDistinct = SelectionsetNames.ToArray().Distinct().ToList();

            if (!success) return;

            #endregion


            #region creating array 

            

            #endregion


            int row_models = 0;
            int col_models = 0;
            

          
        }
        private bool CanImportXMLClashtestsExecute(object p) => true;

        /*
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
                    foreach (SelectionSetNode ssn in SelectionSetNodes)
                    {
                        foreach (SelectionSet ss in ssn.SelectionSets)
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
                        foreach (ExcelWorksheet worksheet in worksheets)
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
                    foreach (string draftV in Drafts)
                    {
                        foreach (string draftH in Drafts)
                        {
                            foreach (ClashTest clashTest in ClashTestsTotal)
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

        */
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
