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
using System.Xml.Serialization;
using System.Windows.Shapes;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Globalization;

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

        public double widthColumn;
        public double WidthColumn
        {
            get { return widthColumn; }
            set { widthColumn = value; OnPropertyChanged(); }
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
        public ObservableCollection<UserControl> LineUserControls { get; set; }
        public ObservableCollection<UserControl> LineUserControls_Names { get; set; }

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

        public List<CT.Clashtest> Clashtests { get; set; } = new List<CT.Clashtest>();

        public MainWindowModel()
        {
            LineUserControls = new ObservableCollection<UserControl>();
            LineUserControls_Names = new ObservableCollection<UserControl>();

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
            ExcelExport = new RelayCommand(OnExcelExportExecuted, CanExcelExportExecute);

            WidthColumn = 210;

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
            LineUserControls_Names.Clear();
            SelectionsetNames.Clear();

            #region reading xml file

            string pathtoexe = Assembly.GetExecutingAssembly().Location;
            string dllname = Assembly.GetExecutingAssembly().GetName().Name.ToString();
            string folderpath = pathtoexe.Replace(dllname + ".exe", "");
            string pathtoxml = folderpath + "batchtest.xml";

            Forms.OpenFileDialog openFileDialog = new Forms.OpenFileDialog();
            openFileDialog.ShowDialog();
            pathtoxml = openFileDialog.FileName;

            XmlDocument xDoc = new XmlDocument();

            try
            {
                xDoc.Load(pathtoxml);
            }
            catch (Exception ex)
            {
                success = false;
                Xceed.Wpf.Toolkit.MessageBox.Show($"Возможно, не выбран файл и вот что удалось узнать:\n{ex.Message}", "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            string json = JsonConvert.SerializeXmlNode(xDoc);

            CT.Root xmlRoot = null;

            try
            {
                xmlRoot = JsonConvert.DeserializeObject<CT.Root>(json);
            }
            catch (Exception ex)
            {
                success = false;
                Xceed.Wpf.Toolkit.MessageBox.Show($"Возможно, не тот файл и вот что удалось узнать:\n{ex.Message}", "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            if (!success) return;
            if (xmlRoot == null) return;

            #endregion

            #region creating array Clashtests from xml file

            foreach (CT.Clashtest ct in xmlRoot.exchange.batchtest.clashtests.clashtest)
            {
                Clashtests.Add(ct);
            }
            foreach (CT.Clashtest ct in xmlRoot.exchange.batchtest.clashtests.clashtest)
            {
                SelectionsetNames.Add(ct.left?.clashselection?.locator);  //Replace("lcop_selection_set_tree/", ""));
            }
            foreach (CT.Clashtest ct in xmlRoot.exchange.batchtest.clashtests.clashtest)
            {
                SelectionsetNames.Add(ct.right?.clashselection?.locator);  //Replace("lcop_selection_set_tree/", ""));
            }

            List<string> SelectionsetsNamesDistinct = SelectionsetNames.ToArray().Distinct().ToList();

            if (!success) return;

            #endregion


            #region creating array 
            int rn = 0;
            foreach (string sel_name in SelectionsetsNamesDistinct)
            {
                Mainviews.MainLineUserControl mluc = new Mainviews.MainLineUserControl();
                Mainviews.MainLineViewModel mlvm = new Mainviews.MainLineViewModel();
                mlvm.HeaderWidth = WidthColumn - 35;
                mlvm.NameOfSelection = sel_name.Replace("lcop_selection_set_tree/", "");
                mlvm.RowNum = rn.ToString();
                mlvm.CellViews = new ObservableCollection<UserControl>();
                mlvm.Clashtests = new List<CT.Clashtest>();

                Mainviews.MainSelectionNameUserControl msnuc = new Mainviews.MainSelectionNameUserControl(); // это верхняя строка с именами selection set

                foreach (CT.Clashtest clashtest in xmlRoot.exchange.batchtest.clashtests.clashtest)
                {
                    if (clashtest.left?.clashselection?.locator != null )
                    {
                        if (clashtest.left.clashselection.locator.Equals(sel_name)) mlvm.Clashtests.Add(clashtest); 
                    }
                }

                foreach(string sel_name_2 in SelectionsetsNamesDistinct)
                {
                    Mainviews.MainCellUserConrol mcuc = new Mainviews.MainCellUserConrol();
                    Mainviews.MainCellViewModel mcvm = new Mainviews.MainCellViewModel();
                    mcvm.Presenter = "";
                    foreach (CT.Clashtest ct in mlvm.Clashtests)
                    {
                        if (ct.right.clashselection.locator.Equals(sel_name_2))
                        {
                            mcvm.Clashtest = ct;
                            mcvm.Presenter = ct.summary?.total;
                        }
                    }
                    mcuc.DataContext = mcvm;
                    mlvm.CellViews.Add(mcuc);
                }

                mluc.DataContext = mlvm;
                msnuc.DataContext = mlvm;
                LineUserControls.Add(mluc);
                LineUserControls_Names.Add(msnuc);
                
                rn++;
            }
            
            foreach(UserControl uc_line in LineUserControls)
            {
                Mainviews.MainLineViewModel mlvm = (Mainviews.MainLineViewModel)uc_line.DataContext;

                foreach (UserControl uc_cell in mlvm.CellViews)
                {
                    Mainviews.MainCellViewModel mcvm = (Mainviews.MainCellViewModel)uc_cell.DataContext;
                    double dt = Convert.ToDouble(mcvm.Clashtest?.tolerance, CultureInfo.InvariantCulture);
                    if (dt != 0.0)
                    {
                        if (xmlRoot.exchange.batchtest.units.Equals("m"))
                        {
                            if (dt * 1000 <= 50.5) uc_cell.Background = Color50;
                            if (dt * 1000 <= 30.5) uc_cell.Background = Color30;
                            if (dt * 1000 <= 15.5) uc_cell.Background = Color15;
                            if (dt * 1000 > 50.5) uc_cell.Background = Color80;
                            uc_cell.ToolTip =
                                $"{mcvm.Clashtest.name} \n{dt * 1000} мм\n" +
                                $"{mcvm.Clashtest.summary.testtype}: {mcvm.Clashtest.summary.teststatus}\n" +
                                $"Новых: {mcvm.Clashtest.summary.@new}\n" +
                                $"Активных: {mcvm.Clashtest.summary.active}\n" +
                                $"Проверенных: {mcvm.Clashtest.summary.reviewed}\n" +
                                $"Подтвержденных (не коллизия): {mcvm.Clashtest.summary.approved}\n" +
                                $"Исправленных: {mcvm.Clashtest.summary.resolved}";
                        }
                        else if (xmlRoot.exchange.batchtest.units.Equals("ft"))
                        {
                            if (dt * 304.8 <= 50.5) uc_cell.Background = Color50;
                            if (dt * 304.8 <= 30.5) uc_cell.Background = Color30;
                            if (dt * 304.8 <= 15.5) uc_cell.Background = Color15;
                            if (dt * 304.8 > 50.5) uc_cell.Background = Color80;
                            uc_cell.ToolTip =
                                $"{mcvm.Clashtest.name} \n{dt * 304.8} мм\n" +
                                $"{mcvm.Clashtest.summary.testtype}: {mcvm.Clashtest.summary.teststatus}\n" +
                                $"Новых: {mcvm.Clashtest.summary.@new}\n" +
                                $"Активных: {mcvm.Clashtest.summary.active}\n" +
                                $"Проверенных: {mcvm.Clashtest.summary.reviewed}\n" +
                                $"Подтвержденных (не коллизия): {mcvm.Clashtest.summary.approved}\n" +
                                $"Исправленных: {mcvm.Clashtest.summary.resolved}";
                        }
                        else
                        {
                            uc_cell.ToolTip =
                                $"{mcvm.Clashtest.name} \n" +
                                $"{mcvm.Clashtest.summary.testtype}: {mcvm.Clashtest.summary.teststatus}\n" +
                                $"Новых: {mcvm.Clashtest.summary.@new}\n" +
                                $"Активных: {mcvm.Clashtest.summary.active}\n" +
                                $"Проверенных: {mcvm.Clashtest.summary.reviewed}\n" +
                                $"Подтвержденных (не коллизия): {mcvm.Clashtest.summary.approved}\n" +
                                $"Исправленных: {mcvm.Clashtest.summary.resolved}";
                        }
                    }
                }
            }
            
            #endregion

            if (!success) return;

        }
        private bool CanImportXMLClashtestsExecute(object p) => true;


        public ICommand ExcelExport { get; set; }
        private void OnExcelExportExecuted(object p)
        {
            /*
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelPackage excel = new ExcelPackage();

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
            foreach (UserControl uc_line in LineUserControls)
            {
                Mainviews.MainLineViewModel mlvm = (Mainviews.MainLineViewModel)uc_line.DataContext;

                foreach (ExcelWorksheet worksheet in worksheets)
                {
                    worksheet.Cells[rowCount, colCount].Value = "";
                    worksheet.Cells[rowCount, colCount + 1].Value = mlvm.NameOfSelection;

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

            // горизонтальные заголовки сверху
            rowCount = 1;
            colCount = 3;
            foreach (UserControl uc_line in LineUserControls)
            {
                Mainviews.MainLineViewModel mlvm = (Mainviews.MainLineViewModel)uc_line.DataContext;

                foreach (ExcelWorksheet worksheet in worksheets)
                {
                    worksheet.Cells[rowCount, colCount].Value = "";
                    worksheet.Cells[rowCount + 1, colCount].Value = mlvm.NameOfSelection;
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
            #endregion

            

            rowCount = 3;
            colCount = 3;
            foreach (UserControl uc_line in LineUserControls)
            {
                Mainviews.MainLineViewModel mlvm = (Mainviews.MainLineViewModel)uc_line.DataContext;
                colCount = 3;
                foreach (UserControl uc_cell in mlvm.CellViews)
                {
                    Mainviews.MainCellViewModel mcvm = (Mainviews.MainCellViewModel)uc_cell.DataContext;
                    if (mcvm.Clashtest != null)
                    {
                        int summary = Convert.ToInt32(mcvm.Clashtest.summary.@new) + 
                            Convert.ToInt32(mcvm.Clashtest.summary.active) +
                            Convert.ToInt32(mcvm.Clashtest.summary.reviewed); // + approved;
                        foreach (ExcelWorksheet worksheet in worksheets)
                        {
                            if (worksheet.Name.Equals("Matrix"))
                                worksheet.Cells[rowCount, colCount].Value = summary + "(" + mcvm.Clashtest.summary.resolved + ")" + "\n" + mcvm.Clashtest.summary.total;
                            else if (worksheet.Name.Equals("Matrix-New"))
                                worksheet.Cells[rowCount, colCount].Value = Convert.ToInt32(mcvm.Clashtest.summary.@new);
                            else if (worksheet.Name.Equals("Matrix-Active"))
                                worksheet.Cells[rowCount, colCount].Value = Convert.ToInt32(mcvm.Clashtest.summary.active);
                            else if (worksheet.Name.Equals("Matrix-Reviewed"))
                                worksheet.Cells[rowCount, colCount].Value = Convert.ToInt32(mcvm.Clashtest.summary.reviewed);
                            else if (worksheet.Name.Equals("Matrix-Approved"))
                                worksheet.Cells[rowCount, colCount].Value = Convert.ToInt32(mcvm.Clashtest.summary.approved);
                            else if (worksheet.Name.Equals("Matrix-Resolved"))
                                worksheet.Cells[rowCount, colCount].Value = Convert.ToInt32(mcvm.Clashtest.summary.resolved);
                            else if (worksheet.Name.Equals("Matrix-Total"))
                                worksheet.Cells[rowCount, colCount].Value = Convert.ToInt32(mcvm.Clashtest.summary.total);

                            if (Convert.ToInt32(mcvm.Clashtest.summary.total) == 0)
                            {
                                worksheet.Cells[rowCount, colCount].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(236, 236, 236));
                            }
                            else if (mcvm.Clashtest.summary.resolved == mcvm.Clashtest.summary.total)
                            {
                                worksheet.Cells[rowCount, colCount].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                worksheet.Cells[rowCount, colCount].Style.Fill.SetBackground(System.Drawing.Color.FromArgb(202, 255, 191));
                            }
                            else if (Convert.ToInt32(mcvm.Clashtest.summary.resolved) == 0)
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
                    colCount++;
                }
                rowCount++;
            }

            if (Clashtests == null) return;

            // выравнивание содержимого ячеек результатов проверок
            for (int i = 1; i <= Clashtests.Count + 2; i++)
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
                rowCount = LineUserControls.Count + 4;

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

            Forms.SaveFileDialog saveFileDialog = new Forms.SaveFileDialog();
            Forms.DialogResult saving_result = saveFileDialog.ShowDialog();
            string pathtoxls = saveFileDialog.FileName;

            if (saving_result == Forms.DialogResult.OK)
            {
                if (!pathtoxls.Contains(".xlsx")) pathtoxls += ".xlsx";
                FileInfo excelFile = new FileInfo(pathtoxls);
                excel.SaveAs(excelFile);
            }


            var choose = Xceed.Wpf.Toolkit.MessageBox.Show("Сохранен файл с отчетом:\n" + pathtoxls, "Вопросик", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (choose == MessageBoxResult.Yes)
            {
                System.Diagnostics.Process.Start(pathtoxls);
            }
            */
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
