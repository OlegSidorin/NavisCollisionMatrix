using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Navigation;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using СollisionMatrix.Models;

namespace СollisionMatrix
{
    public class MatrixCreatingViewModel : ObservableObject
    {
        JS.Root root;
        public JS.Root Root { get { return root; } set { root = value; OnPropertyChanged(); } }
        public ObservableCollection<MatrixSelectionLineModel> Selections { get; set; }
        public ObservableCollection<UserControl> UserControlsInWholeMatrix { get; set; }
        //public ObservableCollection<UserControl> UserControlsInWholeMatrixHeaders { get; set; }
        public ObservableCollection<UserControl> UserControlsSelectionNames { get; set; }
        public List<Selectionset> Selectionsets { get; set; }
        public List<Clashtest> Clashtests { get; set; }
        public Batchtest Batchtest { get; set; }

        public double widthColumn;
        public double WidthColumn
        {
            get { return widthColumn; }
            set { widthColumn = value; OnPropertyChanged(); }
        }
        public MatrixCreatingViewModel()
        {
            WidthColumn = 183;
            Selectionsets = new List<Selectionset>();
            Clashtests = new List<Clashtest>();
            Batchtest = new Batchtest()
            {
                Tag_name = "Экспорт проверок",
                Tag_internal_name = "Экспорт проверок",
                Tag_units = "ft"
            };

            Selections = new ObservableCollection<MatrixSelectionLineModel>();

            var new1 = new MatrixSelectionLineModel()
            {
                NameOfSelection = "АР | Стены",
                SelectionIntersectionTolerance = new List<string>() 
                {
                    "15", "30", "30", "30"
                }
            };
            var new2 = new MatrixSelectionLineModel()
            {
                NameOfSelection = "АР | Перекрытия",
                SelectionIntersectionTolerance = new List<string>()
                {
                    "", "15", "30", "30"
                }
            };
            var new3 = new MatrixSelectionLineModel()
            {
                NameOfSelection = "ОВ | Воздуховоды, Воздухораспределители, Соединительные детали воздуховодов",
                SelectionIntersectionTolerance = new List<string>()
                {
                    "", "", "15", "30"
                }
            };
            var new4 = new MatrixSelectionLineModel()
            {
                NameOfSelection = "ВК | Трубы, Соединительные детали трубопроводов",
                SelectionIntersectionTolerance = new List<string>()
                {
                    "", "", "", "15"
                }
            };

            Selections.Add(new1);
            Selections.Add(new2);
            Selections.Add(new3);
            Selections.Add(new4);

            UserControlsInWholeMatrix = new ObservableCollection<UserControl>();
            //UserControlsInWholeMatrixHeaders = new ObservableCollection<UserControl>();
            UserControlsSelectionNames = new ObservableCollection<UserControl>();

            foreach (MatrixSelectionLineModel model in Selections)
            {
                MatrixSelectionLineViewModel lineVM = new MatrixSelectionLineViewModel();
                lineVM.MatrixSelectionLineModel = model;
                lineVM.RowNum = Selections.IndexOf(model);
                lineVM.Selections = Selections;
                lineVM.Selection = model;
                
                lineVM.ToleranceViews = new ObservableCollection<UserControl>();
                foreach (string tolerance in model.SelectionIntersectionTolerance)
                {

                    MatrixSelectionCellViewModel cellVM = new MatrixSelectionCellViewModel()
                    {
                        Tolerance = tolerance
                    };
                    MatrixSelectionCellUserControl cellView = new MatrixSelectionCellUserControl();
                    cellView.DataContext = cellVM;

                    lineVM.ToleranceViews.Add(cellView);
                }

                MatrixSelectionLineUserControl userControl = new MatrixSelectionLineUserControl();
                //MatrixSelectionLineHeaderUserControl userControlHeaders = new MatrixSelectionLineHeaderUserControl();
                lineVM.UserControl_MatrixSelectionLineUserControl = userControl;
                lineVM.UserControlsInAllMatrixWithLineUserControls = UserControlsInWholeMatrix;
                lineVM.UserControlsInSelectionNameUserControls = UserControlsSelectionNames;
                userControl.DataContext = lineVM;
                //userControlHeaders.DataContext = lineVM;
                UserControlsInWholeMatrix.Add(userControl);
                //UserControlsInWholeMatrixHeaders.Add(userControlHeaders);
            };

            foreach (UserControl usercontrolline in UserControlsInWholeMatrix)
            {
                MatrixSelectionLineUserControl msluc = (MatrixSelectionLineUserControl)usercontrolline;
                MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;

                MatrixSelectionNameUserControl msnuc = new MatrixSelectionNameUserControl();
                msnuc.DataContext = mslvm;
                UserControlsSelectionNames.Add(msnuc);
            }

            DoIfIClickOnSaveXMLCollisionMatrixButton = new RelayCommand(OnDoIfIClickOnSaveXMLCollisionMatrixButtonExecuted, CanDoIfIClickOnSaveXMLCollisionMatrixButtonExecute);
            DoIfIClickOnOpenXMLCollisionMatrixButton = new RelayCommand(OnDoIfIClickOnOpenXMLCollisionMatrixButtonExecuted, CanDoIfIClickOnOpenXMLCollisionMatrixButtonExecute);

            int linecounter = 0;
            int cellcounter = 0;
            foreach (UserControl usercontrolline in UserControlsInWholeMatrix)
            {
                MatrixSelectionLineUserControl msluc = (MatrixSelectionLineUserControl)usercontrolline;
                MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;
                cellcounter = 0;
                foreach (UserControl usercontrolcell in mslvm.ToleranceViews)
                {
                    MatrixSelectionCellUserControl mscuc = (MatrixSelectionCellUserControl)usercontrolcell;
                    MatrixSelectionCellViewModel mscvm = (MatrixSelectionCellViewModel)mscuc.DataContext;

                    cellcounter += 1;
                }
                linecounter += 1;
            }

        }

        public ICommand DoIfIClickOnSaveXMLCollisionMatrixButton { get; set; }
        private void OnDoIfIClickOnSaveXMLCollisionMatrixButtonExecuted(object p)
        {
            Root = new JS.Root()
            {
                Xml = new JS.Xml()
                {
                    Version = "1.0",
                    Encoding = "UTF-8"
                },
                exchange = new JS.Exchange()
                {
                    XmlnsXsi = @"http://www.w3.org/2001/XMLSchema-instance",
                    XsiNoNamespaceSchemaLocation = @"http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd",
                    Units = "ft",
                    Filename = "",
                    Filepath = "",
                    batchtest = new JS.Batchtest()
                    {
                        Name = "Tests",
                        InternalName = "Tests",
                        Units = "ft",
                        clashtests = new JS.Clashtests()
                        {
                            clashtest = new List<JS.Clashtest>()
                        },
                        selectionsets = new JS.Selectionsets()
                        {
                            selectionset = new List<JS.Selectionset>()
                        }
                    }
                }
            };

            #region fill SelectionSets

            foreach (UserControl usercontrolline in UserControlsInWholeMatrix)
            {
                MatrixSelectionLineUserControl msluc = (MatrixSelectionLineUserControl)usercontrolline;
                MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;

                if (mslvm.JSelectionset == null)
                {
                    #region setting selectionset

                    var ss = new JS.Selectionset()
                    {
                        Name = mslvm.NameOfSelection,
                        Guid = Guid.NewGuid().ToString().ToLower(),
                        findspec = new JS.Findspec()
                        {
                            Mode = "all",
                            Disjoint = "0",
                            conditions = new JS.Conditions()
                            {
                                condition = new List<JS.Condition>()
                            {
                                new JS.Condition()
                                {
                                    Test = "contains",
                                    Flags = "10",
                                    property = new JS.Property()
                                    {
                                        name = new JS.Name()
                                        {
                                            Internal = "LcOaNodeSourceFile",
                                            Text = "Файл источника"
                                        }
                                    },
                                    value = new JS.Value()
                                    {
                                        data = new JS.Data()
                                        {
                                            Type = "wstring",
                                            Text = mslvm.NameOfSelection.Split('|').First().TrimStart(' ').TrimEnd(' ')
                                        }
                                    }
                                }
                            }
                            },
                            locator = @"/"
                        }
                    };

                    if (mslvm.NameOfSelection.Split('|').Count() >= 1)
                    {
                        for (int i = 0; i < mslvm.NameOfSelection.Split('|').Last().Split(',').Count(); i++)
                        {
                            if (i == 0)
                            {
                                ss.findspec.conditions.condition.Add(
                                    new JS.Condition()
                                    {
                                        Test = "equals",
                                        Flags = "10",
                                        category = new JS.Category()
                                        {
                                            name = new JS.Name()
                                            {
                                                Internal = "LcRevitData_Element",
                                                Text = "Объект"
                                            }
                                        },
                                        property = new JS.Property()
                                        {
                                            name = new JS.Name()
                                            {
                                                Internal = "LcRevitPropertyElementCategory",
                                                Text = "Категория"
                                            }
                                        },
                                        value = new JS.Value()
                                        {
                                            data = new JS.Data()
                                            {
                                                Type = "wstring",
                                                Text = mslvm.NameOfSelection.Split('|').Last().Split(',')[i].TrimStart(' ').TrimEnd(' ')
                                            }
                                        }
                                    }
                                    );
                            }
                            else
                            {
                                ss.findspec.conditions.condition.Add(
                                    new JS.Condition()
                                    {
                                        Test = "equals",
                                        Flags = "74",
                                        category = new JS.Category()
                                        {
                                            name = new JS.Name()
                                            {
                                                Internal = "LcRevitData_Element",
                                                Text = "Объект"
                                            }
                                        },
                                        property = new JS.Property()
                                        {
                                            name = new JS.Name()
                                            {
                                                Internal = "LcRevitPropertyElementCategory",
                                                Text = "Категория"
                                            }
                                        },
                                        value = new JS.Value()
                                        {
                                            data = new JS.Data()
                                            {
                                                Type = "wstring",
                                                Text = mslvm.NameOfSelection.Split('|').Last().Split(',')[i].TrimStart(' ').TrimEnd(' ')
                                            }
                                        }
                                    }
                                    );
                            }
                        }
                    }

                    #endregion

                    Root.exchange.batchtest.selectionsets.selectionset.Add(ss);
                }
                else
                {
                    var ss = new JS.Selectionset()
                    {
                        Name = mslvm.NameOfSelection,
                        Guid = mslvm.JSelectionset.Guid,
                        findspec = mslvm.JSelectionset.findspec
                    };

                    Root.exchange.batchtest.selectionsets.selectionset.Add(ss);
                }
            }

            #endregion

            #region fill ClashTests

            string getSelectionName(int i) { return Root.exchange.batchtest.selectionsets.selectionset[i].Name; }
            string toFt(string mm) { if (int.TryParse(mm, out int result)) return (result / 304.8).ToString().Replace(",", "."); else return ""; }

            foreach (UserControl usercontrolline in UserControlsInWholeMatrix)
            {
                MatrixSelectionLineUserControl msluc = (MatrixSelectionLineUserControl)usercontrolline;
                MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;

                int ix = 0;
                foreach (UserControl usercontrolcell in mslvm.ToleranceViews)
                {
                    MatrixSelectionCellUserControl mscuc = (MatrixSelectionCellUserControl)usercontrolcell;
                    MatrixSelectionCellViewModel mscvm = (MatrixSelectionCellViewModel)mscuc.DataContext;

                    if (mscvm.JClashtest == null)
                    {
                        if (!string.IsNullOrEmpty(mscvm.Tolerance))
                        {
                            JS.Clashtest ct = new JS.Clashtest()
                            {
                                Name = $"{mslvm.NameOfSelection} - {getSelectionName(ix)}", // left selset name plus right sel set
                                TestType = "hard",
                                Status = "ok",
                                Tolerance = toFt(mscvm.Tolerance),
                                MergeComposites = "1",
                                linkage = new JS.Linkage() { Mode = "none" },
                                left = new JS.Left()
                                {
                                    clashselection = new JS.Clashselection()
                                    {
                                        Selfintersect = "0",
                                        Primtypes = "1",
                                        locator = $@"lcop_selection_set_tree/{mslvm.NameOfSelection}",
                                    }
                                },
                                right = new JS.Right()
                                {
                                    clashselection = new JS.Clashselection()
                                    {
                                        Selfintersect = "0",
                                        Primtypes = "1",
                                        locator = $@"lcop_selection_set_tree/{getSelectionName(ix)}",
                                    }
                                },
                                rules = null
                            };

                            Root.exchange.batchtest.clashtests.clashtest.Add(ct);
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(mscvm.Tolerance))
                        {
                            JS.Clashtest ct = new JS.Clashtest()
                            {
                                Name = mscvm.JClashtest.Name,
                                TestType = mscvm.JClashtest.TestType,
                                Status = mscvm.JClashtest.Status,
                                Tolerance = toFt(mscvm.Tolerance),
                                MergeComposites = "1",
                                linkage = mscvm.JClashtest.linkage,
                                left = new JS.Left()
                                {
                                    clashselection = new JS.Clashselection()
                                    {
                                        Selfintersect = mscvm.JClashtest.left.clashselection.Selfintersect,
                                        Primtypes = mscvm.JClashtest.left.clashselection.Primtypes,
                                        locator = $@"lcop_selection_set_tree/{mslvm.NameOfSelection}",
                                    }
                                },
                                right = new JS.Right()
                                {
                                    clashselection = new JS.Clashselection()
                                    {
                                        Selfintersect = mscvm.JClashtest.right.clashselection.Selfintersect,
                                        Primtypes = mscvm.JClashtest.right.clashselection.Primtypes,
                                        locator = $@"lcop_selection_set_tree/{getSelectionName(ix)}",
                                    }
                                },
                                rules = mscvm.JClashtest.rules
                            };

                            mscvm.JClashtest = ct;

                            Root.exchange.batchtest.clashtests.clashtest.Add(ct);
                        }
                    }


                    ix++;
                }
            }

            #endregion


            //Selectionsets.Clear();
            //Clashtests.Clear();



            //Debug.WriteLine(Root.exchange.batchtest.clashtests.clashtest.FirstOrDefault().Name);

            //foreach (var item in root.exchange.batchtest.clashtests.clashtest)
            //{
            //    item.Name = item.Name.Replace("2_", "5 : ");
            //}

            string json2 = JsonConvert.SerializeObject(Root, Newtonsoft.Json.Formatting.Indented);

            Debug.WriteLine(json2.ToString());

            XNode node = JsonConvert.DeserializeXNode(json2, null);

            Debug.WriteLine(node.ToString());

            System.Windows.Forms.SaveFileDialog openFileDialog = new System.Windows.Forms.SaveFileDialog();
            System.Windows.Forms.DialogResult dialog_result = openFileDialog.ShowDialog();
            string pathtoxml = openFileDialog.FileName;

            if (dialog_result == System.Windows.Forms.DialogResult.OK)
            {
                if (!pathtoxml.Contains(".xml"))
                {
                    pathtoxml += ".xml";
                }
                if (!File.Exists(pathtoxml))
                {
                    // Create a file to write to.
                    using (StreamWriter sw = File.CreateText(pathtoxml))
                    {
                        sw.Write(node.ToString().Replace(@"<category />", ""));
                    }
                }
            }

            //string path = @"C:\Users\o.sidorin\Downloads\Коммунарка\Navis\output.xml";
            //if (!File.Exists(path))
            //{
            //    // Create a file to write to.
            //    using (StreamWriter sw = File.CreateText(path))
            //    {
            //        sw.Write(node.ToString().Replace(@"<category />", ""));
            //    }
            //}

        }
        private bool CanDoIfIClickOnSaveXMLCollisionMatrixButtonExecute(object p) => true;

        public ICommand DoIfIClickOnOpenXMLCollisionMatrixButton { get; set; }
        private void OnDoIfIClickOnOpenXMLCollisionMatrixButtonExecuted(object p)
        {
            #region reading xml file

            bool success = true;

            string pathtoxml = "";

            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();

            openFileDialog.ShowDialog();

            pathtoxml = openFileDialog.FileName;

            //Xceed.Wpf.Toolkit.MessageBox.Show("Приветики", "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);

            XmlDocument xDoc = new XmlDocument();

            try
            {
                xDoc.Load(pathtoxml);
            }
            catch (Exception ex)
            {
                success = false;
                MessageBox.Show("Возможно, не выбран файл ... \n", "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            if (!success) return;

            #endregion

            string json = JsonConvert.SerializeXmlNode(xDoc);

            Debug.WriteLine(json);

            Root = JsonConvert.DeserializeObject<JS.Root>(json);

            Debug.WriteLine(Root.exchange.batchtest.clashtests.clashtest.FirstOrDefault().Name);

            UserControlsInWholeMatrix.Clear();
            UserControlsSelectionNames.Clear();

            #region reading Selectionsets

            if (Root != null && Root.exchange != null && Root.exchange.batchtest != null && Root.exchange.batchtest.selectionsets != null)
            {
                if (Root.exchange.batchtest.selectionsets.selectionset != null)
                {
                    int i = 0;
                    foreach (JS.Selectionset ss in Root.exchange.batchtest.selectionsets.selectionset)
                    {
                        MatrixSelectionLineViewModel mslvm = new MatrixSelectionLineViewModel()
                        {
                            NameOfSelection = ss.Name,
                            ToleranceViews = new ObservableCollection<UserControl>(),
                            RowNum = i,
                            JSelectionset = ss
                        };
                        
                        MatrixSelectionLineUserControl msluc = new MatrixSelectionLineUserControl();
                        MatrixSelectionNameUserControl msnuc = new MatrixSelectionNameUserControl();
                        mslvm.UserControl_MatrixSelectionLineUserControl = msluc;
                        mslvm.UserControlsInAllMatrixWithLineUserControls = UserControlsInWholeMatrix;
                        mslvm.UserControlsInSelectionNameUserControls = UserControlsSelectionNames;
                        msluc.DataContext = mslvm;
                        msnuc.DataContext = mslvm;

                        for (int ik = 0; ik < Root.exchange.batchtest.selectionsets.selectionset.Count; ik++)
                        {
                            MatrixSelectionCellUserControl mscuc = new MatrixSelectionCellUserControl();
                            MatrixSelectionCellViewModel mscvm = new MatrixSelectionCellViewModel();
                            mscvm.Tolerance = string.Empty;
                            mscuc.DataContext = mscvm;
                            mslvm.ToleranceViews.Add(mscuc);
                        }

                        UserControlsInWholeMatrix.Add(msluc);
                        UserControlsSelectionNames.Add(msnuc);

                        i++;
                    }
                }
            }

            #endregion

            #region reading ClashTests

            string toMm(string ft) { if (double.TryParse(ft.Replace(".", ","), out double result)) return (((int)((result + 0.00001) * 3048) / 10)).ToString(); else return ""; }
            int getIndex(string selectionsetName)
            {
                int ix = 0;
                bool ok = false;
                foreach (MatrixSelectionLineUserControl msluc in UserControlsInWholeMatrix)
                {
                    MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;
                    if (mslvm.NameOfSelection.Equals(selectionsetName))
                    {
                        return ix;
                    }
                    ix++;
                }
                if (!ok) return -1;
                else return ix;
            }
            MatrixSelectionCellViewModel getClashVM(string leftSelectionset, string rightSelectionset)
            {
                foreach (MatrixSelectionLineUserControl msluc in UserControlsInWholeMatrix)
                {
                    MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;

                    if (mslvm.NameOfSelection.Equals(leftSelectionset))
                    {
                        int index = getIndex(rightSelectionset);

                        if (index != -1)
                        {
                            MatrixSelectionCellUserControl mscuc = (MatrixSelectionCellUserControl)mslvm.ToleranceViews[index];
                            MatrixSelectionCellViewModel mscvm = (MatrixSelectionCellViewModel)mscuc.DataContext;
                            return mscvm;
                        }
                        else
                        {
                            return null;
                        }
                    }

                }
                return null;
            }

            if (Root != null && Root.exchange != null && Root.exchange.batchtest != null && Root.exchange.batchtest.clashtests != null)
            {
                if (Root.exchange.batchtest.clashtests.clashtest != null)
                {
                    foreach (JS.Clashtest ct in Root.exchange.batchtest.clashtests.clashtest)
                    {
                        MatrixSelectionCellViewModel mscvm = getClashVM(ct.left.clashselection.locator.Replace(@"lcop_selection_set_tree/", ""), ct.right.clashselection.locator.Replace(@"lcop_selection_set_tree/", ""));
                        if (mscvm != null)
                        {
                            mscvm.Tolerance = toMm(ct.Tolerance);
                            mscvm.JClashtest = ct;
                        }
                    }
                }
            }

            #endregion
        }
        private bool CanDoIfIClickOnOpenXMLCollisionMatrixButtonExecute(object p) => true;

        private string selectedCategory;
        public string SelectedCategory 
        { 
            get { return selectedCategory; } 
            set { selectedCategory = value; OnPropertyChanged(); } 
        }
        public ObservableCollection<string> Categories_in_revit

        {
            get
            {
                return new ObservableCollection<string>()
                {
                    "Стены",
                    "Перекрытия",
                    "Панели витража",
                    "Импосты витража",
                    "Потолки",
                    "Окна",
                    "Двери",
                    "Лестницы",
                    "Крыши",
                    "Несущие колонны",
                    "Каркас несущий",
                    "Фундамент несущей конструкции",
                    "Оборудование",
                    "Арматура воздуховодов",
                    "Воздуховоды",
                    "Воздуховоды по осевой",
                    "Воздухораспределители",
                    "Соединительные детали воздуховодов",
                    "Арматура трубопроводов",
                    "Трубы",
                    "Трубопровод по осевой",
                    "Трубы из базы данных производителя MEP",
                    "Соединительные детали трубопроводов",
                    "Спринклеры",
                    "Сантехнические приборы",
                    "Электрооборудование",
                    "Кабельные лотки",
                    "Соединительные детали кабельных лотков",
                    "Короба",
                    "Соединительные детали коробов",
                    "Осветительные приборы",
                    "Осветительная аппаратура",
                    "Системы пожарной сигнализации",
                    "Датчики",
                    "Устройства связи",
                    "Предохранительные устройства",
                    "Устройства вызова и оповещения",
                    "Телефонные устройства",
                    "Система коммутации",
                    "Обобщенные модели",
                    "Все"
                };
            }
        }

        //public string GetSameCategory(string input_category)
        //{
        //    foreach (string category in Categories_in_revit)
        //    {
        //        if (category.ToLower().Contains(input_category.ToLower())) return category;
        //    }
        //    return "Стены";
        //}

        public void PerformConvertToJson()
        {
            #region reading xml file

            bool success = true;

            string pathtoxml = "";

            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();

            openFileDialog.ShowDialog();

            pathtoxml = openFileDialog.FileName;

            //Xceed.Wpf.Toolkit.MessageBox.Show("Приветики", "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);

            XmlDocument xDoc = new XmlDocument();

            try
            {
                xDoc.Load(pathtoxml);
            }
            catch (Exception ex)
            {
                success = false;
                MessageBox.Show("Возможно, не выбран файл ... \n", "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            if (!success) return;

            #endregion

            string json = JsonConvert.SerializeXmlNode(xDoc);

            Debug.WriteLine(json);


            JS.Root root = JsonConvert.DeserializeObject<JS.Root>(json);

            Debug.WriteLine(root.exchange.batchtest.clashtests.clashtest.FirstOrDefault().Name);

            foreach (var item in root.exchange.batchtest.clashtests.clashtest)
            {
                item.Name = item.Name.Replace("2_", "5 : ");
            }

            string json2 = JsonConvert.SerializeObject(root, Newtonsoft.Json.Formatting.Indented);

            XNode node = JsonConvert.DeserializeXNode(json2, null);

            Debug.WriteLine(node.ToString());

            string path = @"C:\Users\o.sidorin\Downloads\Коммунарка\Navis\output.xml";
            if (!File.Exists(path))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.Write(node.ToString().Replace(@"<category />", ""));
                }
            }

        }

        public void PerformReadXML()
        {
            #region reading xml file

            bool success = true;

            string pathtoxml = "";

            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();

            openFileDialog.ShowDialog();

            pathtoxml = openFileDialog.FileName;

            //Xceed.Wpf.Toolkit.MessageBox.Show("Приветики", "Инфо к сведению", MessageBoxButton.OK, MessageBoxImage.Information);

            XmlDocument xDoc = new XmlDocument();

            try
            {
                xDoc.Load(pathtoxml);
            }
            catch (Exception ex)
            {
                success = false;
                MessageBox.Show("Возможно, не выбран файл ... \n", "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            if (!success) return;

            #endregion

            string json = JsonConvert.SerializeXmlNode(xDoc);

            Debug.WriteLine(json);

            Root = JsonConvert.DeserializeObject<JS.Root>(json);

            Debug.WriteLine(Root.exchange.batchtest.clashtests.clashtest.FirstOrDefault().Name);

            UserControlsInWholeMatrix.Clear();
            UserControlsSelectionNames.Clear();

            #region reading Selectionsets

            if (Root != null && Root.exchange != null && Root.exchange.batchtest != null && Root.exchange.batchtest.selectionsets != null)
            {
                if (Root.exchange.batchtest.selectionsets.selectionset != null)
                {
                    int i = 0;
                    foreach(JS.Selectionset ss in Root.exchange.batchtest.selectionsets.selectionset)
                    {
                        MatrixSelectionLineViewModel mslvm = new MatrixSelectionLineViewModel()
                        {
                            NameOfSelection = ss.Name,
                            ToleranceViews = new ObservableCollection<UserControl>(),
                            RowNum = i,
                            JSelectionset = ss
                        };
                        MatrixSelectionLineUserControl msluc = new MatrixSelectionLineUserControl();
                        MatrixSelectionNameUserControl msnuc = new MatrixSelectionNameUserControl();
                        msluc.DataContext = mslvm;
                        msnuc.DataContext = mslvm;
                        
                        for (int ik = 0; ik < Root.exchange.batchtest.selectionsets.selectionset.Count; ik++)
                        {
                            MatrixSelectionCellUserControl mscuc = new MatrixSelectionCellUserControl();
                            MatrixSelectionCellViewModel mscvm = new MatrixSelectionCellViewModel();
                            mscvm.Tolerance = string.Empty;
                            mscuc.DataContext = mscvm;
                            mslvm.ToleranceViews.Add(mscuc);
                        }

                        UserControlsInWholeMatrix.Add(msluc);
                        UserControlsSelectionNames.Add(msnuc);

                        i++;
                    }
                }
            }

            #endregion

            #region reading ClashTests

            string toMm(string ft) { if (double.TryParse(ft.Replace(".", ","), out double result)) return ((int)(result * 304.8)).ToString(); else return ""; }
            int getIndex(string selectionsetName)
            {
                int ix = 0;
                bool ok = false;
                foreach (MatrixSelectionLineUserControl msluc in UserControlsInWholeMatrix)
                {
                    MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;
                    if (mslvm.NameOfSelection.Equals(selectionsetName))
                    {
                        return ix;
                    }
                    ix++;
                }
                if (!ok) return -1;
                else return ix;
            }
            MatrixSelectionCellViewModel getClashVM(string leftSelectionset, string rightSelectionset)
            {
                foreach(MatrixSelectionLineUserControl msluc in UserControlsInWholeMatrix)
                {
                    MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;

                    if (mslvm.NameOfSelection.Equals(leftSelectionset))
                    {
                        int index = getIndex(rightSelectionset);

                        if (index != -1)
                        {
                            MatrixSelectionCellUserControl mscuc = (MatrixSelectionCellUserControl)mslvm.ToleranceViews[index];
                            MatrixSelectionCellViewModel mscvm = (MatrixSelectionCellViewModel)mscuc.DataContext;
                            return mscvm;
                        }
                        else
                        {
                            return null;
                        }
                    }
                    
                }
                return null;
            }

            if (Root != null && Root.exchange != null && Root.exchange.batchtest != null && Root.exchange.batchtest.clashtests != null)
            {
                if (Root.exchange.batchtest.clashtests.clashtest != null)
                {
                    foreach(JS.Clashtest ct in Root.exchange.batchtest.clashtests.clashtest)
                    {
                        MatrixSelectionCellViewModel mscvm = getClashVM(ct.left.clashselection.locator.Replace(@"lcop_selection_set_tree/", ""), ct.right.clashselection.locator.Replace(@"lcop_selection_set_tree/", ""));
                        if (mscvm != null)
                        {
                            mscvm.Tolerance = toMm(ct.Tolerance);
                            mscvm.JClashtest = ct;
                        }
                    }
                }
            }

            #endregion

        }

        public void PerformSaveXML()
        {
            Root = new JS.Root()
            {
                Xml = new JS.Xml()
                {
                    Version = "1.0",
                    Encoding = "UTF-8"
                },
                exchange = new JS.Exchange()
                {
                    XmlnsXsi = @"http://www.w3.org/2001/XMLSchema-instance",
                    XsiNoNamespaceSchemaLocation = @"http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd",
                    Units = "ft",
                    Filename = "",
                    Filepath = "",
                    batchtest = new JS.Batchtest()
                    {
                        Name = "Tests",
                        InternalName = "Tests",
                        Units = "ft",
                        clashtests = new JS.Clashtests()
                        {
                            clashtest = new List<JS.Clashtest>()
                        },
                        selectionsets = new JS.Selectionsets()
                        {
                            selectionset = new List<JS.Selectionset>()
                        }
                    }
                }
            };

            #region fill SelectionSets

            foreach (UserControl usercontrolline in UserControlsInWholeMatrix)
            {
                MatrixSelectionLineUserControl msluc = (MatrixSelectionLineUserControl)usercontrolline;
                MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;

                if (mslvm.JSelectionset == null)
                {
                    #region setting selectionset

                    var ss = new JS.Selectionset()
                    {
                        Name = mslvm.NameOfSelection,
                        Guid = Guid.NewGuid().ToString().ToLower(),
                        findspec = new JS.Findspec()
                        {
                            Mode = "all",
                            Disjoint = "0",
                            conditions = new JS.Conditions()
                            {
                                condition = new List<JS.Condition>()
                            {
                                new JS.Condition()
                                {
                                    Test = "contains",
                                    Flags = "10",
                                    property = new JS.Property()
                                    {
                                        name = new JS.Name()
                                        {
                                            Internal = "LcOaNodeSourceFile",
                                            Text = "Файл источника"
                                        }
                                    },
                                    value = new JS.Value()
                                    {
                                        data = new JS.Data()
                                        {
                                            Type = "wstring",
                                            Text = mslvm.NameOfSelection.Split('|').First().TrimStart(' ').TrimEnd(' ')
                                        }
                                    }
                                }
                            }
                            },
                            locator = @"/"
                        }
                    };

                    if (mslvm.NameOfSelection.Split('|').Count() >= 1)
                    {
                        for (int i = 0; i < mslvm.NameOfSelection.Split('|').Last().Split(',').Count(); i++)
                        {
                            if (i == 0)
                            {
                                ss.findspec.conditions.condition.Add(
                                    new JS.Condition()
                                    {
                                        Test = "equals",
                                        Flags = "10",
                                        category = new JS.Category()
                                        {
                                            name = new JS.Name()
                                            {
                                                Internal = "LcRevitData_Element",
                                                Text = "Объект"
                                            }
                                        },
                                        property = new JS.Property()
                                        {
                                            name = new JS.Name()
                                            {
                                                Internal = "LcRevitPropertyElementCategory",
                                                Text = "Категория"
                                            }
                                        },
                                        value = new JS.Value()
                                        {
                                            data = new JS.Data()
                                            {
                                                Type = "wstring",
                                                Text = mslvm.NameOfSelection.Split('|').Last().Split(',')[i].TrimStart(' ').TrimEnd(' ')
                                            }
                                        }
                                    }
                                    );
                            }
                            else
                            {
                                ss.findspec.conditions.condition.Add(
                                    new JS.Condition()
                                    {
                                        Test = "equals",
                                        Flags = "74",
                                        category = new JS.Category()
                                        {
                                            name = new JS.Name()
                                            {
                                                Internal = "LcRevitData_Element",
                                                Text = "Объект"
                                            }
                                        },
                                        property = new JS.Property()
                                        {
                                            name = new JS.Name()
                                            {
                                                Internal = "LcRevitPropertyElementCategory",
                                                Text = "Категория"
                                            }
                                        },
                                        value = new JS.Value()
                                        {
                                            data = new JS.Data()
                                            {
                                                Type = "wstring",
                                                Text = mslvm.NameOfSelection.Split('|').Last().Split(',')[i].TrimStart(' ').TrimEnd(' ')
                                            }
                                        }
                                    }
                                    );
                            }
                        }
                    }

                    #endregion

                    Root.exchange.batchtest.selectionsets.selectionset.Add(ss);
                }
                else
                {
                    var ss = new JS.Selectionset()
                    {
                        Name = mslvm.NameOfSelection,
                        Guid = mslvm.JSelectionset.Guid,
                        findspec = mslvm.JSelectionset.findspec
                    };

                    Root.exchange.batchtest.selectionsets.selectionset.Add(ss);
                }
            }

            #endregion

            #region fill ClashTests

            string getSelectionName(int i) { return Root.exchange.batchtest.selectionsets.selectionset[i].Name; }
            string toFt(string mm) { if (double.TryParse(mm, out double result)) return (result / 304.8).ToString().Replace(",", "."); else return ""; }

            foreach (UserControl usercontrolline in UserControlsInWholeMatrix)
            {
                MatrixSelectionLineUserControl msluc = (MatrixSelectionLineUserControl)usercontrolline;
                MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;

                int ix = 0;
                foreach (UserControl usercontrolcell in mslvm.ToleranceViews)
                {
                    MatrixSelectionCellUserControl mscuc = (MatrixSelectionCellUserControl)usercontrolcell;
                    MatrixSelectionCellViewModel mscvm = (MatrixSelectionCellViewModel)mscuc.DataContext;

                    if (mscvm.JClashtest == null)
                    {
                        if (!string.IsNullOrEmpty(mscvm.Tolerance))
                        {
                            JS.Clashtest ct = new JS.Clashtest()
                            {
                                Name = $"{mslvm.NameOfSelection} - {getSelectionName(ix)}", // left selset name plus right sel set
                                TestType = "hard",
                                Status = "ok",
                                Tolerance = toFt(mscvm.Tolerance),
                                MergeComposites = "1",
                                linkage = new JS.Linkage() { Mode = "none" },
                                left = new JS.Left()
                                {
                                    clashselection = new JS.Clashselection()
                                    {
                                        Selfintersect = "0",
                                        Primtypes = "1",
                                        locator = $@"lcop_selection_set_tree/{mslvm.NameOfSelection}",
                                    }
                                },
                                right = new JS.Right()
                                {
                                    clashselection = new JS.Clashselection()
                                    {
                                        Selfintersect = "0",
                                        Primtypes = "1",
                                        locator = $@"lcop_selection_set_tree/{getSelectionName(ix)}",
                                    }
                                },
                                rules = null
                            };

                            Root.exchange.batchtest.clashtests.clashtest.Add(ct);
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(mscvm.Tolerance))
                        {
                            JS.Clashtest ct = new JS.Clashtest()
                            {
                                Name = mscvm.JClashtest.Name,
                                TestType = mscvm.JClashtest.TestType,
                                Status = mscvm.JClashtest.Status,
                                Tolerance = toFt(mscvm.Tolerance),
                                MergeComposites = "1",
                                linkage = mscvm.JClashtest.linkage,
                                left = new JS.Left()
                                {
                                    clashselection = new JS.Clashselection()
                                    {
                                        Selfintersect = mscvm.JClashtest.left.clashselection.Selfintersect,
                                        Primtypes = mscvm.JClashtest.left.clashselection.Primtypes,
                                        locator = $@"lcop_selection_set_tree/{mslvm.NameOfSelection}",
                                    }
                                },
                                right = new JS.Right()
                                {
                                    clashselection = new JS.Clashselection()
                                    {
                                        Selfintersect = mscvm.JClashtest.right.clashselection.Selfintersect,
                                        Primtypes = mscvm.JClashtest.right.clashselection.Primtypes,
                                        locator = $@"lcop_selection_set_tree/{getSelectionName(ix)}",
                                    }
                                },
                                rules = mscvm.JClashtest.rules
                            };

                            mscvm.JClashtest = ct;

                            Root.exchange.batchtest.clashtests.clashtest.Add(ct);
                        }
                    }

                    
                    ix++;
                }
            }

            #endregion


            //Selectionsets.Clear();
            //Clashtests.Clear();



            //Debug.WriteLine(Root.exchange.batchtest.clashtests.clashtest.FirstOrDefault().Name);

            //foreach (var item in root.exchange.batchtest.clashtests.clashtest)
            //{
            //    item.Name = item.Name.Replace("2_", "5 : ");
            //}

            string json2 = JsonConvert.SerializeObject(Root, Newtonsoft.Json.Formatting.Indented);

            Debug.WriteLine(json2.ToString());

            XNode node = JsonConvert.DeserializeXNode(json2, null);

            Debug.WriteLine(node.ToString());

            System.Windows.Forms.SaveFileDialog openFileDialog = new System.Windows.Forms.SaveFileDialog();
            System.Windows.Forms.DialogResult dialog_result = openFileDialog.ShowDialog();
            string pathtoxml = openFileDialog.FileName;

            if (dialog_result == System.Windows.Forms.DialogResult.OK)
            {
                if (!pathtoxml.Contains(".xml"))
                {
                    pathtoxml += ".xml";
                }
                if (!File.Exists(pathtoxml))
                {
                    // Create a file to write to.
                    using (StreamWriter sw = File.CreateText(pathtoxml))
                    {
                        sw.Write(node.ToString().Replace(@"<category />", ""));
                    }
                }
            }

            //string path = @"C:\Users\o.sidorin\Downloads\Коммунарка\Navis\output.xml";
            //if (!File.Exists(path))
            //{
            //    // Create a file to write to.
            //    using (StreamWriter sw = File.CreateText(path))
            //    {
            //        sw.Write(node.ToString().Replace(@"<category />", ""));
            //    }
            //}

        }
    }
}
