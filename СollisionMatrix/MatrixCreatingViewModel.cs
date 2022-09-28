using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Xml;
using System.Xml.Serialization;
using СollisionMatrix.Models;

namespace СollisionMatrix
{
    public class MatrixCreatingViewModel : ObservableObject
    {
        public ObservableCollection<MatrixSelectionLineModel> Selections { get; set; }
        public ObservableCollection<UserControl> UserControlsInWholeMatrix { get; set; }
        public ObservableCollection<UserControl> UserControlsInWholeMatrixHeaders { get; set; }
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
            UserControlsInWholeMatrixHeaders = new ObservableCollection<UserControl>();
            UserControlsSelectionNames = new ObservableCollection<UserControl>();
            foreach (MatrixSelectionLineModel model in Selections)
            {
                MatrixSelectionLineViewModel userControlvm = new MatrixSelectionLineViewModel();
                userControlvm.MatrixSelectionLineModel = model;
                userControlvm.RowNum = Selections.IndexOf(model);
                userControlvm.Selections = Selections;
                userControlvm.Selection = model;
                
                userControlvm.ToleranceViews = new ObservableCollection<UserControl>();
                foreach (string tolerance in model.SelectionIntersectionTolerance)
                {

                    MatrixSelectionCellViewModel cellViewModel = new MatrixSelectionCellViewModel()
                    {
                        Tolerance = tolerance
                    };
                    MatrixSelectionCellUserControl cellView = new MatrixSelectionCellUserControl();
                    cellView.DataContext = cellViewModel;

                    userControlvm.ToleranceViews.Add(cellView);
                }

                MatrixSelectionLineUserControl userControl = new MatrixSelectionLineUserControl();
                MatrixSelectionLineHeaderUserControl userControlHeaders = new MatrixSelectionLineHeaderUserControl();
                userControlvm.UserControl_MatrixSelectionLineUserControl = userControl;
                userControlvm.UserControlsInAllMatrixWithLineUserControls = UserControlsInWholeMatrix;
                userControlvm.UserControlsInSelectionNameUserControls = UserControlsSelectionNames;
                userControl.DataContext = userControlvm;
                userControlHeaders.DataContext = userControlvm;
                UserControlsInWholeMatrix.Add(userControl);
                UserControlsInWholeMatrixHeaders.Add(userControlHeaders);
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

            // refresh Selectionsets and Clashtests from view models
            Selectionsets.Clear();
            Clashtests.Clear();
            int iForToleranceArray = 0;
            foreach (UserControl usercontrolline in UserControlsInWholeMatrix)
            {
                iForToleranceArray = 0;
                MatrixSelectionLineUserControl msluc = (MatrixSelectionLineUserControl)usercontrolline;
                MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;

                var part1ofSelectionName = "?AR";
                var part2ofSelectionName = "Стены";
                var partsOfSelectionName = mslvm.NameOfSelection.Split('|');

                if (partsOfSelectionName.Count() > 1)
                {
                    part1ofSelectionName = partsOfSelectionName.First().TrimStart(' ').TrimEnd(' ');
                    part2ofSelectionName = partsOfSelectionName.Last().TrimStart(' ').TrimEnd(' ');
                }

                string[] draft_category_names = part2ofSelectionName.Split(',');
                List<string> category_names = new List<string>();
                foreach (string category in draft_category_names)
                {
                    category_names.Add(category.TrimStart(' ').TrimEnd(' '));
                }

                if (mslvm.Selectionset != null)
                {
                    mslvm.Selectionset.Tag_name = mslvm.NameOfSelection;
                }

                else if (mslvm.Selectionset == null)
                {
                    mslvm.Selectionset = new Selectionset()
                    {
                        Tag_name = mslvm.NameOfSelection,
                        Tag_guid = "",
                        Findspec = new Findspec()
                        {
                            Tag_mode = "all",
                            Tag_disjoint = "0",
                            Locator = new Locator()
                            {
                                Tag_inner_text = "/"
                            },
                            Conditions = new Conditions()
                            {
                                Conditions_list = new List<Models.Condition>()
                            }
                        }
                    };

                    Models.Condition mc1 = new Models.Condition()
                    {
                        Tag_test = "contains",
                        Tag_flags = "10",
                        Property = new Property()
                        {
                            Name = new Name()
                            {
                                Tag_internal = "LcOaNodeSourceFile",
                                Tag_inner_text = "Файл источника"
                            }
                        },
                        Value = new Value()
                        {
                            Data = new Data()
                            {
                                Tag_type = "wstring",
                                Tag_inner_text = part1ofSelectionName
                            }
                        }
                    };
                    mslvm.Selectionset.Findspec.Conditions.Conditions_list.Add(mc1);

                    Models.Condition mc2 = new Models.Condition()
                    {
                        Tag_test = "equals",
                        Tag_flags = "10",
                        Category = new Category()
                        {
                            Name = new Name()
                            {
                                Tag_internal = "LcRevitData_Element",
                                Tag_inner_text = "Объект"
                            }
                        },
                        Property = new Property()
                        {
                            Name = new Name()
                            {
                                Tag_internal = "LcRevitPropertyElementCategory",
                                Tag_inner_text = "Категория"
                            }
                        },
                        Value = new Value()
                        {
                            Data = new Data()
                            {
                                Tag_type = "wstring",
                                Tag_inner_text = GetSameCategory(category_names.FirstOrDefault())
                            }
                        }
                    };

                    mslvm.Selectionset.Findspec.Conditions.Conditions_list.Add(mc2);

                    if (category_names.Count > 1)
                    {
                        int i = 1;
                        for (i = 1; i < category_names.Count; i++)
                        {
                            Models.Condition mc3 = new Models.Condition()
                            {
                                Tag_test = "equals",
                                Tag_flags = "74",
                                Category = new Category()
                                {
                                    Name = new Name()
                                    {
                                        Tag_internal = "LcRevitData_Element",
                                        Tag_inner_text = "Объект"
                                    }
                                },
                                Property = new Property()
                                {
                                    Name = new Name()
                                    {
                                        Tag_internal = "LcRevitPropertyElementCategory",
                                        Tag_inner_text = "Категория"
                                    }
                                },
                                Value = new Value()
                                {
                                    Data = new Data()
                                    {
                                        Tag_type = "wstring",
                                        Tag_inner_text = GetSameCategory(category_names[i])
                                    }
                                }
                            };

                            mslvm.Selectionset.Findspec.Conditions.Conditions_list.Add(mc3);
                        }
                       
                    }

                }

                Selectionsets.Add(mslvm.Selectionset);

                int colNum = 0;
                foreach (UserControl usercontrolcell in mslvm.ToleranceViews)
                {
                    MatrixSelectionCellUserControl mscuc = (MatrixSelectionCellUserControl)usercontrolcell;
                    MatrixSelectionCellViewModel mscvm = (MatrixSelectionCellViewModel)mscuc.DataContext;

                    Clashtest ct = null;
                    if (mscvm.Clashtest != null)
                    {
                        bool operation_result_toInt_isOk = int.TryParse(mscvm.Tolerance, out int tolerance_int);
                        if (operation_result_toInt_isOk)
                        {
                            string dl_name = "?";
                            string[] dl_names = mslvm.NameOfSelection.Split('|');
                            if (dl_names.Count() > 1) dl_name = dl_names.First().TrimStart(' ').TrimEnd(' ');
                            string dr_name = "?";
                            MatrixSelectionLineViewModel vm = (MatrixSelectionLineViewModel)UserControlsSelectionNames.ElementAt(colNum).DataContext;
                            string[] dr_names = vm.NameOfSelection.Split('|');
                            if (dr_names.Count() > 1) dr_name = dr_names.First().TrimStart(' ').TrimEnd(' ');

                            ct = mscvm.Clashtest;

                            ct.Left.Clashselection.Locator.Tag_inner_text_selection_name = mslvm.NameOfSelection;
                            ct.Left.Clashselection.Locator.Tag_inner_text_draft_name = dl_name;
                            string folder_path_l = "";
                            if (ct.Left.Clashselection.Locator.Tag_inner_text_folders.Count() > 1)
                            {
                                for (var i = 0; i < ct.Left.Clashselection.Locator.Tag_inner_text_folders.Count() - 1; i++)
                                {
                                    folder_path_l += ct.Left.Clashselection.Locator.Tag_inner_text_folders.ElementAt(i) + "/";
                                }
                            }
                            ct.Left.Clashselection.Locator.Tag_inner_text = folder_path_l + mslvm.NameOfSelection;

                            ct.Right.Clashselection.Locator.Tag_inner_text_selection_name = vm.NameOfSelection;
                            ct.Right.Clashselection.Locator.Tag_inner_text_draft_name = dr_name;
                            string folder_path_r = "";
                            if (ct.Right.Clashselection.Locator.Tag_inner_text_folders.Count() > 1)
                            {
                                for (var i = 0; i < ct.Right.Clashselection.Locator.Tag_inner_text_folders.Count() - 1; i++)
                                {
                                    folder_path_r += ct.Right.Clashselection.Locator.Tag_inner_text_folders.ElementAt(i) + "/";
                                }
                            }
                            ct.Right.Clashselection.Locator.Tag_inner_text = folder_path_r + vm.NameOfSelection;
                            ct.Tag_name = ct.Left.Clashselection.Locator.Tag_inner_text_selection_name + " - " + ct.Right.Clashselection.Locator.Tag_inner_text_selection_name;
                            ct.Tag_tolerance = (tolerance_int / 304.78).ToString().Replace(",", ".").Substring(0, 12);
                            ct.Tag_tolerance_in_mm = tolerance_int.ToString();
                            
                        }
                        
                    }
                    else if (mscvm.Clashtest == null)
                    {
                        bool operation_result_toInt_isOk = int.TryParse(mscvm.Tolerance, out int tolerance_int);
                        if (operation_result_toInt_isOk)
                        {
                            string dl_name = "?";
                            string[] dl_names = mslvm.NameOfSelection.Split('|');
                            if (dl_names.Count() > 1) dl_name = dl_names.First().TrimStart(' ').TrimEnd(' ');
                            string dr_name = "?";

                            MatrixSelectionLineViewModel vm = (MatrixSelectionLineViewModel)UserControlsSelectionNames.ElementAt(colNum).DataContext;
                            string[] dr_names = vm.NameOfSelection.Split('|');
                            if (dr_names.Count() > 1) dr_name = dr_names.First().TrimStart(' ').TrimEnd(' ');

                            ct = new Clashtest()
                            {
                                Tag_name = mslvm.NameOfSelection + " - " + vm.NameOfSelection,
                                Tag_test_type = "hard",
                                Tag_status = "new",
                                Tag_merge_composites = "1",
                                Tag_tolerance_in_mm = tolerance_int.ToString(),
                                Tag_tolerance = (tolerance_int / 304.78).ToString().Replace(",", ".").Substring(0, 12),
                                Left = new Left()
                                {
                                    Clashselection = new Clashselection()
                                    {
                                        Tag_selfintersect = "0",
                                        Tag_primtypes = "1",
                                        Locator = new Locator()
                                        {
                                            Tag_inner_text = @"lcop_selection_set_tree/" + mslvm.NameOfSelection,
                                            Tag_inner_text_draft_name = dl_name,
                                            Tag_inner_text_folders = new List<string>()
                                            {
                                                 "lcop_selection_set_tree"
                                            },
                                            Tag_inner_text_selection_name = mslvm.NameOfSelection
                                        }
                                    }
                                },
                                Right = new Right()
                                {
                                    Clashselection = new Clashselection()
                                    {
                                        Tag_selfintersect = "0",
                                        Tag_primtypes = "1",
                                        Locator = new Locator()
                                        {
                                            Tag_inner_text = @"lcop_selection_set_tree/" + vm.NameOfSelection,
                                            Tag_inner_text_draft_name = dr_name,
                                            Tag_inner_text_folders = new List<string>()
                                            {
                                                "lcop_selection_set_tree"
                                            },
                                            Tag_inner_text_selection_name = vm.NameOfSelection
                                        }
                                    }
                                },
                                Linkage = new Linkage()
                                {
                                    Tag_mode = "none"
                                },
                                Rules = new Rules()
                        };
                        }
                    }

                    if (ct != null) Clashtests.Add(ct);

                    colNum += 1;
                }
                    
            }

            string st100 = "";

            // write xml from Selectionsets and Clashtests

            XmlDocument xDoc = new XmlDocument();

            XmlDeclaration xmlDeclaration = xDoc.CreateXmlDeclaration("1.0", "UTF-8", string.Empty);
            XmlNode xmlDeclaration_node = xDoc.AppendChild(xmlDeclaration);


            XmlElement exchange_element = xDoc.CreateElement("exchange");
            exchange_element.SetAttribute(@"xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
            exchange_element.SetAttribute("noNamespaceSchemaLocation", "http://www.w3.org/2001/XMLSchema-instance", "http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd");
            //exchange_element.SetAttribute("xsi:noNamespaceSchemaLocation", "http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd"); // not working
            exchange_element.SetAttribute("units", "ft");
            exchange_element.SetAttribute("filename", "");
            exchange_element.SetAttribute("filepath", "");

            XmlNode exchange_node = xDoc.AppendChild(exchange_element);

            XmlElement batchtest_element = xDoc.CreateElement("batchtest");
            batchtest_element.SetAttribute("name", Batchtest.Tag_name);
            batchtest_element.SetAttribute("internal_name", Batchtest.Tag_internal_name);
            batchtest_element.SetAttribute("units", Batchtest.Tag_units);
            XmlNode batchtest_node = exchange_node.AppendChild(batchtest_element);

            XmlElement clashtests_element = xDoc.CreateElement("clashtests");
            XmlNode clashtests_node = batchtest_node.AppendChild(clashtests_element);
            
            foreach (Clashtest clashtest in Clashtests)
            {
                XmlElement clashtest_element = xDoc.CreateElement("clashtest");
                clashtest_element.SetAttribute("name", clashtest.Tag_name);
                clashtest_element.SetAttribute("test_type", clashtest.Tag_test_type);
                clashtest_element.SetAttribute("status", clashtest.Tag_status);
                clashtest_element.SetAttribute("tolerance", clashtest.Tag_tolerance);
                clashtest_element.SetAttribute("merge_composites", clashtest.Tag_merge_composites);
                XmlNode clashtest_node = clashtests_node.AppendChild(clashtest_element);

                // <linkage mode="none"/>
                XmlElement linkage_element = xDoc.CreateElement("linkage");
                linkage_element.SetAttribute("mode", clashtest.Linkage.Tag_mode);
                XmlNode linkage_node = clashtest_node.AppendChild(linkage_element);

                // <left>
                XmlElement left_element = xDoc.CreateElement("left");
                XmlNode left_node = clashtest_node.AppendChild(left_element);

                // <clashselection selfintersect="0" primtypes="1">
                XmlElement clashselection_element = xDoc.CreateElement("clashselection");
                clashselection_element.SetAttribute("selfintersect", clashtest.Left.Clashselection.Tag_selfintersect);
                clashselection_element.SetAttribute("primtypes", clashtest.Left.Clashselection.Tag_primtypes);
                XmlNode clashselection_node = left_node.AppendChild(clashselection_element);

                // <locator>lcop_selection_set_tree/АР_Стены</locator>
                XmlElement locator_element = xDoc.CreateElement("locator");
                locator_element.InnerText = clashtest.Left.Clashselection.Locator.Tag_inner_text;
                XmlNode locator_node = clashselection_node.AppendChild(locator_element);

                // <right>
                XmlElement right_element = xDoc.CreateElement("right");
                XmlNode right_node = clashtest_node.AppendChild(right_element);

                // <clashselection selfintersect="0" primtypes="1">
                XmlElement clashselectionR_element = xDoc.CreateElement("clashselection");
                clashselectionR_element.SetAttribute("selfintersect", clashtest.Right.Clashselection.Tag_selfintersect);
                clashselectionR_element.SetAttribute("primtypes", clashtest.Right.Clashselection.Tag_primtypes);
                XmlNode clashselectionR_node = right_node.AppendChild(clashselectionR_element);

                // <locator>lcop_selection_set_tree/АР_Стены</locator>
                XmlElement locatorR_element = xDoc.CreateElement("locator");
                locatorR_element.InnerText = clashtest.Right.Clashselection.Locator.Tag_inner_text; ;
                XmlNode locatorR_node = clashselectionR_node.AppendChild(locatorR_element);

                // <rules/>
                XmlElement rules_element = xDoc.CreateElement("rules");
                XmlNode rules_node = clashtest_node.AppendChild(rules_element);
            }

            XmlElement selectionsets_element = xDoc.CreateElement("selectionsets");
            XmlNode selectionsets_node = batchtest_node.AppendChild(selectionsets_element);

            foreach (Selectionset selectionset in Selectionsets)
            {
                XmlElement selectionset_element = xDoc.CreateElement("selectionset");
                selectionset_element.SetAttribute("name", selectionset.Tag_name);
                selectionset_element.SetAttribute("guid", selectionset.Tag_guid);
                XmlNode selectionset_node = selectionsets_node.AppendChild(selectionset_element);

                // <findspec mode="all" disjoint="0">
                XmlElement findspec_element = xDoc.CreateElement("findspec");
                findspec_element.SetAttribute("mode", selectionset.Findspec.Tag_mode);
                findspec_element.SetAttribute("disjoint", selectionset.Findspec.Tag_disjoint);
                XmlNode findspec_node = selectionset_node.AppendChild(findspec_element);

                // <conditions>
                XmlElement conditions_element = xDoc.CreateElement("conditions");
                XmlNode conditions_node = findspec_node.AppendChild(conditions_element);

                foreach (var condition in selectionset.Findspec.Conditions.Conditions_list)
                {
                    // <condition test="contains" flags="10">
                    XmlElement condition_element = xDoc.CreateElement("condition");
                    condition_element.SetAttribute("test", condition.Tag_test);
                    condition_element.SetAttribute("flags", condition.Tag_flags);
                    XmlNode condition_node = conditions_node.AppendChild(condition_element);

                    // <category>
                    if (condition.Category != null)
                    {
                        XmlElement category_element = xDoc.CreateElement("category");
                        XmlNode category_node = condition_node.AppendChild(category_element);
                        // <name internal="LcRevitData_Element">Объект</name>
                        XmlElement name_element = xDoc.CreateElement("name");
                        name_element.SetAttribute("internal", condition.Category.Name.Tag_internal);
                        name_element.InnerText = condition.Category.Name.Tag_inner_text;
                        XmlNode name_node = category_node.AppendChild(name_element);
                    }
                    

                    // <property>
                    if (condition.Property != null)
                    {
                        XmlElement property_element = xDoc.CreateElement("property");
                        XmlNode property_node = condition_node.AppendChild(property_element);
                        // <name internal="LcRevitPropertyElementCategory">Категория</name>
                        XmlElement name_element_p = xDoc.CreateElement("name");
                        name_element_p.SetAttribute("internal", condition.Property.Name.Tag_internal);
                        name_element_p.InnerText = condition.Property.Name.Tag_inner_text;
                        XmlNode name_node_p = property_node.AppendChild(name_element_p);
                    }
                   

                    // <value>
                    if (condition.Value != null)
                    {
                        XmlElement value_element = xDoc.CreateElement("value");
                        XmlNode value_node = condition_node.AppendChild(value_element);
                        // <data type="wstring">Стены</data>
                        XmlElement data_element = xDoc.CreateElement("data");
                        data_element.SetAttribute("type", condition.Value.Data.Tag_type);
                        data_element.InnerText = condition.Value.Data.Tag_inner_text;
                        XmlNode data_node = value_node.AppendChild(data_element);
                    }

                }

                // <locator>/</locator>
                XmlElement locator_element = xDoc.CreateElement("locator");
                locator_element.InnerText = selectionset.Findspec.Locator.Tag_inner_text;
                XmlNode locator_node = findspec_node.AppendChild(locator_element);
            }

            System.Windows.Forms.SaveFileDialog openFileDialog = new System.Windows.Forms.SaveFileDialog();
            System.Windows.Forms.DialogResult dialog_result = openFileDialog.ShowDialog();
            string pathtoxml = openFileDialog.FileName;

            if (dialog_result == System.Windows.Forms.DialogResult.OK)
            {
                if (!pathtoxml.Contains(".xml"))
                {
                    pathtoxml += ".xml";
                }
                xDoc.Save(pathtoxml);
            }

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


            UserControlsInWholeMatrix.Clear();
            UserControlsSelectionNames.Clear();

            Selectionsets.Clear();
            Clashtests.Clear();

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
                                        td *= 304.8;
                                        clashtest.Tag_tolerance_in_mm = ((int)td).ToString();

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

                                                                var draft_names = clashtest.Left.Clashselection.Locator.Tag_inner_text_selection_name.Split('|');
                                                                if (draft_names.Count() > 1)
                                                                {
                                                                    clashtest.Left.Clashselection.Locator.Tag_inner_text_draft_name = draft_names.First().TrimStart(' ').TrimEnd(' ');
                                                                }
                                                                else
                                                                {
                                                                    clashtest.Left.Clashselection.Locator.Tag_inner_text_draft_name = "?";
                                                                }
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

                                                                var draft_names = clashtest.Right.Clashselection.Locator.Tag_inner_text_selection_name.Split('|');
                                                                if (draft_names.Count() > 1)
                                                                {
                                                                    clashtest.Right.Clashselection.Locator.Tag_inner_text_draft_name = draft_names.First().TrimStart(' ').TrimEnd(' ');
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
                                        }

                                        Clashtests.Add(clashtest);
                                    }
                                }
                            }
                            else if (node_in_batchtest_node.Name == "selectionsets")
                            {
                                foreach (XmlNode selectionset_node in node_in_batchtest_node.ChildNodes)
                                {
                                    if (selectionset_node.Name == "selectionset")
                                    {
                                        Selectionset selectionset = new Selectionset();
                                        selectionset.Tag_name = selectionset_node.Attributes.GetNamedItem("name").InnerText;
                                        selectionset.Tag_guid = selectionset_node.Attributes.GetNamedItem("guid").InnerText;
                                        foreach (XmlNode findspec_node in selectionset_node.ChildNodes)
                                        {
                                            if (findspec_node.Name == "findspec")
                                            {
                                                selectionset.Findspec = new Findspec();
                                                selectionset.Findspec.Tag_mode = findspec_node.Attributes.GetNamedItem("mode").InnerText;
                                                selectionset.Findspec.Tag_disjoint = findspec_node.Attributes.GetNamedItem("disjoint").InnerText;
                                                selectionset.Findspec.Conditions = new Conditions();
                                                foreach (XmlNode node_in_findspec_node in findspec_node.ChildNodes)
                                                {
                                                    if (node_in_findspec_node.Name == "conditions")
                                                    {
                                                        selectionset.Findspec.Conditions.Conditions_list = new List<Models.Condition>();
                                                        foreach (XmlNode condition_node in node_in_findspec_node.ChildNodes)
                                                        {
                                                            if (condition_node.Name == "condition")
                                                            {
                                                                Models.Condition condition = new Models.Condition();
                                                                condition.Tag_test = condition_node.Attributes.GetNamedItem("test").InnerText;
                                                                condition.Tag_flags = condition_node.Attributes.GetNamedItem("flags").InnerText;
                                                                foreach (XmlNode node_in_condition_node in condition_node.ChildNodes)
                                                                {
                                                                    if (node_in_condition_node.Name == "category")
                                                                    {
                                                                        condition.Category = new Category();
                                                                        foreach (XmlNode name_node in node_in_condition_node.ChildNodes)
                                                                        {
                                                                            if (name_node.Name == "name")
                                                                            {
                                                                                condition.Category.Name = new Name();
                                                                                condition.Category.Name.Tag_internal = name_node.Attributes.GetNamedItem("internal").InnerText;
                                                                                condition.Category.Name.Tag_inner_text = name_node.InnerText;
                                                                            }
                                                                        }
                                                                    }
                                                                    else if (node_in_condition_node.Name == "property")
                                                                    {
                                                                        condition.Property = new Property();
                                                                        foreach (XmlNode name_node in node_in_condition_node.ChildNodes)
                                                                        {
                                                                            if (name_node.Name == "name")
                                                                            {
                                                                                condition.Property.Name = new Name();
                                                                                condition.Property.Name.Tag_internal = name_node.Attributes.GetNamedItem("internal").InnerText;
                                                                                condition.Property.Name.Tag_inner_text = name_node.InnerText;
                                                                            }
                                                                        }
                                                                    }
                                                                    else if (node_in_condition_node.Name == "value")
                                                                    {
                                                                        condition.Value = new Value();
                                                                        foreach (XmlNode data_node in node_in_condition_node.ChildNodes)
                                                                        {
                                                                            if (data_node.Name == "data")
                                                                            {
                                                                                condition.Value.Data = new Data();
                                                                                condition.Value.Data.Tag_type = data_node.Attributes.GetNamedItem("type").InnerText;
                                                                                condition.Value.Data.Tag_inner_text = data_node.InnerText;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                selectionset.Findspec.Conditions.Conditions_list.Add(condition);
                                                            }
                                                            
                                                        }
                                                        
                                                    }
                                                    else if (node_in_findspec_node.Name == "locator")
                                                    {
                                                        selectionset.Findspec.Locator = new Locator();
                                                        selectionset.Findspec.Locator.Tag_inner_text = node_in_findspec_node.InnerText;
                                                    }
                                                }
                                                
                                            }
                                        }
                                        Selectionsets.Add(selectionset);
                                    }
                                }
                            }
                        }
                    }
                    
                }
                
            }
            catch (Exception ex)
            {
                success = false;
                MessageBox.Show("Возможно, файл не содержит результатов проверки... \n" + ex.ToString(), "Ошибочка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            if (!success) return;

            UserControlsInWholeMatrix.Clear();
            UserControlsSelectionNames.Clear();

            int row_num = 0;
            foreach (Selectionset ss in Selectionsets)
            {
                MatrixSelectionLineUserControl msluc_new = new MatrixSelectionLineUserControl();
                MatrixSelectionLineViewModel mslvm_new = new MatrixSelectionLineViewModel();
                mslvm_new.HeaderWidth = WidthColumn - 65;
                mslvm_new.Selectionset = ss;
                mslvm_new.NameOfSelection = ss.Tag_name;
                mslvm_new.RowNum = row_num;
                mslvm_new.ToleranceViews = new ObservableCollection<UserControl>();
                msluc_new.DataContext = mslvm_new;
                MatrixSelectionNameUserControl msnuc_new = new MatrixSelectionNameUserControl();
                msnuc_new.DataContext = mslvm_new;
                mslvm_new.UserControlsInAllMatrixWithLineUserControls = UserControlsInWholeMatrix;
                mslvm_new.UserControlsInSelectionNameUserControls = UserControlsSelectionNames;
                mslvm_new.UserControl_MatrixSelectionLineUserControl = msluc_new;

                List<Clashtest> clashTests_with_left_name_of_selection_is_equal_name_of_selection = new List<Clashtest>();
                foreach (Clashtest ct in Clashtests)
                {
                    if (ct.Left.Clashselection.Locator.Tag_inner_text_selection_name == ss.Tag_name)
                    {
                        clashTests_with_left_name_of_selection_is_equal_name_of_selection.Add(ct);
                    }
                }

                for (int i = 0; i < Selectionsets.Count(); i++)
                {
                    bool view_with_clastest = false;
                    Clashtest clashtest = null;
                    foreach (Clashtest ct in clashTests_with_left_name_of_selection_is_equal_name_of_selection)
                    {
                        if (ct.Right.Clashselection.Locator.Tag_inner_text_selection_name == Selectionsets.ElementAt(i).Tag_name)
                        {
                            view_with_clastest = true;
                            clashtest = ct;
                        }
                    }
                    if (view_with_clastest)
                    {
                        MatrixSelectionCellUserControl mscuc_new = new MatrixSelectionCellUserControl();
                        MatrixSelectionCellViewModel mscvm_new = new MatrixSelectionCellViewModel();
                        mscvm_new.Clashtest = clashtest;
                        mscvm_new.Tolerance = clashtest.Tag_tolerance_in_mm;
                        mscuc_new.DataContext = mscvm_new;
                        mslvm_new.ToleranceViews.Add(mscuc_new);
                    }
                    else
                    {
                        MatrixSelectionCellUserControl mscuc_new = new MatrixSelectionCellUserControl();
                        MatrixSelectionCellViewModel mscvm_new = new MatrixSelectionCellViewModel();
                        mscvm_new.Tolerance = "";
                        mscuc_new.DataContext = mscvm_new;
                        mslvm_new.ToleranceViews.Add(mscuc_new);
                    }

                }

                UserControlsInWholeMatrix.Add(msluc_new);
                UserControlsSelectionNames.Add(msnuc_new);
                row_num += 1;
            }

            //string output = "";
            //foreach (Selectionset ss in Selectionsets)
            //{
            //    output += ss.Tag_name + "\n";
            //}
            //MessageBox.Show(output);

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
                    "Система коммутации"
                };
            }
        }
        private string GetSameCategory(string input_category)
        {
            foreach (string category in Categories_in_revit)
            {
                if (category.ToLower().Contains(input_category.ToLower())) return category;
            }
            return "Стены";
        }
    }
}
