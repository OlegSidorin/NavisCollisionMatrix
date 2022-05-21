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
        public ObservableCollection<UserControl> UserControlsSelectionNames { get; set; }
        public List<Selectionset> Selectionsets { get; set; }
        public List<Clashtest> Clashtests { get; set; }
        public MatrixCreatingViewModel()
        {
            //ClashTests = new ObservableCollection<ClashTest>();
            Selectionsets = new List<Selectionset>();
            Clashtests = new List<Clashtest>();

            Selections = new ObservableCollection<MatrixSelectionLineModel>();
            var new1 = new MatrixSelectionLineModel()
            {
                NameOfSelection = "AR_Walls",
                SelectionIntersectionTolerance = new List<string>() 
                {
                    "15", "", "", ""
                }
            };
            var new2 = new MatrixSelectionLineModel()
            {
                NameOfSelection = "AR_Floors",
                SelectionIntersectionTolerance = new List<string>()
                {
                    "", "30", "", ""
                }
            };
            var new3 = new MatrixSelectionLineModel()
            {
                NameOfSelection = "AR_Doors",
                SelectionIntersectionTolerance = new List<string>()
                {
                    "", "", "80", ""
                }
            };
            var new4 = new MatrixSelectionLineModel()
            {
                NameOfSelection = "AR_Windows",
                SelectionIntersectionTolerance = new List<string>()
                {
                    "", "", "", "50"
                }
            };

            Selections.Add(new1);
            Selections.Add(new2);
            Selections.Add(new3);
            Selections.Add(new4);

            UserControlsInWholeMatrix = new ObservableCollection<UserControl>();
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

                    MatrixSelectionCellVewModel cellViewModel = new MatrixSelectionCellVewModel()
                    {
                        Tolerance = tolerance
                    };
                    MatrixSelectionCellUserControl cellView = new MatrixSelectionCellUserControl();
                    cellView.DataContext = cellViewModel;

                    userControlvm.ToleranceViews.Add(cellView);
                }

                MatrixSelectionLineUserControl userControl = new MatrixSelectionLineUserControl();
                userControlvm.UserControl_MatrixSelectionLineUserControl = userControl;
                userControlvm.UserControlsInAllMatrixWithLineUserControls = UserControlsInWholeMatrix;
                userControlvm.UserControlsInSelectionNameUserControls = UserControlsSelectionNames;
                userControl.DataContext = userControlvm;
                UserControlsInWholeMatrix.Add(userControl);
            };

            foreach (UserControl usercontrolline in UserControlsInWholeMatrix)
            {
                MatrixSelectionLineUserControl msluc = (MatrixSelectionLineUserControl)usercontrolline;
                MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;

                MatrixSelectionNameUserControl msnuc = new MatrixSelectionNameUserControl();
                msnuc.DataContext = mslvm;
                UserControlsSelectionNames.Add(msnuc);
            }

            DoIfIClickOnCreateXMLCollisionMatrixButton = new RelayCommand(OnDoIfIClickOnCreateXMLCollisionMatrixButtonExecuted, CanDoIfIClickOnCreateXMLCollisionMatrixButtonExecute);
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
                    MatrixSelectionCellVewModel mscvm = (MatrixSelectionCellVewModel)mscuc.DataContext;

                    cellcounter += 1;
                }
                linecounter += 1;
            }

        }

        public ICommand DoIfIClickOnCreateXMLCollisionMatrixButton { get; set; }
        private void OnDoIfIClickOnCreateXMLCollisionMatrixButtonExecuted(object p)
        {
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
            batchtest_element.SetAttribute("name", "Без имени");
            batchtest_element.SetAttribute("internal_name", "Без имени");
            batchtest_element.SetAttribute("units", "ft");
            XmlNode batchtest_node = exchange_node.AppendChild(batchtest_element);

            XmlElement clashtests_element = xDoc.CreateElement("clashtests");
            XmlNode clashtests_node = batchtest_node.AppendChild(clashtests_element);
            int iForToleranceArray = 0;
            foreach(UserControl usercontrolline in UserControlsInWholeMatrix)
            {
                iForToleranceArray = 0;
                MatrixSelectionLineUserControl msluc = (MatrixSelectionLineUserControl)usercontrolline;
                MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;
                foreach (UserControl usercontrolcell in mslvm.ToleranceViews)
                {
                    MatrixSelectionCellUserControl mscuc = (MatrixSelectionCellUserControl)usercontrolcell;
                    MatrixSelectionCellVewModel mscvm = (MatrixSelectionCellVewModel)mscuc.DataContext;
                    if (mscvm.Tolerance != string.Empty)
                    {
                        bool result = double.TryParse(mscvm.Tolerance, out double dt);
                        if (result)
                        {
                            MatrixSelectionLineUserControl uci = (MatrixSelectionLineUserControl)UserControlsInWholeMatrix.ElementAt(iForToleranceArray);
                            MatrixSelectionLineViewModel vmi = (MatrixSelectionLineViewModel)uci.DataContext;
                            string nameOfOtherSelection = vmi.NameOfSelection;
                            XmlElement clashtest_element = xDoc.CreateElement("clashtest");
                            clashtest_element.SetAttribute("name", mslvm.NameOfSelection + " - " + nameOfOtherSelection);
                            clashtest_element.SetAttribute("test_type", "hard");
                            clashtest_element.SetAttribute("status", "new");
                            clashtest_element.SetAttribute("tolerance", (dt / 304.78).ToString().Replace(",", ".").Substring(0,12));
                            clashtest_element.SetAttribute("merge_composites", "1");
                            XmlNode clashtest_node = clashtests_node.AppendChild(clashtest_element);

                            // <linkage mode="none"/>
                            XmlElement linkage_element = xDoc.CreateElement("linkage");
                            linkage_element.SetAttribute("mode", "none");
                            XmlNode linkage_node = clashtest_node.AppendChild(linkage_element);

                            // <left>
                            XmlElement left_element = xDoc.CreateElement("left");
                            XmlNode left_node = clashtest_node.AppendChild(left_element);

                            // <clashselection selfintersect="0" primtypes="1">
                            XmlElement clashselection_element = xDoc.CreateElement("clashselection");
                            clashselection_element.SetAttribute("selfintersect", "0");
                            clashselection_element.SetAttribute("primtypes", "1");
                            XmlNode clashselection_node = left_node.AppendChild(clashselection_element);

                            // <locator>lcop_selection_set_tree/АР_Стены</locator>
                            XmlElement locator_element = xDoc.CreateElement("locator");
                            locator_element.InnerText = @"lcop_selection_set_tree/" + mslvm.NameOfSelection;
                            XmlNode locator_node = clashselection_node.AppendChild(locator_element);

                            // <right>
                            XmlElement right_element = xDoc.CreateElement("right");
                            XmlNode right_node = clashtest_node.AppendChild(right_element);

                            // <clashselection selfintersect="0" primtypes="1">
                            XmlElement clashselectionR_element = xDoc.CreateElement("clashselection");
                            clashselectionR_element.SetAttribute("selfintersect", "0");
                            clashselectionR_element.SetAttribute("primtypes", "1");
                            XmlNode clashselectionR_node = right_node.AppendChild(clashselectionR_element);

                            // <locator>lcop_selection_set_tree/АР_Стены</locator>
                            XmlElement locatorR_element = xDoc.CreateElement("locator");
                            locatorR_element.InnerText = @"lcop_selection_set_tree/" + nameOfOtherSelection;
                            XmlNode locatorR_node = clashselectionR_node.AppendChild(locatorR_element);

                            // <rules/>
                            XmlElement rules_element = xDoc.CreateElement("rules");
                            XmlNode rules_node = clashtest_node.AppendChild(rules_element);
                        }
                    }
                    iForToleranceArray += 1;
                }

                
            }

            XmlElement selectionsets_element = xDoc.CreateElement("selectionsets");
            XmlNode selectionsets_node = batchtest_node.AppendChild(selectionsets_element);

            foreach (UserControl usercontrolline in UserControlsInWholeMatrix)
            {
                MatrixSelectionLineUserControl msluc = (MatrixSelectionLineUserControl)usercontrolline;
                MatrixSelectionLineViewModel mslvm = (MatrixSelectionLineViewModel)msluc.DataContext;
                XmlElement selectionset_element = xDoc.CreateElement("selectionset");
                selectionset_element.SetAttribute("name", mslvm.NameOfSelection);
                selectionset_element.SetAttribute("guid", "");
                XmlNode selectionset_node = selectionsets_node.AppendChild(selectionset_element);

                // <findspec mode="all" disjoint="0">
                XmlElement findspec_element = xDoc.CreateElement("findspec");
                findspec_element.SetAttribute("mode", "all");
                findspec_element.SetAttribute("disjoint", "0");
                XmlNode findspec_node = selectionset_node.AppendChild(findspec_element);

                // <conditions>
                XmlElement conditions_element = xDoc.CreateElement("conditions");
                XmlNode conditions_node = findspec_node.AppendChild(conditions_element);

                var names = mslvm.NameOfSelection.Split('_');
                string source_file_mask = string.Empty;
                string category_mask = string.Empty;
                if (names.Count() > 1)
                {
                    source_file_mask = names.First();
                    category_mask = names.LastOrDefault();
                }
                else
                {
                    source_file_mask = "AR";
                    category_mask = "Walls";
                }

                //---------------------------------------- condition SourceFile --------------

                // <condition test="contains" flags="10">
                XmlElement condition_SF_element = xDoc.CreateElement("condition");
                condition_SF_element.SetAttribute("test", "contains");
                condition_SF_element.SetAttribute("flags", "10");
                XmlNode condition_SF_node = conditions_node.AppendChild(condition_SF_element);

                // <property>
                XmlElement property_SF_element = xDoc.CreateElement("property");
                XmlNode property_SF_node = condition_SF_node.AppendChild(property_SF_element);
                // <name internal="LcOaNodeSourceFile">Файл источника</name>
                XmlElement name_SF__element_p = xDoc.CreateElement("name");
                name_SF__element_p.SetAttribute("internal", "LcOaNodeSourceFile");
                name_SF__element_p.InnerText = "Файл источника";
                XmlNode name_SF_node_p = property_SF_node.AppendChild(name_SF__element_p);

                // <value>
                XmlElement value_SF_element = xDoc.CreateElement("value");
                XmlNode value_SF_node = condition_SF_node.AppendChild(value_SF_element);
                // <data type="wstring">_АР</data>
                XmlElement data_SF_element = xDoc.CreateElement("data");
                data_SF_element.SetAttribute("type", "wstring");
                data_SF_element.InnerText = source_file_mask + "_";
                XmlNode data_SF_node = value_SF_node.AppendChild(data_SF_element);


                //---------------------------------------- condition ElementCategory ----------

                // <condition test="equals" flags="10">
                XmlElement condition_element = xDoc.CreateElement("condition");
                condition_element.SetAttribute("test", "equals");
                condition_element.SetAttribute("flags", "10");
                XmlNode condition_node = conditions_node.AppendChild(condition_element);

                // <category>
                XmlElement category_element = xDoc.CreateElement("category");
                XmlNode category_node = condition_node.AppendChild(category_element);
                // <name internal="LcRevitData_Element">Объект</name>
                XmlElement name_element = xDoc.CreateElement("name");
                name_element.SetAttribute("internal", "LcRevitData_Element");
                name_element.InnerText = "Объект";
                XmlNode name_node = category_node.AppendChild(name_element);

                // <property>
                XmlElement property_element = xDoc.CreateElement("property");
                XmlNode property_node = condition_node.AppendChild(property_element);
                // <name internal="LcRevitPropertyElementCategory">Категория</name>
                XmlElement name_element_p = xDoc.CreateElement("name");
                name_element_p.SetAttribute("internal", "LcRevitPropertyElementCategory");
                name_element_p.InnerText = "Категория";
                XmlNode name_node_p = property_node.AppendChild(name_element_p);

                // <value>
                XmlElement value_element = xDoc.CreateElement("value");
                XmlNode value_node = condition_node.AppendChild(value_element);
                // <data type="wstring">Стены</data>
                XmlElement data_element = xDoc.CreateElement("data");
                data_element.SetAttribute("type", "wstring");
                data_element.InnerText = GetSameCategory(category_mask);
                XmlNode data_node = value_node.AppendChild(data_element);



                // <locator>/</locator>
                XmlElement locator_element = xDoc.CreateElement("locator");
                locator_element.InnerText = "/";
                XmlNode locator_node = findspec_node.AppendChild(locator_element);
            }

            string pathtoxml = "";

            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            var dialog_result = openFileDialog.ShowDialog();
            pathtoxml = openFileDialog.FileName;

            if (dialog_result == System.Windows.Forms.DialogResult.OK) xDoc.Save(pathtoxml);

        }
        private bool CanDoIfIClickOnCreateXMLCollisionMatrixButtonExecute(object p) => true;


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

            try
            {
                XmlElement exchange_element = xDoc.DocumentElement; // <exchange xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd" units="ft" filename="" filepath="">

                foreach (XmlNode batchtest_node in exchange_element) // <batchtest name="Без имени" internal_name="Без имени" units="ft">
                {
                    if (batchtest_node.Name == "batchtest")
                    {
                        foreach (XmlNode node_in_batchtest_node in batchtest_node.ChildNodes) // <clashtests>
                        {
                            if (node_in_batchtest_node.Name == "clashtests")
                            {
                                foreach (XmlNode clashtest_node in node_in_batchtest_node.ChildNodes) // <clashtest name="АР_Вертикальные конструкции - АР_Вертикальные конструкции" test_type="duplicate" status="ok" tolerance="0.000" merge_composites="0">
                                {
                                    if (clashtest_node.Name == "clashtest")
                                    {
                                        Clashtest clashtest = new Clashtest();
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
                                                                clashtest.Left.Clashselection.Locator.Tag_inner_text_selection_name = locator_node.InnerText.Replace("lcop_selection_set_tree/", "");

                                                                var folders_names = clashtest.Left.Clashselection.Locator.Tag_inner_text_selection_name.Split('/');
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
                                                                clashtest.Right.Clashselection.Locator.Tag_inner_text_selection_name = locator_node.InnerText.Replace("lcop_selection_set_tree/", "");

                                                                var folders_names = clashtest.Right.Clashselection.Locator.Tag_inner_text_selection_name.Split('/');
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
                        MatrixSelectionCellVewModel mscvm_new = new MatrixSelectionCellVewModel();
                        mscvm_new.Tolerance = clashtest.Tag_tolerance_in_mm;
                        mscuc_new.DataContext = mscvm_new;
                        mslvm_new.ToleranceViews.Add(mscuc_new);
                    }
                    else
                    {
                        MatrixSelectionCellUserControl mscuc_new = new MatrixSelectionCellUserControl();
                        MatrixSelectionCellVewModel mscvm_new = new MatrixSelectionCellVewModel();
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

        private static List<string> categories_in_revit = new List<string>()
        {
            "Стены", "Перекрытия", "Панели витража", "Импосты витража", "Потолки", "Окна", "Двери", "Лестницы", "Крыши", "Несущие колонны",
            "Каркас несущий", "Лестницы", "Фундамент несущей конструкции", "Оборудование", 
            "Арматура воздуховодов", "Воздуховоды", "Соединительные детали воздуховодов",
            "Арматура трубопроводов", "Трубы", "Соединительные детали трубопроводов", "Спринклеры",
            "Сантехнические приборы", "Электрооборудование", "Кабельные лотки", "Короба", "Осветительные приборы"
        };
        private static string GetSameCategory(string input_category)
        {
            foreach (string category in categories_in_revit)
            {
                if (category.ToLower().Contains(input_category.ToLower())) return category;
            }
            return "Стены";
        }
    }
}
