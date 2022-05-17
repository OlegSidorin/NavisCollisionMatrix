using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Xml;
using System.Xml.Serialization;

namespace СollisionMatrix
{
    public class MatrixCreatingViewModel : ObservableObject
    {
        public ObservableCollection<MatrixSelectionLineModel> Selections { get; set; }
        public ObservableCollection<UserControl> UserControlsInWholeMatrix { get; set; }
        public ObservableCollection<UserControl> UserControlsSelectionNames { get; set; }
        public MatrixCreatingViewModel()
        {
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

            xDoc.Save(@"C:\Users\o.sidorin\Downloads\examplexml.xml");

        }
        private bool CanDoIfIClickOnCreateXMLCollisionMatrixButtonExecute(object p) => true;

    }
}
