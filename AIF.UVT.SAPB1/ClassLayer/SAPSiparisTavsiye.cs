using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.UVT.SAPB1.HelperClass;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class SAPSiparisTavsiye
    {
        [ItemAtt(AIFConn.SAPSiparisTavsiye_FormUID)]
        public SAPbouiCOM.Form frmSAPSiparisTavsiye;

        private static string formuid = "";
        //private SAPbouiCOM.Button oBtnKaliteKnt;
        //public SAPbouiCOM.Button btnIptal;

        [ItemAtt("3")]
        public SAPbouiCOM.Matrix oMatrix;
        public void LoadForms()
        {
            Functions.CreateUserOrSystemFormComponent<SAPSiparisTavsiye>(AIFConn.Sys65217, true, formuid);

            System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
            System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("AIF.UVT.SAPB1.FormsView.SAPSiparisTavsiye.xml");

            System.IO.StreamReader streamreader = new System.IO.StreamReader(stream, true);
            xmldoc.LoadXml(string.Format(streamreader.ReadToEnd(), formuid));
            Handler.SAPApplication.LoadBatchActions(xmldoc.InnerXml);

            streamreader.Close();

            var cml = frmSAPSiparisTavsiye.GetAsXML();
            InitForms();
        }

        public void InitForms()
        {
            try
            {
                #region Miktar Güncelle buton yerleşimi

                frmSAPSiparisTavsiye.Items.Item("Item_0").Top = frmSAPSiparisTavsiye.Items.Item("2").Top;              
                frmSAPSiparisTavsiye.Items.Item("Item_0").Left = frmSAPSiparisTavsiye.Items.Item("2").Left + frmSAPSiparisTavsiye.Items.Item("2").Width + 5;
                frmSAPSiparisTavsiye.Items.Item("Item_0").LinkTo = "2";

                #endregion
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Form yüklenirken oluştu." + ex.Message);
            }
        }
        bool ekleme = false;
        string SASGercekDocEntry = "";
        string SASTaslakDocEntry = "";
        public bool SAP_FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            BubbleEvent = true;

            switch (BusinessObjectInfo.EventType)
            {
                case BoEventTypes.et_ALL_EVENTS:
                    break;

                case BoEventTypes.et_ITEM_PRESSED:
                    break;

                case BoEventTypes.et_KEY_DOWN:
                    break;

                case BoEventTypes.et_GOT_FOCUS:
                    break;

                case BoEventTypes.et_LOST_FOCUS:
                    break;

                case BoEventTypes.et_COMBO_SELECT:
                    break;

                case BoEventTypes.et_CLICK:
                    break;

                case BoEventTypes.et_DOUBLE_CLICK:
                    break;

                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                    break;

                case BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
                    break;

                case BoEventTypes.et_VALIDATE:
                    break;

                case BoEventTypes.et_MATRIX_LOAD:
                    break;

                case BoEventTypes.et_DATASOURCE_LOAD:
                    break;

                case BoEventTypes.et_FORM_LOAD:
                    break;

                case BoEventTypes.et_FORM_UNLOAD:
                    break;

                case BoEventTypes.et_FORM_ACTIVATE:
                    break;

                case BoEventTypes.et_FORM_DEACTIVATE:
                    break;

                case BoEventTypes.et_FORM_CLOSE:
                    break;

                case BoEventTypes.et_FORM_RESIZE:
                    break;

                case BoEventTypes.et_FORM_KEY_DOWN:
                    break;

                case BoEventTypes.et_FORM_MENU_HILIGHT:
                    break;

                case BoEventTypes.et_PRINT:
                    break;

                case BoEventTypes.et_PRINT_DATA:
                    break;

                case BoEventTypes.et_EDIT_REPORT:
                    break;

                case BoEventTypes.et_CHOOSE_FROM_LIST:
                    break;

                case BoEventTypes.et_RIGHT_CLICK:
                    break;

                case BoEventTypes.et_MENU_CLICK:
                    break;

                case BoEventTypes.et_FORM_DATA_ADD:

                    break;

                case BoEventTypes.et_FORM_DATA_UPDATE:

                    break;

                case BoEventTypes.et_FORM_DATA_DELETE:
                    break;

                case BoEventTypes.et_FORM_DATA_LOAD:
                    break;

                case BoEventTypes.et_PICKER_CLICKED:
                    break;

                case BoEventTypes.et_GRID_SORT:
                    break;

                case BoEventTypes.et_Drag:
                    break;

                case BoEventTypes.et_FORM_DRAW:
                    break;

                case BoEventTypes.et_UDO_FORM_BUILD:
                    break;

                case BoEventTypes.et_UDO_FORM_OPEN:
                    break;

                case BoEventTypes.et_B1I_SERVICE_COMPLETE:
                    break;

                case BoEventTypes.et_FORMAT_SEARCH_COMPLETED:
                    break;

                case BoEventTypes.et_PRINT_LAYOUT_KEY:
                    break;

                case BoEventTypes.et_FORM_VISIBLE:
                    break;

                case BoEventTypes.et_ITEM_WEBMESSAGE:
                    break;

                default:
                    break;
            }

            return BubbleEvent;
        }


        public bool SAP_ItemEvent(string FormUID, ref ItemEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;

            switch (pVal.EventType)
            {
                case BoEventTypes.et_ALL_EVENTS:
                    break;

                case BoEventTypes.et_ITEM_PRESSED:
                    break;

                case BoEventTypes.et_KEY_DOWN:
                    break;

                case BoEventTypes.et_GOT_FOCUS:
                    break;

                case BoEventTypes.et_LOST_FOCUS:
                    break;

                case BoEventTypes.et_COMBO_SELECT:
                    break;

                case BoEventTypes.et_CLICK:
                    if (pVal.ItemUID == "Item_0" && pVal.BeforeAction)
                    {
                        string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                        var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                    select new
                                    {
                                        satirSira= (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "0" select new XElement(y.Element("Value"))).First().Value,
                                        Olustur = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "2" select new XElement(y.Element("Value"))).First().Value,
                                        SiparisTuru = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "8" select new XElement(y.Element("Value"))).First().Value,
                                        KalemNumarasi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "7" select new XElement(y.Element("Value"))).First().Value,
                                        Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "5" select new XElement(y.Element("Value"))).First().Value),

                                    }).ToList();

                        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        oRS = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string TavsiyeParamQry = " select \"U_Sec\",\"U_SipTur\" from \"@AIF_TAVSIYEPARAM\" ";
                        oRS.DoQuery(TavsiyeParamQry);

                        string xmll = oRS.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                        XDocument xDoc = XDocument.Parse(xmll);
                        XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                        var rowsRecordSet = (from t in xDoc.Descendants(ns + "Row")
                                             select new
                                             {
                                                 secili = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Sec" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                 SiparisTuru = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SipTur" select new XElement(y.Element(ns + "Value"))).First().Value,

                                             }).ToList();

                        if (rowsRecordSet.Where(x=>x.secili=="Y").Count() > 0)
                        {
                            foreach (var item in rowsRecordSet)
                            {

                            }
                        }
                    }
                    break;

                case BoEventTypes.et_DOUBLE_CLICK:
                    break;

                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                    break;

                case BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
                    break;

                case BoEventTypes.et_VALIDATE:
                    break;

                case BoEventTypes.et_MATRIX_LOAD:
                    break;

                case BoEventTypes.et_DATASOURCE_LOAD:
                    break;

                case BoEventTypes.et_FORM_LOAD:
                    break;

                case BoEventTypes.et_FORM_UNLOAD:
                    break;

                case BoEventTypes.et_FORM_ACTIVATE:
                    break;

                case BoEventTypes.et_FORM_DEACTIVATE:
                    break;

                case BoEventTypes.et_FORM_CLOSE:
                    break;

                case BoEventTypes.et_FORM_RESIZE:
                    break;

                case BoEventTypes.et_FORM_KEY_DOWN:
                    break;

                case BoEventTypes.et_FORM_MENU_HILIGHT:
                    break;

                case BoEventTypes.et_PRINT:
                    break;

                case BoEventTypes.et_PRINT_DATA:
                    break;

                case BoEventTypes.et_EDIT_REPORT:
                    break;

                case BoEventTypes.et_CHOOSE_FROM_LIST:
                    break;

                case BoEventTypes.et_RIGHT_CLICK:
                    break;

                case BoEventTypes.et_MENU_CLICK:
                    break;

                case BoEventTypes.et_FORM_DATA_ADD:
                    break;

                case BoEventTypes.et_FORM_DATA_UPDATE:
                    break;

                case BoEventTypes.et_FORM_DATA_DELETE:
                    break;

                case BoEventTypes.et_FORM_DATA_LOAD:
                    break;

                case BoEventTypes.et_PICKER_CLICKED:
                    break;

                case BoEventTypes.et_GRID_SORT:
                    break;

                case BoEventTypes.et_Drag:
                    break;

                case BoEventTypes.et_FORM_DRAW:
                    break;

                case BoEventTypes.et_UDO_FORM_BUILD:
                    break;

                case BoEventTypes.et_UDO_FORM_OPEN:
                    break;

                case BoEventTypes.et_B1I_SERVICE_COMPLETE:
                    break;

                case BoEventTypes.et_FORMAT_SEARCH_COMPLETED:
                    break;

                case BoEventTypes.et_PRINT_LAYOUT_KEY:
                    break;

                case BoEventTypes.et_FORM_VISIBLE:
                    break;

                case BoEventTypes.et_ITEM_WEBMESSAGE:
                    break;

                default:
                    break;
            }

            return BubbleEvent;
        }

        public void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }
    }
}
