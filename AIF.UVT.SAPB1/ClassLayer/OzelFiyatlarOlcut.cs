using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class OzelFiyatlarOlcut
    {
        [ItemAtt(AIFConn.OzelFiyatlar_FormUID)]
        public SAPbouiCOM.Form frmOzelFiyatlarOlcut;

        public SAPbouiCOM.EditText edtTemplateCode;
        //[ItemAtt("1")]
        //public SAPbouiCOM.Button btnAddOrUpdate;
        public static SAPbouiCOM.Form baseForm;
        static string formuid = "";
        public void LoadForms()
        {
            if (Program.mKod == "010OTATURVT")
            {
                Functions.CreateUserOrSystemFormComponent<OzelFiyatlarOlcut>(AIFConn.Sys669, true, formuid);

                System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
                System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("AIF.UVT.SAPB1.FormsView.OzelFiyatlarSecim.xml");

                System.IO.StreamReader streamreader = new System.IO.StreamReader(stream, true);
                xmldoc.LoadXml(string.Format(streamreader.ReadToEnd(), formuid));
                Handler.SAPApplication.LoadBatchActions(xmldoc.InnerXml);

                //streamreader.Close(); 
                InitForms();
            }
        }
        public void InitForms()
        {
            try
            {
                edtTemplateCode = (SAPbouiCOM.EditText)frmOzelFiyatlarOlcut.Items.Item("Item_1").Specific;
            }
            catch (Exception ex)
            {
            }
        }

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

        bool isok = false;
        double totalDiscRate = 0;
        string val = "";
        public bool SAP_ItemEvent(string FormUID, ref ItemEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;

            switch (pVal.EventType)
            {
                case BoEventTypes.et_ALL_EVENTS:
                    break;
                case BoEventTypes.et_ITEM_PRESSED:
                    //if (pVal.BeforeAction && pVal.ItemUID == "1")
                    //{
                    //    string templateCode = cmbTemplateCode.Value.Trim();
                    //    SAPbouiCOM.Matrix oMatrixMuhatapOzelFiyatlar = (SAPbouiCOM.Matrix)baseForm.Items.Item("14").Specific;

                    //    for (int i = 1; i <= oMatrixMuhatapOzelFiyatlar.RowCount; i++)
                    //    {
                    //        ((SAPbouiCOM.EditText)oMatrixMuhatapOzelFiyatlar.Columns.Item("U_TemplateCode").Cells.Item(i).Specific).Value = templateCode;
                    //    }
                    //}
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
                    if (Program.mKod == "010OTATURVT")
                    {
                        if (pVal.ItemUID == "1")
                        {
                            isok = true;
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
                    if (pVal.BeforeAction)
                    {
                        frmOzelFiyatlarOlcut = Handler.SAPApplication.Forms.Item(pVal.FormUID);
                        //frmInit = false; 
                        formuid = pVal.FormUID;
                        AIFConn.Sys669.LoadForms();
                    }
                    break;
                case BoEventTypes.et_FORM_UNLOAD:
                    break;
                case BoEventTypes.et_FORM_ACTIVATE:
                    break;
                case BoEventTypes.et_FORM_DEACTIVATE:
                    break;
                case BoEventTypes.et_FORM_CLOSE:
                    if (Program.mKod == "010OTATURVT")
                    {
                        if (!pVal.BeforeAction && isok)
                        {
                            SAPbouiCOM.ProgressBar oProgressBar;
                            SAPbouiCOM.Matrix oMatrixMuhatapOzelFiyatlar = (SAPbouiCOM.Matrix)baseForm.Items.Item("14").Specific;
                            var xml = oMatrixMuhatapOzelFiyatlar.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                            var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                        select new
                                        {
                                            TemplateCode = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "U_TemplateCode" select new XElement(y.Element("Value"))).First().Value,
                                            row = x.ElementsBeforeSelf().Count() + 1
                                        }).ToList();

                            oProgressBar = Handler.SAPApplication.StatusBar.CreateProgressBar("Şablon Kodlar Aktarılıyor...", rows.Where(x => x.TemplateCode != "").Count(), true);
                            try
                            {

                                int Progress = 0;
                                oProgressBar.Text = "Şablon Kodlar Aktarılıyor...";

                                baseForm.Freeze(true);
                                string templateCode = edtTemplateCode.Value;

                                //for (int i = 1; i <= oMatrixMuhatapOzelFiyatlar.RowCount - 1; i++)
                                //{
                                foreach (var item in rows.Where(x => x.TemplateCode == ""))
                                {
                                    ((SAPbouiCOM.EditText)oMatrixMuhatapOzelFiyatlar.Columns.Item("U_TemplateCode").Cells.Item(item.row).Specific).Value = templateCode;
                                    ((SAPbouiCOM.EditText)oMatrixMuhatapOzelFiyatlar.Columns.Item("4").Cells.Item(item.row).Specific).Value = totalDiscRate.ToString();
                                    Progress += 1;
                                    oProgressBar.Value = Progress;
                                }
                                //}
                                oMatrixMuhatapOzelFiyatlar.Columns.Item("1").Cells.Item(1).Click();
                            }
                            catch (Exception)
                            {
                            }
                            finally
                            {
                                oProgressBar.Stop();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar);
                                GC.Collect();
                                baseForm.Freeze(false);
                            }
                            isok = false;
                        }
                    }
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
                    if (Program.mKod == "010OTATURVT")
                    {
                        if (pVal.ItemUID == "Item_1" && !pVal.BeforeAction)
                        {
                            try
                            {
                                var xml = Handler.SAPApplication.Forms.ActiveForm.GetAsXML();
                                SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                                string DocEntry = "";
                                try
                                {
                                    DocEntry = oDataTable.GetValue("DocEntry", 0).ToString();
                                }
                                catch (Exception)
                                {
                                }
                                try
                                {
                                    edtTemplateCode.Value = DocEntry;
                                }
                                catch (Exception)
                                {
                                }
                                totalDiscRate = parservalues<double>(oDataTable.GetValue("U_DiscRate", 0).ToString());
                            }
                            catch (Exception)
                            {
                            }

                        }
                    }
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

        private static CultureInfo _parsCult;
        public static T parservalues<T>(string val) where T : struct
        {

            object returnvals = 0;
            try
            {
                if (string.IsNullOrEmpty(val)) return (T)Convert.ChangeType(returnvals, typeof(T));

                val = Regex.Replace(val, @"[^0-9.,]+", "").Trim();

                CultureInfo parser = parsCult;
                if (new Regex("^(?!0|\\.00)[0-9]+(,\\d{3})*([.][0-9]{0,9})$").IsMatch(val))
                {
                    parser = System.Globalization.CultureInfo.InvariantCulture;
                }

                if (typeof(T) == typeof(decimal))
                {
                    returnvals = decimal.Parse(val, parser);
                }
                else if (typeof(T) == typeof(double))
                {
                    returnvals = double.Parse(val, parser);
                }

                else if (typeof(T) == typeof(float))
                {
                    returnvals = float.Parse(val, parser);
                }
                else if (typeof(T) == typeof(int))
                {
                    returnvals = (int)decimal.Parse(val, parser);
                }
            }
            catch (Exception ex)
            {
            }

            return (T)returnvals;

        }

        private static CultureInfo parsCult
        {
            get
            {
                if (_parsCult == null)
                {
                    CultureInfo ci = CultureInfo.InvariantCulture;
                    _parsCult = (CultureInfo)ci.Clone();

                    _parsCult.NumberFormat.CurrencyDecimalSeparator = ",";
                    _parsCult.NumberFormat.NumberDecimalSeparator = ",";
                    _parsCult.NumberFormat.PercentDecimalSeparator = ",";
                    _parsCult.NumberFormat.CurrencyGroupSeparator = ".";
                    _parsCult.NumberFormat.NumberGroupSeparator = ".";
                    _parsCult.NumberFormat.PercentGroupSeparator = ".";
                }

                return _parsCult;

            }
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
