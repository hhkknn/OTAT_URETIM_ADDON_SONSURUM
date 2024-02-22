using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class Reports
    {
        [ItemAtt(AIFConn.ReportsXMLUID)]
        public SAPbouiCOM.Form frmReports;

        [ItemAtt("Item_4")]
        public SAPbouiCOM.Grid oGrid;
        [ItemAtt("Item_5")]
        public SAPbouiCOM.EditText edtStartDate;
        [ItemAtt("Item_3")]
        public SAPbouiCOM.EditText edtEndDate;
        [ItemAtt("Item_1")]
        public SAPbouiCOM.CheckBox checkUrunAgaclariDahil;

        SAPbouiCOM.DataTable oGridDataTable = null;
        SAPbouiCOM.DataTable oTempGridDataTable = null;
        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.ReportsXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.ReportsXML));
            Functions.CreateUserOrSystemFormComponent<Reports>(AIFConn.Reports);

            InitForms();
        }
        public void InitForms()
        {
            try
            {
                oGridDataTable = (SAPbouiCOM.DataTable)frmReports.DataSources.DataTables.Item("DT_0");
                oTempGridDataTable = (SAPbouiCOM.DataTable)frmReports.DataSources.DataTables.Add("DT_1");
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

        string val = "";
        string Dt0XML = "<?xmlversion=\"1.0\" encoding=\"UTF-16\" ?><DataTableUid=\"DT_0\"><Columns>{0}</Columns><Rows>{1}</Rows></DataTable>";
        string Dt1XML = "<?xmlversion=\"1.0\" encoding=\"UTF-16\" ?><DataTableUid=\"DT_1\"><Columns>{0}</Columns><Rows>{1}</Rows></DataTable>";
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
                    if (pVal.ItemUID == "Item_6" && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.ProgressBar oProgressBar = null;
                        try
                        {
                            frmReports.Freeze(true);
                            string date1 = edtStartDate.Value;
                            string date2 = edtEndDate.Value;

                            string qry = "EXEC [GunlukStok] '" + date1 + "', '" + date2 + "'";


                            oGridDataTable.Clear();
                            oGridDataTable.ExecuteQuery(qry);


                            DateTime startdate = new DateTime(Convert.ToInt32(date1.Substring(0, 4)), Convert.ToInt32(date1.Substring(4, 2)), Convert.ToInt32(date1.Substring(6, 2)));
                            DateTime endDate = new DateTime(Convert.ToInt32(date2.Substring(0, 4)), Convert.ToInt32(date2.Substring(4, 2)), Convert.ToInt32(date2.Substring(6, 2)));

                            TimeSpan t = endDate - startdate;
                            int diff = Convert.ToInt32(t.TotalDays);
                            double UrunAgaciToplam = 0;
                            string urunKodu = "";
                            string DueDate = "";
                            double OnHand = 0;
                            int colnum = 3;
                            ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            int color = ColorTranslator.ToOle(System.Drawing.Color.Red);
                            oProgressBar = Handler.SAPApplication.StatusBar.CreateProgressBar("Rapor getiriliyor...", oGrid.DataTable.Rows.Count, true);
                            int Progress = 0;
                            oProgressBar.Text = "Rapor getiriliyor...";
                            for (int i = 0; i <= oGrid.DataTable.Rows.Count - 1; i++)
                            {
                                OnHand = parservalues<double>(oGrid.DataTable.GetValue("Stokta", i).ToString());
                                urunKodu = oGrid.DataTable.GetValue("ÜrünKodu", i).ToString();
                                for (int z = 1; z <= diff; z++)
                                {
                                    oGrid.CommonSetting.SetCellFontColor(i + 1, colnum + 1, -1);

                                    if (checkUrunAgaclariDahil.Checked)
                                    {
                                        DueDate = oGrid.Columns.Item(colnum).TitleObject.Caption;

                                        ConstVariables.oRecordset.DoQuery("Select SUM(PlannedQty) from OWOR as T0 where T0.ItemCode = '" + urunKodu + "' and T0.DueDate = '" + DueDate + "' and T0.Status='P'");

                                        UrunAgaciToplam = parservalues<double>(ConstVariables.oRecordset.Fields.Item(0).Value.ToString());
                                    }

                                    double tempval = parservalues<double>(oGrid.DataTable.GetValue(colnum, i).ToString());
                                    double result = OnHand - tempval + UrunAgaciToplam;
                                    oGrid.DataTable.SetValue(colnum, i, result);

                                    if (result < 0)
                                    {
                                        oGrid.CommonSetting.SetCellFontColor(i + 1, colnum + 1, color);
                                    }


                                    OnHand = OnHand - tempval + UrunAgaciToplam;
                                    colnum++;
                                }
                                Progress += 1;
                                oProgressBar.Value = Progress;
                                colnum = 3;
                            }

                            SAPbouiCOM.EditTextColumn oEditCol;

                            oEditCol = ((SAPbouiCOM.EditTextColumn)(oGrid.Columns.Item("ÜrünKodu")));

                            oEditCol.LinkedObjectType = "4";

                            //var xml = oGridDataTable.SerializeAsXML(BoDataTableXmlSelect.dxs_All);





                            //var rows = (from x in XDocument.Parse(xml).Descendants("Columns")
                            //            select new
                            //            {
                            //                uid = (from y in x.Elements("Column") 

                            //            }).ToList();


                            //var columns = from x in XDocument.Parse(xml).Descendants("Columns") select x;

                            //foreach (var item in columns)
                            //{
                            //    var rrr = item.Element("Column");
                            //    var zaaasd = rrr.Attribute("Uid").Value;
                            //}
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            oProgressBar.Stop();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar);
                            GC.Collect();
                            frmReports.Freeze(false);
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

    }
}
