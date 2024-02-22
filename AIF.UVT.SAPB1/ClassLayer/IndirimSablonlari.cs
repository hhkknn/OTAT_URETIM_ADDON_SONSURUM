using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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
    public class IndirimSablonlari
    {
        [ItemAtt(AIFConn.IndirimSablonUID)]
        public SAPbouiCOM.Form frmIndirimSablonu;

        [ItemAtt("Item_11")]
        public SAPbouiCOM.Matrix oMatrix;
        [ItemAtt("Item_7")]
        public SAPbouiCOM.EditText EdtPriceBefDi;
        [ItemAtt("Item_8")]
        public SAPbouiCOM.EditText EdtTotalDisc;
        [ItemAtt("Item_9")]
        public SAPbouiCOM.EditText EdtDiscRate;
        [ItemAtt("Item_10")]
        public SAPbouiCOM.EditText EdtPriceAfterDisc;
        [ItemAtt("Item_3")]
        public SAPbouiCOM.EditText EdtDocEntry;
        [ItemAtt("Item_13")]
        public SAPbouiCOM.EditText EdtCode;
        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.IndirimSablonXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.IndirimSablonXML));
            Functions.CreateUserOrSystemFormComponent<IndirimSablonlari>(AIFConn.IndirimTmp);
            InitForms();
        }
        public void InitForms()
        {
            try
            {
                SAPbouiCOM.Column oColumn;

                oColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item("Col_3");

                oColumn.ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                oMatrix.AutoResizeColumns();

                EdtPriceBefDi.Value = "100";
                //obaseMatrix = (SAPbouiCOM.Matrix)baseForm.Items.Item(matrixuid).Specific;
                //EdtPriceBefDi.Value = parservalues<double>(((SAPbouiCOM.EditText)obaseMatrix.Columns.Item("14").Cells.Item(line).Specific).Value).ToString();

                //string discountNo = ((SAPbouiCOM.EditText)obaseMatrix.Columns.Item("U_AIF_DiscountNo").Cells.Item(line).Specific).Value;


                //if (discountNo != "")
                //{
                //    try
                //    {
                //        frmIndirimSihirbazi.Freeze(true);
                //        findmode = true;
                //        frmIndirimSihirbazi.Mode = BoFormMode.fm_FIND_MODE;
                //        EdtDocEntry.Value = discountNo;
                //        frmIndirimSihirbazi.Items.Item("1").Click();
                //    }
                //    catch (Exception)
                //    { 
                //    }
                //    finally
                //    { 
                //        frmIndirimSihirbazi.Freeze(false);
                //    }
                //}
                //findmode = false;
                //Hesapla();

                //frmAktivteParametre.EnableMenu("1283", false);
                //frmAktivteParametre.EnableMenu("1284", false);
                //frmAktivteParametre.EnableMenu("1286", false);


                //ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                //ConstVariables.oRecordset.DoQuery("Select \"DocEntry\" from \"@AIF_UVT_ACTPARAM\"");

                //if (ConstVariables.oRecordset.RecordCount > 0)
                //{
                //    frmAktivteParametre.Mode = BoFormMode.fm_FIND_MODE;
                //    EdtDocEntry.Value = "1";
                //    btnAddOrUpdate.Item.Click();

                //}
            }
            catch (Exception ex)
            {
            }
        }
        bool findmode = false;
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
                    if (!BusinessObjectInfo.BeforeAction)
                    {
                    }
                    break;
                case BoEventTypes.et_FORM_DATA_UPDATE:
                    if (!BusinessObjectInfo.BeforeAction)
                    {

                    }
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
        public bool SAP_ItemEvent(string FormUID, ref ItemEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;

            switch (pVal.EventType)
            {
                case BoEventTypes.et_ALL_EVENTS:
                    break;
                case BoEventTypes.et_ITEM_PRESSED:
                    if (pVal.ItemUID == "1" && !pVal.BeforeAction && findmode == false)
                    {
                    }
                    break;
                case BoEventTypes.et_KEY_DOWN:
                    break;
                case BoEventTypes.et_GOT_FOCUS:
                    break;
                case BoEventTypes.et_LOST_FOCUS:
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    if (pVal.ColUID == "Col_1" && pVal.ItemUID == "Item_11" && !pVal.BeforeAction)
                    {
                    }
                    break;
                case BoEventTypes.et_CLICK:
                    //if (pVal.BeforeAction && pVal.ItemUID == "1")
                    //{
                    //    Hesapla();
                    //    string xml = frmIndirimSihirbazi.DataSources.DBDataSources.Item(0).GetAsXML();

                    //    string xml1 = frmIndirimSihirbazi.DataSources.DBDataSources.Item(1).GetAsXML();


                    //}
                    break;
                case BoEventTypes.et_DOUBLE_CLICK:
                    break;
                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                    break;
                case BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
                    break;
                case BoEventTypes.et_VALIDATE:
                    if (pVal.ColUID == "Col_2" && pVal.ItemChanged && !pVal.BeforeAction)
                    {
                        Hesapla();
                    }
                    else if (pVal.ColUID == "Col_3" && pVal.ItemChanged && !pVal.BeforeAction)
                    {
                        Hesapla();
                    }
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
            try
            {
                if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && pVal.BeforeAction)
                {
                    int row = oMatrix.GetNextSelectedRow();
                    if (row != -1)
                    {
                        oMatrix.DeleteRow(row);
                    }
                }
                else if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && pVal.BeforeAction)
                {
                    frmIndirimSablonu.DataSources.DBDataSources.Item(1).Clear();
                    oMatrix.AddRow();

                    if (oMatrix.RowCount == 1)
                    {
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(1).Specific).Value = EdtPriceBefDi.Value == "" ? "0" : EdtPriceBefDi.Value.ToString();
                        calculationLists.Add(new CalculationList { SubTotal = parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(1).Specific).Value) });
                    }
                    else
                    {
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific).Value = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_4").Cells.Item(oMatrix.RowCount - 1).Specific).Value == "" ? "0" : ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_4").Cells.Item(oMatrix.RowCount - 1).Specific).Value;
                        calculationLists.Add(new CalculationList { SubTotal = parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific).Value) });
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true; try
            {
                var oForm = Handler.SAPApplication.Forms.ActiveForm;

                if (eventInfo.ItemUID != "")
                {
                    try
                    {
                        SAPbouiCOM.Matrix item = (SAPbouiCOM.Matrix)oForm.Items.Item(eventInfo.ItemUID).Specific;
                    }
                    catch (Exception)
                    {
                        Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_DeleteRow");
                        return;
                    }


                }
                else
                {
                    Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_DeleteRow");
                    Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_AddRow");
                    return;
                }


                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_DeleteRow");
                    Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_AddRow");
                    return;
                }
                SAPbouiCOM.MenuItem oMenuItem = default(SAPbouiCOM.MenuItem);

                SAPbouiCOM.Menus oMenus = default(SAPbouiCOM.Menus);

                try
                {

                    SAPbouiCOM.MenuCreationParams oCreationPackage = default(SAPbouiCOM.MenuCreationParams);

                    oCreationPackage = (SAPbouiCOM.MenuCreationParams)Handler.SAPApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    try
                    {
                        oCreationPackage.UniqueID = "AIFRGHTCLK_DeleteRow";

                        oCreationPackage.String = "Satır Sil";

                        oCreationPackage.Enabled = true;

                        oMenuItem = Handler.SAPApplication.Menus.Item("1280");

                        oMenus = oMenuItem.SubMenus;

                        oMenus.AddEx(oCreationPackage);

                    }
                    catch
                    {
                    }

                    try
                    {

                        oCreationPackage.UniqueID = "AIFRGHTCLK_AddRow";

                        oCreationPackage.String = "Satır Ekle";

                        oCreationPackage.Enabled = true;

                        oMenuItem = Handler.SAPApplication.Menus.Item("1280");

                        oMenus = oMenuItem.SubMenus;

                        oMenus.AddEx(oCreationPackage);
                    }
                    catch (Exception)
                    {
                    }
                }
                catch (Exception ex)
                {
                }
            }
            catch (Exception ex)
            {

            }
        }

        List<CalculationList> calculationLists = new List<CalculationList>();
        public class CalculationList
        {
            public double SubTotal { get; set; }
            public string DiscType { get; set; }
            public double DiscRate { get; set; }
            public double DiscTotal { get; set; }
            public double SubTotal2 { get; set; }
            public string lineid { get; set; }
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

        private void Hesapla()
        {
            try
            {
                frmIndirimSablonu.Freeze(true);
                string disctype = "";
                double pricebeforeTotal = 0;
                double totalddiscrate = 0;
                double totaldisc = 0;
                double priceaftertotl = 0;
                double subTotal = 0;
                double discrateline = 0;
                double priceline = 0;
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    disctype = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific).Value;

                    discrateline = parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(i).Specific).Value);

                    if (disctype == "1")
                        if (disctype == "1")
                        {
                            subTotal = parservalues<double>(EdtPriceBefDi.Value.ToString());
                        }
                        else
                        {
                            subTotal = parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value);
                        }


                    priceline = (subTotal / 100) * discrateline;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_3").Cells.Item(i).Specific).Value = priceline.ToString();

                    subTotal = parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value);

                    subTotal = subTotal - priceline;


                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_4").Cells.Item(i).Specific).Value = subTotal.ToString();

                    try
                    {

                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i + 1).Specific).Value = subTotal.ToString();
                    }
                    catch (Exception)
                    {
                    }
                }

                string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                            select new CalculationList
                            {
                                SubTotal = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value),
                                DiscType = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                                DiscRate = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value),
                                DiscTotal = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value),
                                SubTotal2 = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value),

                                lineid = (x.ElementsBeforeSelf().Count() + 1).ToString()


                            }).ToList();

                //rows.ForEach(x => x.SubTotal2 = x.SubTotal - x.DiscTotal);
                pricebeforeTotal = parservalues<double>(EdtPriceBefDi.Value);
                totaldisc = rows.Sum(x => x.DiscTotal);
                priceaftertotl = rows.Where(y => y.SubTotal2 > 0).Select(x => x.SubTotal2).LastOrDefault();

                EdtTotalDisc.Value = parservalues<double>(totaldisc.ToString()).ToString();
                EdtPriceAfterDisc.Value = parservalues<double>(priceaftertotl.ToString()).ToString();
                EdtDiscRate.Value = ((totaldisc / pricebeforeTotal) * 100).ToString();

                #region XML ile kayıt
                //return;
                //frmIndirimSihirbazi.Freeze(true);
                //string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                //var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                //            select new CalculationList
                //            {
                //                SubTotal = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value),
                //                DiscType = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                //                DiscRate = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value),





                //                DiscTotal = ((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value == "1" || (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value == "2") ?
                //                 parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value) / 100 * parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value) : parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value),

                //                lineid = (x.ElementsBeforeSelf().Count() + 1).ToString()


                //            }).ToList();

                //rows.ForEach(x => x.SubTotal2 = x.SubTotal - x.DiscTotal);

                //string headerxml = "<?xml version=\"1.0\" encoding=\"UTF-16\" ?><dbDataSources uid=\"@AIF_SALES_DISC\"><rows>{0}</rows></dbDataSources>";

                //#region Masterdata header
                ////string header = "<row><cells><cell><uid>Code</uid><value>{0}</value></cell><cell><uid>DocEntry</uid><value>{1}</value></cell><cell><uid>Canceled</uid><value>{2}</value></cell><cell><uid>Object</uid><value>{3}</value></cell><cell><uid>UserSign</uid><value>{4}</value></cell><cell><uid>Transfered</uid><value>{5}</value></cell><cell><uid>CreateDate</uid><value>{6}</value></cell><cell><uid>CreateTime</uid><value>{7}</value></cell><cell><uid>UpdateDate</uid><value>{8}</value></cell><cell><uid>UpdateTime</uid><value>{9}</value></cell><cell><uid>DataSource</uid><value>{10}</value></cell><cell><uid>U_PriceBefDisc</uid><value>{11}</value></cell><cell><uid>U_TotalDisc</uid><value>{12}</value></cell><cell><uid>U_DiscRate</uid><value>{13}</value></cell><cell><uid>U_PriceAfterDisc</uid><value>{14}</value></cell></cells></row>"; 
                //#endregion

                //string header = "<row><cells><cell><uid>DocEntry</uid><value>{0}</value></cell><cell><uid>U_PriceBefDisc</uid><value>{1}</value></cell><cell><uid>U_TotalDisc</uid><value>{2}</value></cell><cell><uid>U_DiscRate</uid><value>{3}</value></cell><cell><uid>U_PriceAfterDisc</uid><value>{4}</value></cell></cells></row>";

                //double pricebeforeTotal = parservalues<double>(EdtPriceBefDi.Value);
                //double totalddiscrate = rows.Sum(x => x.DiscRate);
                //double totaldisc = rows.Sum(x => x.DiscTotal);
                //double priceaftertotl = rows.Where(y => y.SubTotal2 > 0).Select(x => x.SubTotal2).LastOrDefault();

                //EdtTotalDisc.Value = parservalues<double>(totaldisc.ToString()).ToString();
                //EdtPriceAfterDisc.Value = parservalues<double>(priceaftertotl.ToString()).ToString();
                //EdtDiscRate.Value = ((totaldisc / pricebeforeTotal) * 100).ToString();

                //#region Master Data Heeader Data
                ////string headerdata = string.Join("", string.Format(header, EdtCode.Value, EdtDocEntry.Value, "N", "AIF_SALES_DISC", "1", "N", "20200427", "1800", "20200427", "1900", "I", EdtPriceBefDi.Value, EdtTotalDisc.Value, EdtDiscRate.Value, EdtPriceAfterDisc.Value)); 
                //#endregion
                //string headerdata = string.Join("", string.Format(header, EdtDocEntry.Value, EdtPriceBefDi.Value, EdtTotalDisc.Value, EdtDiscRate.Value, EdtPriceAfterDisc.Value));

                //string headerfinal = string.Format(headerxml, headerdata);

                //string xmlss = frmIndirimSihirbazi.DataSources.DBDataSources.Item(1).GetAsXML();

                //string rowsheader = "<?xml version=\"1.0\" encoding=\"UTF-16\" ?><dbDataSources uid=\"@AIF_SALES_DISC1\"><rows>{0}</rows></dbDataSources>";

                //#region Master data row
                ////string row = "<row><cells><cell><uid>U_SubTotal</uid><value>{0}</value></cell><cell><uid>U_DiscType</uid><value>{1}</value></cell><cell><uid>U_DiscRate</uid><value>{2}</value></cell><cell><uid>U_DiscTotal</uid><value>{3}</value></cell><cell><uid>U_SubTotal2</uid><value>{4}</value></cell><cell><uid>LineId</uid><value>{5}</value></cell><cell><uid>Code</uid><value>{6}</value></cell><cell><uid>Object</uid><value>{7}</value></cell></cells></row>"; 
                //#endregion

                //string row = "<row><cells><cell><uid>U_SubTotal</uid><value>{0}</value></cell><cell><uid>U_DiscType</uid><value>{1}</value></cell><cell><uid>U_DiscRate</uid><value>{2}</value></cell><cell><uid>U_DiscTotal</uid><value>{3}</value></cell><cell><uid>U_SubTotal2</uid><value>{4}</value></cell><cell><uid>LineId</uid><value>{5}</value></cell><cell><uid>DocEntry</uid><value>{6}</value></cell><cell><uid>Object</uid><value>{7}</value></cell></cells></row>";

                //string data = string.Join("", rows.Select(s => string.Format(row, s.SubTotal, s.DiscType, s.DiscRate, s.DiscTotal, s.SubTotal2, s.lineid, EdtCode.Value, "AIF_SALES_DISC")));

                //string final = string.Format(rowsheader, data);

                ////frmIndirimSihirbazi.DataSources.DBDataSources.Item("@AIF_SALES_DISC").Clear();
                ////frmIndirimSihirbazi.DataSources.DBDataSources.Item("@AIF_SALES_DISC1").Clear();


                //XmlDocument doc = new XmlDocument();
                //doc.LoadXml(headerfinal);

                //XmlDocument doc1 = new XmlDocument();
                //doc1.LoadXml(final);

                //frmIndirimSihirbazi.DataSources.DBDataSources.Item("@AIF_SALES_DISC").LoadFromXML(doc.InnerXml);
                //frmIndirimSihirbazi.DataSources.DBDataSources.Item("@AIF_SALES_DISC1").LoadFromXML(doc1.InnerXml);

                //oMatrix.LoadFromDataSource();
                //oMatrix.AutoResizeColumns(); 
                #endregion

            }
            catch (Exception ex)
            {

            }
            finally
            {
                frmIndirimSablonu.Freeze(false);
            }

        }
    }
}
