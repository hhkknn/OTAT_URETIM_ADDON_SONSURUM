using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using AIF.UVT.SAPB1.HelperClass;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class UretimSiparisi
    {
        //[ItemAtt(AIFConn.UretimSiparisi_FormUID)]
        public SAPbouiCOM.Form frmUretimSiparisi;

        static string formuid = "";
        SAPbouiCOM.Matrix oPartiMatrix = null;
        SAPbouiCOM.EditText edtPlanlananMiktar = null;
        SAPbouiCOM.EditText edtUrunKodu = null;
        public void LoadForms()
        {
            //Functions.CreateUserOrSystemFormComponent<UretimSiparisi>(AIFConn.Sys65211, true, formuid);

            System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
            System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("AIF.UVT.SAPB1.FormsView.UretimSiparisi.xml");

            System.IO.StreamReader streamreader = new System.IO.StreamReader(stream, true);
            xmldoc.LoadXml(string.Format(streamreader.ReadToEnd(), formuid));
            Handler.SAPApplication.LoadBatchActions(xmldoc.InnerXml);

            streamreader.Close();

            //SAPbouiCOM.Folder oPartiFolder = (SAPbouiCOM.Folder)frmUretimSiparisi.Items.Item("Item_0").Specific;
            //oPartiFolder.GroupWith("36");
            //oform.PaneLevel = 1;
            //oPartiFolder.Item.Left = frmUretimSiparisi.Items.Item("36").Left;
            //oPartiMatrix = (SAPbouiCOM.Matrix)frmUretimSiparisi.Items.Item("Item_3").Specific;
            edtPlanlananMiktar = (SAPbouiCOM.EditText)frmUretimSiparisi.Items.Item("12").Specific;
            edtUrunKodu = (SAPbouiCOM.EditText)frmUretimSiparisi.Items.Item("6").Specific;
            var cml = frmUretimSiparisi.GetAsXML();
            InitForms();
        }
        string UrtPrtSekli = "";

        public void InitForms()
        {
            try
            {
                //frmUretimSiparisi.Items.Item("Item_0").Top = fr.Items.Item("2").Top;
                //oEdtCardCode = (SAPbouiCOM.EditText)frmSatisSiparisi.Items.Item("4").Specific;
                string sql = "Select \"U_UrtPrtSekli\" from \"@AIF_UVT_PARAM\"";
                ConstVariables.oRecordset.DoQuery(sql);

                if (ConstVariables.oRecordset.RecordCount > 0)
                {
                    UrtPrtSekli = ConstVariables.oRecordset.Fields.Item("U_UrtPrtSekli").Value.ToString();
                }
                if (UrtPrtSekli == "2")
                {
                    #region üretim siparişi partilendirme buton yerleşimi

                    frmUretimSiparisi.Items.Item("btnUrtPrc").Top = frmUretimSiparisi.Items.Item("2").Top;
                    frmUretimSiparisi.Items.Item("btnUrtPrc").Height = frmUretimSiparisi.Items.Item("2").Height;
                    frmUretimSiparisi.Items.Item("btnUrtPrc").Left = frmUretimSiparisi.Items.Item("2").Left + 70;
                    frmUretimSiparisi.Items.Item("btnUrtPrc").LinkTo = "2";




                    #endregion
                }
                try
                {
                    frmUretimSiparisi.Items.Add("sttOrjMik", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    frmUretimSiparisi.Items.Add("edtOrjMik", SAPbouiCOM.BoFormItemTypes.it_EDIT);

                    SAPbouiCOM.StaticText OrjinalStatic = ((SAPbouiCOM.StaticText)frmUretimSiparisi.Items.Item("sttOrjMik").Specific);
                    OrjinalStatic.Caption = "Orjinal Miktar";
                    frmUretimSiparisi.Items.Item("sttOrjMik").Width = frmUretimSiparisi.Items.Item("255000116").Width;
                    frmUretimSiparisi.Items.Item("sttOrjMik").Top = frmUretimSiparisi.Items.Item("255000116").Top + frmUretimSiparisi.Items.Item("255000116").Height + 5;
                    frmUretimSiparisi.Items.Item("sttOrjMik").Height = frmUretimSiparisi.Items.Item("255000116").Height;
                    frmUretimSiparisi.Items.Item("sttOrjMik").Left = frmUretimSiparisi.Items.Item("255000116").Left;
                    frmUretimSiparisi.Items.Item("sttOrjMik").LinkTo = "255000116";

                    frmUretimSiparisi.Items.Item("edtOrjMik").Width = frmUretimSiparisi.Items.Item("255000122").Width;
                    frmUretimSiparisi.Items.Item("edtOrjMik").Top = frmUretimSiparisi.Items.Item("sttOrjMik").Top;
                    frmUretimSiparisi.Items.Item("edtOrjMik").Height = frmUretimSiparisi.Items.Item("sttOrjMik").Height;
                    frmUretimSiparisi.Items.Item("edtOrjMik").Left = frmUretimSiparisi.Items.Item("sttOrjMik").Left + frmUretimSiparisi.Items.Item("sttOrjMik").Width + 5;
                    frmUretimSiparisi.Items.Item("edtOrjMik").LinkTo = frmUretimSiparisi.Items.Item("sttOrjMik").UniqueID;
                    frmUretimSiparisi.Items.Item("edtOrjMik").Enabled = false;
                    string db = frmUretimSiparisi.DataSources.DBDataSources.Item(0).TableName;
                    ((SAPbouiCOM.EditText)frmUretimSiparisi.Items.Item("edtOrjMik").Specific).DataBind.SetBound(true, db, "U_OrjMiktar");

                    try
                    {
                      
                        SAPbouiCOM.Form KTAForm = Handler.SAPApplication.Forms.Item("-65211");
                        KTAForm.Items.Item("edtOrjMik").Enabled = false;
                    }
                    catch (Exception ex)
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
                    //if (BusinessObjectInfo.BeforeAction)
                    //{
                    //    BubbleEvent = false;

                    //}
                    if (!BusinessObjectInfo.BeforeAction)
                    {
                        try
                        {
                            frmUretimSiparisi.Items.Item("edtOrjMik").Enabled = false;
                            try
                            {
                                SAPbouiCOM.Form KTAForm = Handler.SAPApplication.Forms.GetForm("-65211", 0);
                                KTAForm.Items.Item("edtOrjMik").Enabled = false;
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                    break;
                case BoEventTypes.et_FORM_DATA_UPDATE:
                    //BubbleEvent = false;
                    break;
                case BoEventTypes.et_FORM_DATA_DELETE:
                    break;
                case BoEventTypes.et_FORM_DATA_LOAD:
                    if (!BusinessObjectInfo.BeforeAction)
                    {
                        try
                        {
                            frmUretimSiparisi.Items.Item("edtOrjMik").Enabled = false;
                            try
                            {
                                SAPbouiCOM.Form KTAForm = Handler.SAPApplication.Forms.GetForm("-65211", 0);
                                KTAForm.Items.Item("edtOrjMik").Enabled = false;
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                        catch (Exception ex)
                        {
                        }

                    }
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
                    #region ne oldugunu bilmediğimden kapattımm,chn
                    //if (pVal.ItemUID == "Item_0" && !pVal.BeforeAction)
                    //{
                    //    frmUretimSiparisi.PaneLevel = 3;
                    //    try
                    //    {
                    //        if (oPartiMatrix.RowCount == 0)
                    //        {
                    //            oPartiMatrix.AddRow();

                    //            try
                    //            {
                    //                ((SAPbouiCOM.EditText)oPartiMatrix.Columns.Item("Col_0").Cells.Item(oPartiMatrix.RowCount).Specific).Value = DateTime.Now.ToString("yyyyMMdd") + "-" + edtUrunKodu.Value + "-" + "1";
                    //            }
                    //            catch (Exception)
                    //            {
                    //            }

                    //            try
                    //            {
                    //                ((SAPbouiCOM.EditText)oPartiMatrix.Columns.Item("Col_2").Cells.Item(oPartiMatrix.RowCount).Specific).Value = parseNumber.parservalues<double>(edtPlanlananMiktar.Value).ToString();
                    //            }
                    //            catch (Exception)
                    //            {
                    //            }

                    //            oPartiMatrix.AutoResizeColumns();
                    //        }
                    //    }
                    //    catch (Exception)
                    //    {
                    //    }
                    //} 
                    #endregion
                    break;
                case BoEventTypes.et_KEY_DOWN:
                    break;
                case BoEventTypes.et_GOT_FOCUS:
                    break;
                case BoEventTypes.et_LOST_FOCUS:
                    #region ne oldugunu bilmediğimden kapattım chn
                    //if (pVal.ColUID == "Col_2")
                    //{
                    //    return false;
                    //    double qty = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oPartiMatrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific).Value);


                    //    if (qty > 0)
                    //    {
                    //        string xmlColor = oPartiMatrix.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);
                    //        XDocument xDoc = XDocument.Parse(xmlColor);
                    //        var table = (from x in xDoc.Elements("Matrix").Elements("Rows").Elements("Row")
                    //                     select new
                    //                     {
                    //                         belgeNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                    //                         PartiNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                    //                         Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value),
                    //                     }).ToList();

                    //        double tableQty = table.Sum(x => x.Miktar);
                    //        double PlannedQty = parseNumber.parservalues<double>(edtPlanlananMiktar.Value);

                    //        if (PlannedQty - tableQty != 0)
                    //        {
                    //            oPartiMatrix.AddRow();
                    //            oPartiMatrix.ClearRowData(oPartiMatrix.RowCount);
                    //        }
                    //    }

                    //} 
                    #endregion
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    break;
                case BoEventTypes.et_CLICK:
                    if (UrtPrtSekli == "2")
                    {
                        if (pVal.ItemUID == "btnUrtPrc" && !pVal.BeforeAction)
                        {
                            try
                            {
                                if (frmUretimSiparisi.Mode == BoFormMode.fm_OK_MODE)
                                {
                                    int siparisNo = Convert.ToInt32(((SAPbouiCOM.EditText)frmUretimSiparisi.Items.Item("18").Specific).Value);
                                    string siparisTarihi = ((SAPbouiCOM.EditText)frmUretimSiparisi.Items.Item("24").Specific).Value.ToString();
                                    double siparisMik = parseNumber_Seperator.ConvertToDouble(((SAPbouiCOM.EditText)frmUretimSiparisi.Items.Item("12").Specific).Value.ToString());
                                    string urunKodu = ((SAPbouiCOM.EditText)frmUretimSiparisi.Items.Item("6").Specific).Value.ToString();
                                    string urunTanimi = ((SAPbouiCOM.EditText)frmUretimSiparisi.Items.Item("8").Specific).Value.ToString();

                                    AIFConn.UrtSipCog.LoadForms(siparisNo, siparisTarihi, siparisMik, urunKodu, urunTanimi);
                                }
                                else
                                {
                                    Handler.SAPApplication.MessageBox("Sadece Tamam modunda işlem yapılabilir.");
                                }
                            }
                            catch (Exception ex)
                            {
                                //Handler.SAPApplication.MessageBox("Hata oluştu." + ex.Message);
                            }
                        }
                    }
                     if (pVal.ItemUID == "1" && pVal.BeforeAction)
                    {
                        if (frmUretimSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            //string PlanlananMik = ((SAPbouiCOM.EditText)frmUretimSiparisi.Items.Item("12").Specific).Value.ToString();
                            //double planlananMiktar = parseNumber_Seperator.ConvertToDouble(PlanlananMik);
                            //((SAPbouiCOM.EditText)frmUretimSiparisi.Items.Item("edtOrjMik").Specific).Value = planlananMiktar.ToString();
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
                        frmUretimSiparisi = Handler.SAPApplication.Forms.Item(pVal.FormUID);
                        //frmInit = false; 
                        var xml = frmUretimSiparisi.GetAsXML();
                        formuid = pVal.FormUID;
                        AIFConn.Sys65211.LoadForms();
                    }
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
                    //if (pVal.ItemUID == "6" && !pVal.BeforeAction)
                    //{
                    //    try
                    //    {
                    //        SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                    //        oPartiMatrix.Clear();
                    //        oPartiMatrix.AddRow();

                    //        var itemcode = oDataTable.GetValue("Code", 0).ToString();
                    //        ((SAPbouiCOM.EditText)oPartiMatrix.Columns.Item("Col_0").Cells.Item(oPartiMatrix.RowCount).Specific).Value = DateTime.Now.ToString("yyyyMMdd") + "-" + itemcode + "-" + "1";
                    //        ((SAPbouiCOM.EditText)oPartiMatrix.Columns.Item("Col_2").Cells.Item(oPartiMatrix.RowCount).Specific).Value = parseNumber.parservalues<double>(edtPlanlananMiktar.Value).ToString();
                    //    }
                    //    catch (Exception)
                    //    {  
                    //    }
                    //}
                    //if (pVal.ItemUID == "6" && pVal.BeforeAction)
                    //{
                    //    if (oPartiMatrix.RowCount > 0)
                    //    {
                    //        var retval = Handler.SAPApplication.MessageBox("Parti İşlemleri ekranında girilmiş olan veriler kaldırılacaktır, devam etmek istiyor musunuz?", 1, "Evet", "Hayır");

                    //        if (retval == 2)
                    //        {
                    //            BubbleEvent = false;
                    //        }

                    //    }
                    //}
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
            if (pVal.MenuUID=="1281" || pVal.MenuUID=="1282" || pVal.MenuUID=="1290" || pVal.MenuUID == "1288" || pVal.MenuUID == "1289" || pVal.MenuUID == "1291" || pVal.MenuUID == "1304")
            {
                try
                {
                    frmUretimSiparisi.Items.Item("edtOrjMik").Enabled = false;
                    try
                    {
                        SAPbouiCOM.Form KTAForm = Handler.SAPApplication.Forms.GetForm("-65211", 0);
                        KTAForm.Items.Item("edtOrjMik").Enabled = false;
                    }
                    catch (Exception ex)
                    {
                    }
                }
                catch (Exception ex)
                {
                }

            }
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }
    }
}
