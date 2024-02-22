using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using AIF.UVT.SAPB1.HelperClass;
using AIF.UVT.SAPB1.Models;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class SatinalmaIskontoGiris
    {
        [ItemAtt(AIFConn.SatinalmaIskontoGirisEkraniUID)]
        public static SAPbouiCOM.Form frmSatinalmaIskontoGiris;

        //[ItemAtt("Item_1")]
        //public SAPbouiCOM.ComboBox oComboCalisan;

        [ItemAtt("Item_8")]
        public SAPbouiCOM.Matrix oMatrixUrunler;

        [ItemAtt("Item_10")]
        public SAPbouiCOM.EditText oEditDocEntry;

        [ItemAtt("Item_5")]
        public SAPbouiCOM.EditText oEditGecerlilikBaslangicTarihi;

        [ItemAtt("Item_7")]
        public SAPbouiCOM.EditText oEditGecerlilikBitisTarihi;

        [ItemAtt("Item_20")]
        public SAPbouiCOM.ComboBox oComboDurum;

        [ItemAtt("Item_11")]
        public SAPbouiCOM.EditText oEditSaticiKodu;

        [ItemAtt("Item_25")]
        public SAPbouiCOM.EditText oEditSaticiTanimi;

        [ItemAtt("Item_14")]
        public SAPbouiCOM.ComboBox oComboFiyatListesi;
        //[ItemAtt("Item_22")]
        //public SAPbouiCOM.ComboBox oComboIndirimTipi;

        //[ItemAtt("Item_16")]
        //public SAPbouiCOM.CheckBox oCheckboxMusteri;

        //[ItemAtt("Item_11")]
        //public SAPbouiCOM.CheckBox oCheckboxBolge;

        //private SAPbouiCOM.DataTable oDataTable = null;

        string header = @"<?xml version=""1.0"" encoding=""UTF-16"" ?><dbDataSources uid=""@AIF_SATINALMAISK1""><rows>{0}</rows></dbDataSources>";


        string row = "<row><cells><cell><uid>U_UrunKodu</uid><value>{0}</value></cell><cell><uid>U_UrunTanimi</uid><value>{1}</value></cell><cell><uid>U_MiktarOlcegi</uid><value>{2}</value></cell><cell><uid>U_BirinciIsk</uid><value>{3}</value></cell><cell><uid>U_IkinciIsk</uid><value>{4}</value></cell><cell><uid>U_UcuncuIsk</uid><value>{5}</value></cell><cell><uid>U_DorduncuIsk</uid><value>{6}</value></cell><cell><uid>U_BesinciIsk</uid><value>{7}</value></cell><cell><uid>U_BirimFiyat</uid><value>{8}</value></cell><cell><uid>U_ToplamIsk</uid><value>{9}</value></cell><cell><uid>U_Fiyat</uid><value>{10}</value></cell></cells></row>";

        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.SatinalmaIskontoGirisEkraniFrmXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.SatinalmaIskontoGirisEkraniFrmXML));
            Functions.CreateUserOrSystemFormComponent<SatinalmaIskontoGiris>(AIFConn.StnAlmIsk);

            InitForms();
        }
        string sql = "";

        public void InitForms()
        {
            try
            {
                //oMatrix.AddRow();

                frmSatinalmaIskontoGiris.EnableMenu("1283", false);
                frmSatinalmaIskontoGiris.EnableMenu("1284", false);
                frmSatinalmaIskontoGiris.EnableMenu("1286", false);

                List<Helper.ValidValue> list;
                list = Helper.GetValidValuesFromRS("Select \"ListNum\" as \"value\", \"ListName\" as \"description\"  from OPLN order by \"ListNum\"");

                Helper nesne = new Helper();
                nesne.ComboAction(frmSatinalmaIskontoGiris, "Item_14", Helper.ActionCombo.add, list);

                oMatrixUrunler.AutoResizeColumns();
                oComboDurum.Select(0, BoSearchKey.psk_Index);

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
                    if (BusinessObjectInfo.BeforeAction)
                    {
                        var urunKodu = frmSatinalmaIskontoGiris.DataSources.DBDataSources.Item(1).GetValue("U_UrunKodu", 0).ToString();

                        if (urunKodu == "")
                        {
                            frmSatinalmaIskontoGiris.DataSources.DBDataSources.Item(1).RemoveRecord(frmSatinalmaIskontoGiris.DataSources.DBDataSources.Item(1).Size - 1);
                        }
                    }
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
        string Val = "";

        string oncekiComboDegeri = "";
        List<FiyatListesi> fiyatListesis = new List<FiyatListesi>();
        public bool SAP_ItemEvent(string FormUID, ref ItemEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;

            switch (pVal.EventType)
            {
                case BoEventTypes.et_ALL_EVENTS:
                    break;
                case BoEventTypes.et_ITEM_PRESSED:
                    if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    {
                        try
                        {
                            oEditGecerlilikBaslangicTarihi.Item.Click();
                            oEditDocEntry.Item.Enabled = false;
                        }
                        catch (Exception)
                        {
                        }
                    }
                    break;
                case BoEventTypes.et_KEY_DOWN:
                    break;
                case BoEventTypes.et_GOT_FOCUS:
                    break;
                case BoEventTypes.et_LOST_FOCUS:
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    if (!pVal.BeforeAction && pVal.ItemUID == "Item_14")
                    {
                        if (oComboFiyatListesi.Value.Trim() != "")
                        {
                            if (oComboFiyatListesi.Value.Trim() != oncekiComboDegeri)
                            {
                                if (oncekiComboDegeri != "")
                                {
                                    var retval = Handler.SAPApplication.MessageBox("Yeni fiyat listesine göre hesaplamalar yeniden yapılacaktır. Devam etmek istiyor musunuz?", 1, "Evet", "Hayır");

                                    if (retval == 2)
                                    {
                                        frmSatinalmaIskontoGiris.DataSources.DBDataSources.Item(0).SetValue("U_FiyatListesi", 0, oncekiComboDegeri);
                                        return false;
                                    }
                                }
                                string sql = "Select T1.\"ItemCode\",T1.\"Price\" from \"OPLN\" as T0 INNER JOIN \"ITM1\" as T1 ON T0.\"ListNum\" = T1.\"PriceList\" where T0.\"ListNum\" = '" + oComboFiyatListesi.Value.Trim() + "'";

                                ConstVariables.oRecordset.DoQuery(sql);

                                string xml = ConstVariables.oRecordset.GetAsXML();

                                XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                                XDocument xDoc = XDocument.Parse(ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData));

                                fiyatListesis = (from t in xDoc.Descendants(ns + "Row")
                                                 select new FiyatListesi()
                                                 {
                                                     itemCode = (from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "ItemCode" select new XElement(k.Element(ns + "Value"))).First().Value,
                                                     fiyat = parseNumber.parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "Price" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                 }).ToList();

                                FiyatlariHesapla();
                            }
                        }
                    }
                    else if (pVal.BeforeAction && pVal.ItemUID == "Item_14")
                    {
                        oncekiComboDegeri = oComboFiyatListesi.Value.Trim();
                    }
                    break;
                case BoEventTypes.et_CLICK:
                    if (!pVal.BeforeAction && pVal.ItemUID == "Item_18")
                    {
                        AIFConn.StnAlmUrn.LoadForms();
                    }
                    else if (!pVal.BeforeAction && pVal.ItemUID == "Item_19")
                    {
                        AIFConn.StnAlmUrn.LoadForms("guncelleme");
                    }
                    else if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    {
                        if (frmSatinalmaIskontoGiris.Mode != BoFormMode.fm_FIND_MODE)
                        {
                            if (oEditSaticiKodu.Value == "")
                            {
                                Handler.SAPApplication.MessageBox("Satıcı kodu boş geçilemez.");
                                oEditGecerlilikBaslangicTarihi.Item.Click();

                            }
                            else if (oEditGecerlilikBaslangicTarihi.Value == "")
                            {
                                Handler.SAPApplication.MessageBox("Geçerlilik Başlangıç tarihi boş geçilemez.");
                                oEditGecerlilikBaslangicTarihi.Item.Click();
                            }
                            else if (oEditGecerlilikBitisTarihi.Value == "")
                            {
                                Handler.SAPApplication.MessageBox("Geçerlilik Bitiş tarihi boş geçilemez.");
                                oEditGecerlilikBaslangicTarihi.Item.Click();
                            }
                            else
                            {
                                DateTime dtBaslangic = new DateTime(Convert.ToInt32(oEditGecerlilikBaslangicTarihi.Value.Substring(0, 4)), Convert.ToInt32(oEditGecerlilikBaslangicTarihi.Value.Substring(4, 2)), Convert.ToInt32(oEditGecerlilikBaslangicTarihi.Value.Substring(6, 2)));
                                DateTime dtBitis = new DateTime(Convert.ToInt32(oEditGecerlilikBitisTarihi.Value.Substring(0, 4)), Convert.ToInt32(oEditGecerlilikBitisTarihi.Value.Substring(4, 2)), Convert.ToInt32(oEditGecerlilikBitisTarihi.Value.Substring(6, 2)));

                                if (dtBitis < dtBaslangic)
                                {
                                    Handler.SAPApplication.MessageBox("Bitiş tarihi başlangıç tarihinden küçük olamaz.");
                                    oEditGecerlilikBitisTarihi.Item.Click();
                                }
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
                    if (!pVal.BeforeAction && pVal.ItemChanged && pVal.ItemUID == "Item_8" && (pVal.ColUID == "Col_2" || pVal.ColUID == "Col_3" || pVal.ColUID == "Col_4" || pVal.ColUID == "Col_5" || pVal.ColUID == "Col_6"))
                    {
                        try
                        {
                            frmSatinalmaIskontoGiris.Freeze(true);
                            double _birinciIsk = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixUrunler.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific).Value.ToString());
                            double _ikinciIsk = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixUrunler.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific).Value.ToString());
                            double _ucuncuIsk = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixUrunler.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific).Value.ToString());
                            double _dorduncuIsk = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixUrunler.Columns.Item("Col_5").Cells.Item(pVal.Row).Specific).Value.ToString());
                            double _besinciIsk = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixUrunler.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific).Value.ToString());
                            double _birimFiyat = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixUrunler.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific).Value.ToString());

                            double _toplamIsk = 100 - (100 * (1 - (_birinciIsk / 100)) * (1 - (_ikinciIsk / 100)) * (1 - (_ucuncuIsk / 100)) * (1 - (_dorduncuIsk / 100)) * (1 - (_besinciIsk / 100)));

                            //((SAPbouiCOM.EditText)oMatrixUrunler.Columns.Item("Col_7").Cells.Item(pVal.Row).Specific).Value = _toplamIsk.ToString();
                            double _fiyat = _birimFiyat - _toplamIsk;

                            if (_fiyat < 0)
                            {
                                _fiyat = _fiyat * (-1);
                            }
                            oMatrixUrunler.SetCellWithoutValidation(pVal.Row, "Col_7", _toplamIsk.ToString(System.Globalization.CultureInfo.InvariantCulture).Replace(",", "."));
                            oMatrixUrunler.SetCellWithoutValidation(pVal.Row, "Col_9", _fiyat.ToString(System.Globalization.CultureInfo.InvariantCulture).Replace(",", "."));
                            oMatrixUrunler.AutoResizeColumns();
                        }
                        catch (Exception ex)
                        {
                        }
                        finally
                        {
                            frmSatinalmaIskontoGiris.Freeze(false);
                        }

                        if (oEditDocEntry.Value.ToString() != "")
                        {
                            frmSatinalmaIskontoGiris.Mode = BoFormMode.fm_UPDATE_MODE;
                        }
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
                    if (pVal.ItemUID == "Item_8" && pVal.ColUID == "Col_0" && pVal.BeforeAction)
                    {

                    }
                    else if (pVal.ItemUID == "Item_8" && pVal.ColUID == "Col_0" && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                        string Val = "";
                        try
                        {
                            Val = oDataTable.GetValue("ItemCode", 0).ToString();

                            var xml = oMatrixUrunler.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                            var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                        select new SatinalmaIskontolu_UrunEkle()
                                        {
                                            UrunKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                        }).ToList();

                            if (rows.Where(x => x.UrunKodu == Val).Count() > 0)
                            {
                                Handler.SAPApplication.MessageBox("Bu ürün daha önce eklenmiştir.");
                                return false;
                            }
                        }
                        catch (Exception)
                        {
                        }
                        try
                        {
                            ((SAPbouiCOM.EditText)oMatrixUrunler.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value = Val;
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            Val = oDataTable.GetValue("ItemName", 0).ToString();
                        }
                        catch (Exception)
                        {
                        }
                        try
                        {
                            ((SAPbouiCOM.EditText)oMatrixUrunler.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = Val;
                        }
                        catch (Exception)
                        {
                        }

                        oMatrixUrunler.SetCellWithoutValidation(pVal.Row, "Col_2", "0");
                        oMatrixUrunler.SetCellWithoutValidation(pVal.Row, "Col_3", "0");
                        oMatrixUrunler.SetCellWithoutValidation(pVal.Row, "Col_4", "0");
                        oMatrixUrunler.SetCellWithoutValidation(pVal.Row, "Col_5", "0");
                        oMatrixUrunler.SetCellWithoutValidation(pVal.Row, "Col_6", "0");
                        oMatrixUrunler.SetCellWithoutValidation(pVal.Row, "Col_7", "0");
                        oMatrixUrunler.SetCellWithoutValidation(pVal.Row, "Col_8", "0");
                        oMatrixUrunler.AutoResizeColumns();
                    }
                    else if (pVal.ItemUID == "Item_11" && pVal.BeforeAction)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                        oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                        SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                        oCFL = frmSatinalmaIskontoGiris.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                        SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                        SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                        SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                        oCFL.SetConditions(oEmptyConts);
                        oCons = oCFL.GetConditions();

                        oCon = oCons.Add();
                        oCon.Alias = "CardType";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "S";

                        oCFL.SetConditions(oCons);
                    }
                    else if (pVal.ItemUID == "Item_11" && !pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                            string Val = "";
                            Val = oDataTable.GetValue("CardCode", 0).ToString();
                            try
                            {
                                oEditSaticiKodu.Value = Val;
                            }
                            catch (Exception)
                            {
                            }

                            Val = oDataTable.GetValue("CardName", 0).ToString();
                            try
                            {
                                oEditSaticiTanimi.Value = Val;
                            }
                            catch (Exception)
                            {
                            }
                        }
                        catch (Exception)
                        {
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


        public void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && pVal.BeforeAction)
                {
                    if (itemUID == "Item_8")
                    {
                        int row = oMatrixUrunler.GetNextSelectedRow();
                        if (row != -1)
                        {
                            oMatrixUrunler.DeleteRow(row);
                            if (frmSatinalmaIskontoGiris.Mode == BoFormMode.fm_OK_MODE)
                            {
                                frmSatinalmaIskontoGiris.Mode = BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                    }

                }
                else if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && pVal.BeforeAction)
                {
                    if (itemUID == "Item_8")
                    {
                        frmSatinalmaIskontoGiris.DataSources.DBDataSources.Item(1).Clear();
                        oMatrixUrunler.AddRow();
                    }
                }
                else if (!pVal.BeforeAction && pVal.MenuUID == "1281")
                {
                    try
                    {
                        oEditDocEntry.Item.Enabled = true;
                        oEditDocEntry.Item.Click();
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
            }

        }
        string itemUID = "";


        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                var oForm = Handler.SAPApplication.Forms.ActiveForm;

                if (eventInfo.ItemUID != "")
                {
                    try
                    {
                        itemUID = eventInfo.ItemUID;
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

        public void UrunleriEkle(List<SatinalmaIskontolu_UrunEkle> urunEkles, string birinciIsk = "", string IkinciIsk = "", string UcuncuIsk = "", string DorduncuIsk = "", string BesinciIsk = "", string ToplamIsk = "", string MiktarOlcegi = "")
        {
            urunEkles.ForEach(x => x.Sira = 2);

            var xml = oMatrixUrunler.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

            var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                        select new SatinalmaIskontolu_UrunEkle()
                        {
                            UrunKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                            UrunTanimi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                            MiktarOlcegi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value,
                            birinciIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value,
                            ikinciIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value,
                            ucuncuIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,
                            dorduncuIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                            besinciIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value,
                            toplamIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value,
                            dahaOnceEkli = "Y",
                            Sira = 1
                        }).ToList();
            string ayniUrunlerHatasi = "";
            List<SatinalmaIskontolu_UrunEkle> _urunEkles = new List<SatinalmaIskontolu_UrunEkle>();
            foreach (var item in urunEkles)
            {
                if (rows.Where(x => x.UrunKodu == item.UrunKodu).Count() > 0)
                {
                    if (ayniUrunlerHatasi == "")
                    {
                        ayniUrunlerHatasi = "Aşağıda bulunan ürünler daha önce listeye eklenmiştir.";
                        ayniUrunlerHatasi += Environment.NewLine;
                        ayniUrunlerHatasi += item.UrunKodu + "-" + item.UrunTanimi;
                    }
                    else
                    {
                        ayniUrunlerHatasi += Environment.NewLine;
                        ayniUrunlerHatasi += item.UrunKodu + "-" + item.UrunTanimi;
                    }
                }
                else
                {
                    _urunEkles.Add(item);
                }
            }

            _urunEkles.AddRange(rows.ToArray());

            var urunEkleFinal = _urunEkles.OrderBy(x => x.Sira).ToList();

            string rowsFinal = string.Join("", urunEkleFinal.Select(y => string.Format(row, y.UrunKodu, y.UrunTanimi, y.dahaOnceEkli != "Y" ? MiktarOlcegi : y.MiktarOlcegi.ToString().Replace(",", "."), y.dahaOnceEkli != "Y" ? birinciIsk : y.birinciIskonto.ToString().Replace(",", "."), y.dahaOnceEkli != "Y" ? IkinciIsk : y.ikinciIskonto.ToString().Replace(",", "."), y.dahaOnceEkli != "Y" ? UcuncuIsk : y.ucuncuIskonto.ToString().Replace(",", "."), y.dahaOnceEkli != "Y" ? DorduncuIsk : y.dorduncuIskonto.ToString().Replace(",", "."), y.dahaOnceEkli != "Y" ? BesinciIsk : y.besinciIskonto.ToString().Replace(",", "."), parseNumber.parservalues<double>(fiyatListesis.Where(z => z.itemCode == y.UrunKodu).Select(u => u.fiyat).FirstOrDefault().ToString().Replace(",", ".")), y.dahaOnceEkli != "Y" ? ToplamIsk : y.toplamIskonto.ToString().Replace(",", "."), y.dahaOnceEkli != "Y" ? parseNumber.parservalues<double>(fiyatListesis.Where(z => z.itemCode == y.UrunKodu).Select(u => u.fiyat).FirstOrDefault().ToString().Replace(",", ".")) * (100 - parseNumber.parservalues<double>(ToplamIsk)) / 100 : (parseNumber.parservalues<double>(fiyatListesis.Where(z => z.itemCode == y.UrunKodu).Select(u => u.fiyat).FirstOrDefault().ToString().Replace(",", ".")) * (100 - parseNumber.parservalues<double>(y.toplamIskonto)))).ToString().Replace(",", ".")));

            //string rowsFinal = string.Join("", urunEkleFinal.Select(y => string.Format(row, y.UrunKodu, y.UrunTanimi, y.dahaOnceEkli != "Y" ? MiktarOlcegi : parseNumber.parservalues<double>(y.MiktarOlcegi), y.dahaOnceEkli != "Y" ? birinciIsk : parseNumber.parservalues<double>(y.birinciIskonto), y.dahaOnceEkli != "Y" ? IkinciIsk : parseNumber.parservalues<double>(y.ikinciIskonto), y.dahaOnceEkli != "Y" ? UcuncuIsk : parseNumber.parservalues<double>(y.ucuncuIskonto), y.dahaOnceEkli != "Y" ? DorduncuIsk : parseNumber.parservalues<double>(y.dorduncuIskonto), y.dahaOnceEkli != "Y" ? BesinciIsk : parseNumber.parservalues<double>(y.besinciIskonto), parseNumber.parservalues<double>(fiyatListesis.Where(z => z.itemCode == y.UrunKodu).Select(u => u.fiyat).FirstOrDefault().ToString()), y.dahaOnceEkli != "Y" ? ToplamIsk : parseNumber.parservalues<double>(y.toplamIskonto), y.dahaOnceEkli != "Y" ? parseNumber.parservalues<double>(fiyatListesis.Where(z => z.itemCode == y.UrunKodu).Select(u => u.fiyat).FirstOrDefault().ToString()) * (100 - ToplamIsk) / 100 : parseNumber.parservalues<double>(fiyatListesis.Where(z => z.itemCode == y.UrunKodu).Select(u => u.fiyat).FirstOrDefault().ToString()) * (100 - parseNumber.parservalues<double>(y.toplamIskonto)))));

            try
            {
                frmSatinalmaIskontoGiris.Freeze(true);
                frmSatinalmaIskontoGiris.DataSources.DBDataSources.Item("@AIF_SATINALMAISK1").LoadFromXML(string.Format(header, rowsFinal));
                if (oEditDocEntry.Value.ToString() != "")
                {
                    frmSatinalmaIskontoGiris.Mode = BoFormMode.fm_UPDATE_MODE;
                }

                oMatrixUrunler.AutoResizeColumns();
            }
            catch (Exception)
            {
            }
            finally
            {
                frmSatinalmaIskontoGiris.Freeze(false);
            }

            if (ayniUrunlerHatasi != "")
            {
                Handler.SAPApplication.MessageBox(ayniUrunlerHatasi);
            }
        }

        public void UrunleriGuncelle(List<SatinalmaIskontolu_UrunEkle> urunGuncelles, string birinciIsk = "", string IkinciIsk = "", string UcuncuIsk = "", string DorduncuIsk = "", string BesinciIsk = "", string ToplamIsk = "", string miktarOlcegi = "")
        {
            var xml = oMatrixUrunler.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

            var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                        select new SatinalmaIskontolu_UrunEkle()
                        {
                            UrunKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                            UrunTanimi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                            MiktarOlcegi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value,
                            birinciIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value,
                            ikinciIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value,
                            ucuncuIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,
                            dorduncuIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                            besinciIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value,
                            toplamIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value,
                            fiyat = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_9" select new XElement(y.Element("Value"))).First().Value,
                            dahaOnceEkli = "Y",
                            Sira = 1
                        }).ToList();

            foreach (var item in urunGuncelles)
            {
                rows.Where(x => x.UrunKodu == item.UrunKodu).ToList().ForEach(y =>
                {
                    y.MiktarOlcegi = miktarOlcegi;
                    y.birinciIskonto = birinciIsk;
                    y.ikinciIskonto = IkinciIsk;
                    y.ucuncuIskonto = UcuncuIsk;
                    y.dorduncuIskonto = DorduncuIsk;
                    y.besinciIskonto = BesinciIsk;
                    y.toplamIskonto = ToplamIsk;
                });
            }

            rows.ToList().ForEach(x => x.fiyat = (parseNumber.parservalues<double>(x.birimFiyat) * (100 - parseNumber.parservalues<double>(x.toplamIskonto)) / 100).ToString());

            string rowsFinal = string.Join("", rows.Select(y => string.Format(row, y.UrunKodu, y.UrunTanimi, y.MiktarOlcegi.ToString().Replace(",", "."), y.birinciIskonto.ToString().Replace(",", "."), y.ikinciIskonto.ToString().Replace(",", "."), y.ucuncuIskonto.ToString().Replace(",", "."), y.dorduncuIskonto.ToString().Replace(",", "."), y.besinciIskonto.ToString().Replace(",", "."), parseNumber.parservalues<double>(y.birimFiyat), y.toplamIskonto.ToString().Replace(",", "."), y.fiyat.ToString().Replace(",", "."))));

            //string rowsFinal = string.Join("", rows.Select(y => string.Format(row, y.UrunKodu, y.UrunTanimi, y.MiktarOlcegi, y.birinciIskonto, y.ikinciIskonto, y.ucuncuIskonto, y.dorduncuIskonto, y.besinciIskonto, y.birimFiyat, y.toplamIskonto, y.fiyat)));

            //string rowsFinal = string.Join("", rows.Select(y => string.Format(row, y.UrunKodu, y.UrunTanimi, parseNumber.parservalues<double>(y.MiktarOlcegi), parseNumber.parservalues<double>(y.birinciIskonto), parseNumber.parservalues<double>(y.ikinciIskonto), parseNumber.parservalues<double>(y.ucuncuIskonto), parseNumber.parservalues<double>(y.dorduncuIskonto), parseNumber.parservalues<double>(y.besinciIskonto), parseNumber.parservalues<double>(y.birimFiyat), parseNumber.parservalues<double>(y.toplamIskonto), parseNumber.parservalues<double>(y.fiyat))));

            try
            {
                frmSatinalmaIskontoGiris.Freeze(true);
                frmSatinalmaIskontoGiris.DataSources.DBDataSources.Item("@AIF_SATINALMAISK1").LoadFromXML(string.Format(header, rowsFinal));
                if (oEditDocEntry.Value.ToString() != "")
                {
                    frmSatinalmaIskontoGiris.Mode = BoFormMode.fm_UPDATE_MODE;
                }

                oMatrixUrunler.AutoResizeColumns();
            }
            catch (Exception)
            {
            }
            finally
            {
                frmSatinalmaIskontoGiris.Freeze(false);
            }
        }

        public void FiyatlariHesapla()
        {
            var xml = oMatrixUrunler.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

            var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                        select new SatinalmaIskontolu_UrunEkle()
                        {
                            UrunKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                            UrunTanimi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                            MiktarOlcegi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value,
                            birinciIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value,
                            ikinciIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value,
                            ucuncuIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,
                            dorduncuIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                            besinciIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value,
                            toplamIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value,
                            birimFiyat = fiyatListesis.Where(z => z.itemCode == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(u => u.fiyat).FirstOrDefault().ToString(),
                            dahaOnceEkli = "Y",
                            Sira = 1
                        }).ToList();

            if (rows.Count == 0)
            {
                return;
            }

            rows.ToList().ForEach(x => x.fiyat = (parseNumber.parservalues<double>(x.birimFiyat) * (100 - parseNumber.parservalues<double>(x.toplamIskonto)) / 100).ToString());


            string rowsFinal = string.Join("", rows.Select(y => string.Format(row, y.UrunKodu, y.UrunTanimi, y.MiktarOlcegi.ToString().Replace(",", "."), y.birinciIskonto.ToString().Replace(",", "."), y.ikinciIskonto.ToString().Replace(",", "."), y.ucuncuIskonto.ToString().Replace(",", "."), y.dorduncuIskonto.ToString().Replace(",", "."), y.besinciIskonto.ToString().Replace(",", "."), y.birimFiyat.ToString().Replace(",", "."), y.toplamIskonto.ToString().Replace(",", "."), y.fiyat.ToString().Replace(",", "."))));

            //string rowsFinal = string.Join("", rows.Select(y => string.Format(row, y.UrunKodu, y.UrunTanimi, parseNumber.parservalues<double>(y.MiktarOlcegi), parseNumber.parservalues<double>(y.birinciIskonto), parseNumber.parservalues<double>(y.ikinciIskonto), parseNumber.parservalues<double>(y.ucuncuIskonto), parseNumber.parservalues<double>(y.dorduncuIskonto), parseNumber.parservalues<double>(y.besinciIskonto), parseNumber.parservalues<double>(y.birimFiyat), parseNumber.parservalues<double>(y.toplamIskonto), parseNumber.parservalues<double>(y.fiyat))));

            try
            {
                frmSatinalmaIskontoGiris.Freeze(true);
                frmSatinalmaIskontoGiris.DataSources.DBDataSources.Item("@AIF_SATINALMAISK1").LoadFromXML(string.Format(header, rowsFinal));

                oMatrixUrunler.AutoResizeColumns();
            }
            catch (Exception)
            {
            }
            finally
            {
                frmSatinalmaIskontoGiris.Freeze(false);
            }
        }
    }
}
