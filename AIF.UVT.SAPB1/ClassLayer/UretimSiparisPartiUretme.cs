using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using AIF.UVT.SAPB1.HelperClass;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing.Design;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class UretimSiparisPartiUretme
    {
        [ItemAtt(AIFConn.UretimSiparisPartiUretmeUID)]
        public SAPbouiCOM.Form frmuretimSiparisPartiUretme;

        //[ItemAtt("Item_10")]
        //public SAPbouiCOM.EditText EdtDocEntry;
        //[ItemAtt("1")]
        //public SAPbouiCOM.Button btnAddOrUpdate;

        [ItemAtt("Item_1")]
        public SAPbouiCOM.EditText EdtUretimSiparisNo;

        [ItemAtt("Item_3")]
        public SAPbouiCOM.EditText EdtUretimSiparisTarihi;

        [ItemAtt("Item_5")]
        public SAPbouiCOM.EditText EdtUretimSiparisMiktari;

        [ItemAtt("Item_6")]
        public SAPbouiCOM.Matrix oMatrixUretimSiparisNo;

        [ItemAtt("Item_13")]
        public SAPbouiCOM.EditText EdtUrunKodu;

        [ItemAtt("Item_16")]
        public SAPbouiCOM.EditText EdtUrunTanimi;

        string hataverildi = "";
        public void LoadForms()
        {
            if (hataverildi == "")
            {
                SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery("select \"U_UrtPrtSekli\" as \"UrtPrtSekli\" from \"@AIF_UVT_PARAM\"");

                string UrtPrtSekli = oRS.Fields.Item("UrtPrtSekli").Value.ToString();

                if (UrtPrtSekli == "2")
                {
                    Handler.SAPApplication.MessageBox("Üretim siparişi Yönetme Şekli Üretim Siparişini Grupla Olduğunda Bu Özellik Kullanılamaz.");
                    hataverildi = "Hata";
                    return;
                }

                ConstVariables.oFnc.LoadSAPXML(AIFConn.UretimSiparisPartiUretmeXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.UretimSiparisPartiUretmeXML));
                Functions.CreateUserOrSystemFormComponent<UretimSiparisPartiUretme>(AIFConn.UrtSipPart);

                InitForms();


            }
            else
            {
                hataverildi = "";
            }
        }
        public void InitForms()
        {
            try
            {
                frmuretimSiparisPartiUretme.EnableMenu("1283", false);
                frmuretimSiparisPartiUretme.EnableMenu("1284", false);
                frmuretimSiparisPartiUretme.EnableMenu("1286", false);


                //ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                //ConstVariables.oRecordset.DoQuery("Select \"DocEntry\" from \"@AIF_UVT_ACTPARAM\"");

                //if (ConstVariables.oRecordset.RecordCount > 0)
                //{
                //    frmAktivteParametre.Mode = BoFormMode.fm_FIND_MODE;
                //    EdtDocEntry.Value = "1";
                //    btnAddOrUpdate.Item.Click();

                //}

                oMatrixUretimSiparisNo.AutoResizeColumns();
            }
            catch (Exception ex)
            {
            }
        }

        string header = @"<?xml version=""1.0"" encoding=""UTF-16"" ?><dbDataSources uid=""@AIF_URT_PART1""><rows>{0}</rows></dbDataSources>";

        public static string operasyonplaniTarihi = "";

        string row = "<row><cells><cell><uid>U_PartiNo</uid><value>{0}</value></cell><cell><uid>U_Miktar</uid><value>{1}</value></cell><cell><uid>U_PartiKatSayi</uid><value>{2}</value></cell><cell><uid>U_PartiBasSaati</uid><value>{3}</value></cell><cell><uid>U_PartiBitSaati</uid><value>{4}</value></cell><cell><uid>U_Durum</uid><value>{5}</value></cell></cells></row>";
        bool satirdakiMiktarOtomatikDEgisiyor = false;

        private DateTime startdateGetir()
        {
            DateTime startdate = new DateTime(1900, 01, 01);

            try
            {
                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sql = "SELECT \"StartDate\" FROM \"OWOR\" WHERE \"DocEntry\" = '" + EdtUretimSiparisNo.Value.ToString() + "' ";

                ConstVariables.oRecordset.DoQuery(sql);

                if (ConstVariables.oRecordset.RecordCount > 0)
                {
                    startdate = Convert.ToDateTime(ConstVariables.oRecordset.Fields.Item(0).Value);
                }

                return startdate;
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Başlangıç tarihi getirilirken hata oluştu." + ex.Message);
            }

            return startdate;

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
                    if (!BusinessObjectInfo.BeforeAction)
                    {
                        XmlDocument key = new XmlDocument();
                        key.LoadXml(BusinessObjectInfo.ObjectKey);
                        string docentry = key.SelectNodes("AIF_URT_PARTParams").Item(0).SelectNodes("DocEntry").Item(0).InnerXml;

                        SAPbobsCOM.ProductionOrders oWOR = (SAPbobsCOM.ProductionOrders)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);

                        oWOR.GetByKey(Convert.ToInt32(EdtUretimSiparisNo.Value));

                        //oWOR.PlannedQuantity = parseNumber.parservalues<double>(EdtUretimSiparisMiktari.Value);

                        oWOR.UserFields.Fields.Item("U_PartiBelgeNo").Value = docentry;

                        ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        ConstVariables.oRecordset.DoQuery("Select ISNULL(\"U_ISTASYON\",'') from \"OITT\" where \"Code\" = '" + oWOR.ItemNo + "'");

                        string istasyon = ConstVariables.oRecordset.Fields.Item(0).Value.ToString();


                        if (istasyon != "")
                        {
                            oWOR.UserFields.Fields.Item("U_ISTASYON").Value = istasyon;
                        }

                        var resp = oWOR.Update();
                    }
                    break;
                case BoEventTypes.et_FORM_DATA_UPDATE:
                    break;
                case BoEventTypes.et_FORM_DATA_DELETE:
                    break;
                case BoEventTypes.et_FORM_DATA_LOAD:
                    try
                    {
                        if (!BusinessObjectInfo.BeforeAction)
                        {
                            for (int i = 1; i <= oMatrixUretimSiparisNo.RowCount - 1; i++)
                            {
                                oMatrixUretimSiparisNo.CommonSetting.SetCellEditable(i, 2, false);
                            }
                        }
                        else if (BusinessObjectInfo.BeforeAction)
                        {
                            oMatrixUretimSiparisNo.Columns.Item(2).Editable = true;
                        }
                    }
                    catch (Exception)
                    {
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
        public string NoktalamaIsaretiDegistir(string val1)
        {
            string Val = Program.SAPnfi == "," ? val1.ToString().Replace(",", ".") : Program.SAPnfi == "." ? val1.ToString().Replace(".", ",") : val1.ToString();

            return Val;
        }
        bool col2LostFocus = false;
        bool sonrakiSatir = false;
        bool oncekiSatir = false;
        double siparismiktari = -1;
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
                    //if (pVal.ColUID == "Col_1" && !pVal.BeforeAction)
                    //{
                    //    string Durum = ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_5").Cells.Item(pVal.Row).Specific).Value;

                    //    if (Durum == "C")
                    //    {

                    //        //oMatrixUretimSiparisNo.CommonSetting.SetCellEditable(pVal.Row, 2, false);
                    //    } 
                    //}

                    if (pVal.ColUID == "Col_3" && !pVal.BeforeAction)
                    {
                        if (col2LostFocus && sonrakiSatir)
                        {
                            oMatrixUretimSiparisNo.SetCellFocus(oMatrixUretimSiparisNo.RowCount, 2);
                            col2LostFocus = false;
                            sonrakiSatir = false;
                        }
                        else if (col2LostFocus && oncekiSatir)
                        {
                            oMatrixUretimSiparisNo.SetCellFocus(oMatrixUretimSiparisNo.RowCount, 2);
                            col2LostFocus = false;
                            oncekiSatir = false;
                        }
                    }
                    else if (pVal.ItemUID == "Item_5" && !pVal.BeforeAction)
                    {

                        siparismiktari = parseNumber.parservalues<double>(EdtUretimSiparisMiktari.Value);

                    }
                    break;
                case BoEventTypes.et_LOST_FOCUS:
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    break;
                case BoEventTypes.et_CLICK:
                    if (pVal.ItemUID == "1" && pVal.BeforeAction)
                    {
                        //string xml = oMatrixUretimSiparisNo.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                        //var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                        //            select new
                        //            {
                        //                PartiNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                        //                Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value),
                        //            }).ToList();

                        //double siparisMiktari = parseNumber.parservalues<double>(EdtUretimSiparisMiktari.Value);
                        //double matrixtoplammiktar = parseNumber.parservalues<double>(rows.Sum(x => x.Miktar).ToString());
                        //double sonuc = siparisMiktari - matrixtoplammiktar;
                        //if (sonuc != 0)
                        //{
                        //    Handler.SAPApplication.MessageBox(string.Format("Belgenin tüm sipariş miktarı karşılanmadığından işleme devam edilemez. Kalan Miktar {0}'dır.", sonuc));
                        //    BubbleEvent = false;
                        //}
                    }
                    break;
                case BoEventTypes.et_DOUBLE_CLICK:
                    break;
                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                    break;
                case BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
                    break;
                case BoEventTypes.et_VALIDATE:
                    if (pVal.ColUID == "Col_1" && pVal.ItemChanged && !pVal.BeforeAction)
                    {
                        try
                        {
                            #region startdate ile parti no üretme 18 şubat 2022
                            DateTime baslangictarihi = startdateGetir();

                            if (baslangictarihi.Year == 1900)
                            {
                                Handler.SAPApplication.MessageBox("Başlangıç tarihi geçerli değildir");
                                return false;
                            }
                            #endregion

                            frmuretimSiparisPartiUretme.Freeze(true);
                            var deger = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value);
                            if (deger > 0)
                            {
                                string xml = oMatrixUretimSiparisNo.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                            select new
                                            {
                                                PartiNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                                Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value),
                                                PartiKatSayisi = Math.Round(parseNumber.parservalues<double>(EdtUretimSiparisMiktari.Value) / parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value), 6),
                                                PartiBasSaati = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value,
                                                PartiBitSaati = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,
                                                PartiDurumu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                                            }).ToList();


                                double siparisMiktari = parseNumber.parservalues<double>(EdtUretimSiparisMiktari.Value);
                                double matrixtoplammiktar = parseNumber.parservalues<double>(rows.Sum(x => x.Miktar).ToString());
                                double sonuc = siparisMiktari - matrixtoplammiktar;
                                if (sonuc > 0)
                                {

                                    if (frmuretimSiparisPartiUretme.Mode == BoFormMode.fm_ADD_MODE)
                                    {
                                        #region eski parti no üretme
                                        //rows.Insert(rows.Count, new { PartiNo = DateTime.Now.ToString("yyyyMMdd") + "-" + EdtUretimSiparisNo.Value + "-" + (rows.Count + 1), Miktar = siparisMiktari - matrixtoplammiktar, PartiKatSayisi = Math.Round(siparisMiktari / (siparisMiktari - matrixtoplammiktar), 6), PartiBasSaati = "", PartiBitSaati = "", PartiDurumu = "" });
                                        #endregion

                                        #region startdate ile parti no üretme 18 şubat 2022
                                        rows.Insert(rows.Count, new { PartiNo = baslangictarihi.ToString("yyyyMMdd") + "-" + EdtUretimSiparisNo.Value + "-" + (rows.Count + 1), Miktar = siparisMiktari - matrixtoplammiktar, PartiKatSayisi = Math.Round(siparisMiktari / (siparisMiktari - matrixtoplammiktar), 6), PartiBasSaati = "", PartiBitSaati = "", PartiDurumu = "" });
                                        #endregion
                                    }
                                    else
                                    {
                                        string val = ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_0").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value.ToString();

                                        var satir = val.Split('-');
                                        val = satir[satir.Length - 1];


                                        string partino = satir[0].ToString() + "-" + satir[1].ToString();

                                        partino = partino + "-" + Convert.ToString(Convert.ToInt32(val) + 1);

                                        rows.Insert(rows.Count, new { PartiNo = partino, Miktar = siparisMiktari - matrixtoplammiktar, PartiKatSayisi = Math.Round(siparisMiktari / (siparisMiktari - matrixtoplammiktar), 6), PartiBasSaati = "", PartiBitSaati = "", PartiDurumu = "" });


                                    }

                                    string data = string.Join("", rows.Select(s => string.Format(row, s.PartiNo, NoktalamaIsaretiDegistir(s.Miktar.ToString()), NoktalamaIsaretiDegistir(s.PartiKatSayisi.ToString()), s.PartiBasSaati.ToString(), s.PartiBitSaati.ToString(), s.PartiDurumu.ToString())));

                                    frmuretimSiparisPartiUretme.DataSources.DBDataSources.Item("@AIF_URT_PART1").LoadFromXML(string.Format(header, data));

                                    oMatrixUretimSiparisNo.AutoResizeColumns();

                                    EdtUrunKodu.Item.Click();
                                    for (int i = 1; i <= oMatrixUretimSiparisNo.RowCount - 1; i++)
                                    {
                                        oMatrixUretimSiparisNo.CommonSetting.SetCellEditable(i, 2, false);
                                    }

                                    col2LostFocus = true;
                                    sonrakiSatir = true;
                                    //oMatrixUretimSiparisNo.SetCellFocus(pVal.Row + 1, 2);
                                    //oMatrixUretimSiparisNo.Columns.Item("Col_1").Cells.Item(pVal.Row + 1).Click();
                                }
                                else if (sonuc < 0)
                                {
                                    Handler.SAPApplication.MessageBox(string.Format("Kalan Miktar {0}'dır kalan kadar değer otomatik atanmıştır.", deger - (sonuc * -1)));
                                    ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = NoktalamaIsaretiDegistir(parseNumber.parservalues<double>((deger - (sonuc * -1)).ToString()).ToString());
                                }
                            }
                            else if (deger == 0)
                            {
                                if (pVal.Row > 1)
                                {
                                    #region sıfır girildiğinde satır silme kaldırıldı
                                    //oMatrixUretimSiparisNo.DeleteRow(pVal.Row);
                                    //oMatrixUretimSiparisNo.CommonSetting.SetCellEditable(pVal.Row - 1, 2, true); 
                                    #endregion


                                    #region miktar girilen satırdan önceki satır durumu boş ise hesaplama yapar
                                    string oncekisatirDurum = ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_5").Cells.Item(pVal.Row - 1).Specific).Value.ToString();

                                    if (oncekisatirDurum != "")
                                    {
                                        string xml = oMatrixUretimSiparisNo.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                        var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                                    select new
                                                    {
                                                        PartiNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                                        Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value),
                                                        PartiKatSayisi = parseNumber.parservalues<double>(EdtUretimSiparisMiktari.Value) / parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value),
                                                    }).ToList();

                                        double sonsatir = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_1").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value);
                                        double siparisMiktari = parseNumber.parservalues<double>(EdtUretimSiparisMiktari.Value);
                                        double matrixtoplammiktar = parseNumber.parservalues<double>(rows.Sum(x => x.Miktar).ToString());
                                        double sonuc = siparisMiktari - (matrixtoplammiktar - sonsatir);
                                        ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_1").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = sonuc.ToString();
                                        //((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_2").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = (siparisMiktari / sonuc).ToString();
                                        ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_2").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = NoktalamaIsaretiDegistir(parseNumber.parservalues<double>(((siparisMiktari / sonuc).ToString())).ToString());
                                        Handler.SAPApplication.MessageBox(string.Format("Kalan Miktar 0 veya boş girilemez."));
                                    }
                                    else
                                    {
                                        oMatrixUretimSiparisNo.DeleteRow(pVal.Row);
                                        oMatrixUretimSiparisNo.CommonSetting.SetCellEditable(pVal.Row - 1, 2, true);


                                        string xml = oMatrixUretimSiparisNo.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                        var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                                    select new
                                                    {
                                                        PartiNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                                        Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value),
                                                        PartiKatSayisi = parseNumber.parservalues<double>(EdtUretimSiparisMiktari.Value) / parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value),
                                                    }).ToList();

                                        double sonsatir = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_1").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value);
                                        double siparisMiktari = parseNumber.parservalues<double>(EdtUretimSiparisMiktari.Value);
                                        double matrixtoplammiktar = parseNumber.parservalues<double>(rows.Sum(x => x.Miktar).ToString());
                                        double sonuc = siparisMiktari - (matrixtoplammiktar - sonsatir);
                                        ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_1").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = sonuc.ToString();
                                        //((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_2").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = (siparisMiktari / sonuc).ToString();
                                        ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_2").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = NoktalamaIsaretiDegistir(parseNumber.parservalues<double>(((siparisMiktari / sonuc).ToString())).ToString());
                                    }
                                    #endregion
                                }
                                else if (pVal.Row == 1)
                                {
                                    ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_1").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = EdtUretimSiparisMiktari.Value.ToString();
                                    ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_2").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = "1";
                                    oMatrixUretimSiparisNo.AutoResizeColumns();
                                }


                                col2LostFocus = true;
                                oncekiSatir = true;
                            }
                        }
                        catch (Exception ex)
                        {
                            Handler.SAPApplication.StatusBar.SetText("Hesaplama yapılırken hata oluştu. Hata Kodu A0018", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        }
                        finally
                        {
                            frmuretimSiparisPartiUretme.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "Item_5" && pVal.ItemChanged && !pVal.BeforeAction)
                    {
                        try
                        {
                            #region startdate ile parti no üretme 18 şubat 2022
                            DateTime baslangictarihi = startdateGetir();

                            if (baslangictarihi.Year == 1900)
                            {
                                Handler.SAPApplication.MessageBox("Başlangıç tarihi geçerli değildir");
                                return false;
                            }
                            #endregion

                            if (satirdakiMiktarOtomatikDEgisiyor)
                            {
                                satirdakiMiktarOtomatikDEgisiyor = false;
                                return false;
                            }
                            if (frmuretimSiparisPartiUretme.Mode == BoFormMode.fm_FIND_MODE || frmuretimSiparisPartiUretme.Mode == BoFormMode.fm_OK_MODE)
                            {
                                return false;
                            }

                            frmuretimSiparisPartiUretme.Freeze(true);

                            if (oMatrixUretimSiparisNo.RowCount > 0)
                            {
                                if (frmuretimSiparisPartiUretme.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    oMatrixUretimSiparisNo.Clear();
                                    oMatrixUretimSiparisNo.AddRow();

                                    #region eski parti no üretme 
                                    //((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_0").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = DateTime.Now.ToString("yyyMMdd") + "-" + EdtUretimSiparisNo.Value + "-" + "1";
                                    #endregion

                                    #region startdate ile parti no üretme 18 şubat 2022
                                    ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_0").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = baslangictarihi.ToString("yyyMMdd") + "-" + EdtUretimSiparisNo.Value + "-" + "1";
                                    #endregion

                                    ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_1").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = EdtUretimSiparisMiktari.Value.ToString();
                                    oMatrixUretimSiparisNo.AutoResizeColumns();

                                    oMatrixUretimSiparisNo.CommonSetting.SetCellEditable(1, 2, true);
                                }

                                if (frmuretimSiparisPartiUretme.Mode == BoFormMode.fm_UPDATE_MODE)
                                {

                                    string xml = oMatrixUretimSiparisNo.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                    var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                                select new
                                                {
                                                    PartiNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                                    Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value),
                                                    PartiDurumu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                                                    SatirNo = x.ElementsBeforeSelf().Count() + 1
                                                }).ToList();

                                    rows = rows.OrderByDescending(y => y.SatirNo).ToList();

                                    for (int i = 1; i <= oMatrixUretimSiparisNo.RowCount; i++)
                                    {
                                        oMatrixUretimSiparisNo.CommonSetting.SetCellEditable(i, 2, true);

                                    }


                                    if (oMatrixUretimSiparisNo.RowCount > 0)
                                    {
                                        if (rows.Where(x => x.PartiDurumu != "").Count() <= 0)
                                        {
                                            int retval = Handler.SAPApplication.MessageBox("Parti miktarları yeniden hesaplanacaktır.Devam etmek istiyor musunuz?", 1, "Evet", "Hayır");

                                            if (retval != 1)
                                            {
                                                for (int i = 1; i <= oMatrixUretimSiparisNo.RowCount - 1; i++)
                                                {
                                                    oMatrixUretimSiparisNo.CommonSetting.SetCellEditable(i, 2, false);

                                                }


                                                satirdakiMiktarOtomatikDEgisiyor = true;
                                                double toplamMiktar = parseNumber.parservalues<double>(rows.Sum(x => x.Miktar).ToString());

                                                EdtUretimSiparisMiktari.Value = toplamMiktar.ToString();

                                                return true;
                                            }
                                        }
                                    }
                                    foreach (var item in rows)
                                    {
                                        if (item.PartiDurumu != "C" && item.PartiDurumu != "B")
                                        {
                                            oMatrixUretimSiparisNo.DeleteRow(item.SatirNo);
                                        }

                                    }

                                    if (oMatrixUretimSiparisNo.RowCount == 0)
                                    {

                                        oMatrixUretimSiparisNo.Clear();
                                        oMatrixUretimSiparisNo.AddRow();
                                        oMatrixUretimSiparisNo.CommonSetting.SetCellEditable(oMatrixUretimSiparisNo.RowCount, 2, true);

                                        #region eski partino üretme
                                        //((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_0").Cells.Item(1).Specific).Value = DateTime.Now.ToString("yyyMMdd") + "-" + EdtUretimSiparisNo.Value + "-" + "1";
                                        #endregion

                                        #region startdate ile parti no üretme 18 şubat 2022
                                        ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_0").Cells.Item(1).Specific).Value = baslangictarihi.ToString("yyyMMdd") + "-" + EdtUretimSiparisNo.Value + "-" + "1";
                                        #endregion

                                        ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_1").Cells.Item(1).Specific).Value = EdtUretimSiparisMiktari.Value.ToString();
                                        ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_2").Cells.Item(1).Specific).Value = Convert.ToDouble(1).ToString();
                                    }
                                    else
                                    {
                                        for (int i = 1; i <= oMatrixUretimSiparisNo.RowCount; i++)
                                        {
                                            oMatrixUretimSiparisNo.CommonSetting.SetCellEditable(i, 2, false);

                                        }


                                        string val = ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_0").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value.ToString();

                                        var satir = val.Split('-');
                                        val = satir[satir.Length - 1];

                                        frmuretimSiparisPartiUretme.DataSources.DBDataSources.Item("@AIF_URT_PART1").Clear();
                                        oMatrixUretimSiparisNo.AddRow();


                                        #region eski partino üretme
                                        //((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_0").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = DateTime.Now.ToString("yyyMMdd") + "-" + EdtUretimSiparisNo.Value + "-" + Convert.ToString(Convert.ToInt32(val) + 1);
                                        #endregion

                                        #region startdate ile parti no üretme 18 şubat 2022
                                        ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_0").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = baslangictarihi.ToString("yyyMMdd") + "-" + EdtUretimSiparisNo.Value + "-" + Convert.ToString(Convert.ToInt32(val) + 1);
                                        #endregion

                                        xml = oMatrixUretimSiparisNo.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                        rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                                select new
                                                {
                                                    PartiNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                                    Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value),
                                                    PartiDurumu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                                                    SatirNo = x.ElementsBeforeSelf().Count() + 1
                                                }).ToList();

                                        double baslanmisVeBitmisPartiToplami = parseNumber.parservalues<double>(rows.Sum(x => x.Miktar).ToString());

                                        double yeniMiktar = parseNumber.parservalues<double>(EdtUretimSiparisMiktari.Value);

                                        double sonuc = yeniMiktar - baslanmisVeBitmisPartiToplami;

                                        if (sonuc > 0)
                                        {
                                            ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_1").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = sonuc.ToString();

                                            double sonsatir = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_1").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value);
                                            double siparisMiktari = parseNumber.parservalues<double>(EdtUretimSiparisMiktari.Value);
                                            double matrixtoplammiktar = parseNumber.parservalues<double>(rows.Sum(x => x.Miktar).ToString());

                                            //double sonuc2 = yeniMiktar - ( baslanmisVeBitmisPartiToplami - sonuc); 
                                            double sonuc2 = siparisMiktari - (matrixtoplammiktar - sonsatir);

                                            ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_2").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = NoktalamaIsaretiDegistir(parseNumber.parservalues<double>(((siparisMiktari / sonuc2).ToString())).ToString());

                                        }
                                        else
                                        {

                                            if (yeniMiktar < baslanmisVeBitmisPartiToplami)
                                            {
                                                if (siparismiktari != -1)
                                                {
                                                    EdtUretimSiparisMiktari.Value = siparismiktari.ToString();
                                                    Handler.SAPApplication.MessageBox("Sipariş miktarı, üretilen/devam eden miktarlardan küçük olamaz.");

                                                    //oMatrixUretimSiparisNo.DeleteRow(oMatrixUretimSiparisNo.RowCount);
                                                }
                                            }

                                            if (yeniMiktar == baslanmisVeBitmisPartiToplami)
                                            {
                                                //for (int i = 1; i <= oMatrixUretimSiparisNo.RowCount; i++)
                                                //{
                                                xml = oMatrixUretimSiparisNo.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                                rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                                        select new
                                                        {
                                                            PartiNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                                            Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value),
                                                            PartiDurumu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                                                            SatirNo = x.ElementsBeforeSelf().Count() + 1
                                                        }).ToList();


                                                foreach (var item in rows)
                                                {
                                                    if (item.PartiDurumu != "C" && item.PartiDurumu != "B")
                                                    {
                                                        oMatrixUretimSiparisNo.DeleteRow(item.SatirNo);
                                                    }

                                                }
                                                //var miktarolmayan = rows.Where(x => x.Miktar == "").Select(y => y.SatirNo).FirstOrDefault();


                                                //if (((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_2").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value == "")
                                                //{
                                                //    oMatrixUretimSiparisNo.DeleteRow(oMatrixUretimSiparisNo.RowCount);
                                                //}
                                                //}
                                            }
                                        }
                                    }


                                }
                            }

                            oMatrixUretimSiparisNo.AutoResizeColumns();
                        }
                        catch (Exception ex)
                        {
                        }
                        finally
                        {
                            frmuretimSiparisPartiUretme.Freeze(false);
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
                    if (pVal.ItemUID == "Item_1" && pVal.BeforeAction)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                        oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                        SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                        oCFL = frmuretimSiparisPartiUretme.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                        SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                        SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                        SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                        oCFL.SetConditions(oEmptyConts);
                        oCons = oCFL.GetConditions();

                        oCon = oCons.Add();
                        //oCon.Alias = "Status";
                        //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        //oCon.CondVal = "R";


                        //oCon.Relationship = BoConditionRelationship.cr_AND;
                        //oCon = oCons.Add();
                        oCon.Alias = "U_PartiBelgeNo";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_IS_NULL;
                        //commit

                        oCFL.SetConditions(oCons);
                    }
                    else if (pVal.ItemUID == "Item_1" && !pVal.BeforeAction)
                    {
                        try
                        {
                            oMatrixUretimSiparisNo.Columns.Item(2).Editable = true;

                            //for (int i = 1; i <= oMatrixUretimSiparisNo.RowCount; i++)
                            //{
                            //    oMatrixUretimSiparisNo.CommonSetting.SetCellEditable(i, 2, true);
                            //}

                            oMatrixUretimSiparisNo.Clear();
                            SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                            string Val = "";
                            Val = oDataTable.GetValue("DocNum", 0).ToString();
                            try
                            {
                                EdtUretimSiparisNo.Value = Val;
                            }
                            catch (Exception)
                            {
                            }

                            try
                            {
                                DateTime v = Convert.ToDateTime(oDataTable.GetValue("DueDate", 0));
                                EdtUretimSiparisTarihi.Value = v.ToString("yyyyMMdd");
                            }
                            catch (Exception)
                            {
                            }

                            try
                            {
                                double value = parseNumber.parservalues<double>(oDataTable.GetValue("PlannedQty", 0).ToString());
                                EdtUretimSiparisMiktari.Value = value.ToString();
                            }
                            catch (Exception)
                            {
                            }

                            Val = oDataTable.GetValue("ItemCode", 0).ToString();
                            try
                            {
                                EdtUrunKodu.Value = Val;
                            }
                            catch (Exception)
                            {
                            }


                            Val = oDataTable.GetValue("ProdName", 0).ToString();
                            try
                            {
                                EdtUrunTanimi.Value = Val;
                            }
                            catch (Exception)
                            {
                            }

                            #region startdate ile parti no üretme 18 şubat 2022
                            DateTime baslangictarihi = startdateGetir();

                            if (baslangictarihi.Year == 1900)
                            {
                                Handler.SAPApplication.MessageBox("Başlangıç tarihi geçerli değildir");
                                return false;
                            }
                            #endregion

                            oMatrixUretimSiparisNo.AddRow();
                            //oMatrixUretimSiparisNo.CommonSetting.SetCellEditable(1, 2, true);
                            #region eski parti no üretme
                            //((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_0").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = DateTime.Now.ToString("yyyMMdd") + "-" + EdtUretimSiparisNo.Value + "-" + "1";
                            #endregion

                            #region startdate ile parti no üretme 18 şubat 2022
                            ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_0").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = baslangictarihi.ToString("yyyMMdd") + "-" + EdtUretimSiparisNo.Value + "-" + "1";
                            #endregion

                            ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_1").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = EdtUretimSiparisMiktari.Value.ToString();
                            ((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_2").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = "1";
                            oMatrixUretimSiparisNo.AutoResizeColumns();


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
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }
    }
}
