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
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class GunlukSutRaporu
    {
        [ItemAtt(AIFConn.GunlukSutRaporuUID)]
        public SAPbouiCOM.Form frmGunlukSutRaporu;

        //[ItemAtt("Item_10")]
        //public SAPbouiCOM.EditText EdtDocEntry;
        //[ItemAtt("1")]
        //public SAPbouiCOM.Button btnAddOrUpdate;

        [ItemAtt("Item_3")]
        public SAPbouiCOM.EditText EdtBaslangicTarihi;
        [ItemAtt("Item_5")]
        public SAPbouiCOM.EditText EdtBitisTarihi;
        [ItemAtt("Item_0")]
        public SAPbouiCOM.Matrix oMatrix;

        SAPbouiCOM.DataTable oDataTable = null;
        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.GunlukSutRaporuXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.GunlukSutRaporuXML));
            Functions.CreateUserOrSystemFormComponent<GunlukSutRaporu>(AIFConn.GnlkSut);

            InitForms();
        }
        string header = @"<?xml version=""1.0"" encoding=""UTF-16"" ?><DataTable Uid=""DATA""><Columns><Column Uid=""UrunKodu"" Type=""1"" MaxLength=""254""/><Column Uid=""UrunTanim"" Type=""1"" MaxLength=""254""/><Column Uid=""Miktar"" Type=""5"" MaxLength=""0""/><Column Uid=""Depo"" Type=""1"" MaxLength=""100""/></Columns><Rows>{0}</Rows></DataTable>";


        string row = "<Row><Cells><Cell><ColumnUid>UrunKodu</ColumnUid><Value>{0}</Value></Cell><Cell><ColumnUid>UrunTanim</ColumnUid><Value>{1}</Value></Cell><Cell><ColumnUid>Miktar</ColumnUid><Value>{2}</Value></Cell><Cell><ColumnUid>Depo</ColumnUid><Value>{3}</Value></Cell></Cells></Row>";
        public void InitForms()
        {
            try
            {
                frmGunlukSutRaporu.EnableMenu("1283", false);
                frmGunlukSutRaporu.EnableMenu("1284", false);
                frmGunlukSutRaporu.EnableMenu("1286", false);
                oDataTable = frmGunlukSutRaporu.DataSources.DataTables.Add("DATA");
                oDataTable.Columns.Add("SipNo", BoFieldsType.ft_Integer);
                oDataTable.Columns.Add("UrunKodu", BoFieldsType.ft_AlphaNumeric, 100);
                oDataTable.Columns.Add("UrunAdi", BoFieldsType.ft_AlphaNumeric, 254);
                oDataTable.Columns.Add("YagliSut", BoFieldsType.ft_Rate);
                oDataTable.Columns.Add("YagsizSut", BoFieldsType.ft_Rate);
                oDataTable.Columns.Add("YariMamul", BoFieldsType.ft_AlphaNumeric, 254);


                string xml = oDataTable.SerializeAsXML(BoDataTableXmlSelect.dxs_All);

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
                        Listele(EdtBaslangicTarihi.Value, EdtBitisTarihi.Value);
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

        public class UrunAgaciSonucu
        {
            public int UretimSiparisNo { get; set; }

            public string UrunKodu { get; set; }

            public string UrunAdi { get; set; }

            public double YagliSut { get; set; }

            public double YagsizSut { get; set; }

            public string YariMamul { get; set; }

        }

        List<UrunAgaciSonucu> urunAgaciSonucus = new List<UrunAgaciSonucu>();


        private List<UrunAgaciSonucu> UrunAgaciIslemleri(string YariMamulKodu, string UretimSiparisNo, string UrunKodu, string UrunAdi, double DuzenlenenMiktar)
        {
            string bilesenKodu = "";
            string AltbilesenKodu = "";
            double miktar = 0;
            double baslikmiktar = 0;
            double xDegeri = 0;
            string sql = "Select T1.\"Code\" as \"BilesenKodu\",T1.\"Quantity\" as Miktar,T0.\"Qauntity\" as BaslikMiktar from OITT as T0 INNER JOIN ITT1 as T1 ON T0.[Code] = T1.[Father] INNER JOIN OITM as T2 ON T2.ItemCode = T1.Father where T0.\"Code\" = '" + YariMamulKodu + "' and T2.\"ItmsGrpCod\" = '106' and T1.Code like 'YRM%'";

            ConstVariables.oRecordset = (Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
            ConstVariables.oRecordset1 = (Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
            ConstVariables.oRecordset.DoQuery(sql);

            string sql2 = "";
            while (!ConstVariables.oRecordset.EoF)
            {
                miktar = parseNumber.parservalues<double>(ConstVariables.oRecordset.Fields.Item("Miktar").Value.ToString());
                bilesenKodu = ConstVariables.oRecordset.Fields.Item("BilesenKodu").Value.ToString();
                if (bilesenKodu == "YRM-10003" || bilesenKodu == "YRM-10002")
                {

                    #region İlk bileşende bulunursa.
                    string xmll = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                    XDocument xDoc = XDocument.Parse(xmll);
                    XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                    var rows = (from t in xDoc.Descendants(ns + "Row")
                                select new
                                {
                                    BilesenKodu = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "BilesenKodu" select new XElement(y.Element(ns + "Value"))).First().Value,
                                    Miktar = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "Miktar" select new XElement(y.Element(ns + "Value"))).First().Value),
                                    BaslikMiktar = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "BaslikMiktar" select new XElement(y.Element(ns + "Value"))).First().Value),
                                }).ToList();


                    if (rows.Where(x => x.BilesenKodu == "YRM-10003").Count() > 0 && bilesenKodu == "YRM-10003")
                    {
                        double yagliSutMiktari = 0;
                        double kayitedilismisYagliSutMiktari = 0;
                        double sonuc = 0;
                        if (urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Count() > 0)
                        //if (urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Count() > 0)
                        {

                            xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());

                            xDegeri = Math.Round(xDegeri, 6);

                            //yagliSutMiktari = miktar / xDegeri;
                            //kayitedilismisYagliSutMiktari = parseNumber.parservalues<double>(urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Select(y => y.YagliSut).FirstOrDefault().ToString());
                            //sonuc = kayitedilismisYagliSutMiktari + miktar;

                            sonuc = DuzenlenenMiktar / xDegeri;

                            sonuc = Math.Round(sonuc, 6);

                            urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).ToList().ForEach(y => y.YagliSut = sonuc);

                            //yagliSutMiktari = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());

                            //urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).ToList().ForEach(y => y.YagliSut = y.YagliSut + yagliSutMiktari);
                        }
                        else
                        {
                            xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());

                            xDegeri = Math.Round(xDegeri, 6);

                            //yagliSutMiktari = miktar / xDegeri;

                            sonuc = DuzenlenenMiktar / xDegeri;

                            sonuc = Math.Round(sonuc, 6);

                            urunAgaciSonucus.Add(new UrunAgaciSonucu { UretimSiparisNo = Convert.ToInt32(UretimSiparisNo), YagliSut = sonuc, UrunKodu = UrunKodu, UrunAdi = UrunAdi, YariMamul = YariMamulKodu });
                        }
                    }
                    else if (rows.Where(x => x.BilesenKodu == "YRM-10002").Count() > 0 && bilesenKodu == "YRM-10002")
                    {
                        double yagsizSutMiktari = 0;
                        double kayitedilismisYagsizSutMiktari = 0;
                        double sonuc = 0;
                        if (urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Count() > 0)
                        //if (urunAgaciSonucus.Count > 0)
                        {
                            yagsizSutMiktari = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.Miktar).FirstOrDefault().ToString());
                            //kayitedilismisYagsizSutMiktari = parseNumber.parservalues<double>(urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Select(y => y.YagsizSut).FirstOrDefault().ToString());
                            //sonuc = kayitedilismisYagsizSutMiktari + miktar;


                            xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());

                            xDegeri = Math.Round(xDegeri, 6);

                            yagsizSutMiktari = miktar;

                            sonuc = yagsizSutMiktari / xDegeri;

                            sonuc = Math.Round(sonuc, 6);

                            urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).ToList().ForEach(y => y.YagsizSut = sonuc);

                        }
                        else
                        {

                            xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.Miktar).FirstOrDefault().ToString());

                            xDegeri = Math.Round(xDegeri, 6);


                            //yagsizSutMiktari = DuzenlenenMiktar;

                            sonuc = DuzenlenenMiktar / xDegeri;

                            sonuc = Math.Round(sonuc, 6);

                            urunAgaciSonucus.Add(new UrunAgaciSonucu { UretimSiparisNo = Convert.ToInt32(UretimSiparisNo), YagsizSut = sonuc, UrunKodu = UrunKodu, UrunAdi = UrunAdi, YariMamul = YariMamulKodu });
                        }
                    }
                    #endregion

                }
                else
                {
                #region Alt ürün ağacı varsa
                tekrarsorgula:
                    if (AltbilesenKodu == "")
                    {
                        sql2 = "Select T1.\"Code\" as \"BilesenKodu\",T1.\"Quantity\" as Miktar,T0.\"Qauntity\" as BaslikMiktar  from OITT as T0 INNER JOIN ITT1 as T1 ON T0.[Code] = T1.[Father] INNER JOIN OITM as T2 ON T2.ItemCode = T1.Father where T0.\"Code\" = '" + bilesenKodu + "' and T2.\"ItmsGrpCod\" = '106'";

                    }
                    else
                    {
                        sql2 = "Select T1.\"Code\" as \"BilesenKodu\",T1.\"Quantity\" as Miktar,T0.\"Qauntity\" as BaslikMiktar  from OITT as T0 INNER JOIN ITT1 as T1 ON T0.[Code] = T1.[Father] INNER JOIN OITM as T2 ON T2.ItemCode = T1.Father where T0.\"Code\" = '" + AltbilesenKodu + "' and T2.\"ItmsGrpCod\" = '106'";
                    }

                    ConstVariables.oRecordset1 = (Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
                    ConstVariables.oRecordset1.DoQuery(sql2);

                    AltbilesenKodu = ConstVariables.oRecordset1.Fields.Item("BilesenKodu").Value.ToString();
                    miktar = parseNumber.parservalues<double>(ConstVariables.oRecordset1.Fields.Item("Miktar").Value.ToString());

                    if (AltbilesenKodu == "YRM-10003" || AltbilesenKodu == "YRM-10002")
                    {
                        string xmll = ConstVariables.oRecordset1.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                        XDocument xDoc = XDocument.Parse(xmll);
                        XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                        var rows = (from t in xDoc.Descendants(ns + "Row")
                                    select new
                                    {
                                        BilesenKodu = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "BilesenKodu" select new XElement(y.Element(ns + "Value"))).First().Value,
                                        Miktar = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "Miktar" select new XElement(y.Element(ns + "Value"))).First().Value),
                                        BaslikMiktar = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "BaslikMiktar" select new XElement(y.Element(ns + "Value"))).First().Value),
                                    }).ToList();


                        if (rows.Where(x => x.BilesenKodu == "YRM-10003").Count() > 0)
                        {
                            double yagliSutMiktari = 0;
                            double kayitedilismisYagliSutMiktari = 0;
                            double sonuc = 0;
                            if (urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Count() > 0)
                            //if (urunAgaciSonucus.Count > 0)
                            {
                                xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());

                                //yagliSutMiktari = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());
                                //kayitedilismisYagliSutMiktari = parseNumber.parservalues<double>(urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Select(y => y.YagliSut).FirstOrDefault().ToString());
                                //sonuc = kayitedilismisYagliSutMiktari + yagliSutMiktari;
                                xDegeri = Math.Round(xDegeri, 6);

                                sonuc = DuzenlenenMiktar / xDegeri;
                                urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).ToList().ForEach(y => y.YagliSut = sonuc);
                            }
                            else
                            {
                                xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());

                                xDegeri = Math.Round(xDegeri, 6);

                                //yagliSutMiktari = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());
                                sonuc = DuzenlenenMiktar / xDegeri;
                                urunAgaciSonucus.Add(new UrunAgaciSonucu { UretimSiparisNo = Convert.ToInt32(UretimSiparisNo), YagliSut = sonuc, UrunKodu = UrunKodu, UrunAdi = UrunAdi, YariMamul = YariMamulKodu });
                            }
                        }

                        if (rows.Where(x => x.BilesenKodu == "YRM-10002").Count() > 0)
                        {
                            double yagsizSutMiktari = 0;
                            double kayitedilismisYagsizSutMiktari = 0;
                            double sonuc = 0;
                            if (urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Count() > 0)
                            //if (urunAgaciSonucus.Count > 0)
                            {

                                xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.Miktar).FirstOrDefault().ToString());

                                yagsizSutMiktari = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.Miktar).FirstOrDefault().ToString());
                                //kayitedilismisYagsizSutMiktari = parseNumber.parservalues<double>(urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Select(y => y.YagsizSut).FirstOrDefault().ToString());
                                //sonuc = kayitedilismisYagsizSutMiktari + yagsizSutMiktari;
                                xDegeri = Math.Round(xDegeri, 6);

                                sonuc = DuzenlenenMiktar / xDegeri;

                                urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).ToList().ForEach(y => y.YagsizSut = sonuc);

                                //yagsizSutMiktari = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.Miktar).FirstOrDefault().ToString());

                                //urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).ToList().ForEach(y => y.YagsizSut = y.YagsizSut + yagsizSutMiktari);

                            }
                            else
                            {
                                xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.Miktar).FirstOrDefault().ToString());

                                xDegeri = Math.Round(xDegeri, 6);

                                //yagsizSutMiktari = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.Miktar).FirstOrDefault().ToString());
                                sonuc = DuzenlenenMiktar / xDegeri;

                                urunAgaciSonucus.Add(new UrunAgaciSonucu { UretimSiparisNo = Convert.ToInt32(UretimSiparisNo), YagsizSut = sonuc, UrunKodu = UrunKodu, UrunAdi = UrunAdi, YariMamul = YariMamulKodu });
                            }
                        }

                    }
                    else
                    {
                        if (AltbilesenKodu != "")
                        {
                            ConstVariables.oRecordset1.MoveNext();
                            goto tekrarsorgula;
                        }
                    }

                    #endregion
                }

                ConstVariables.oRecordset.MoveNext();
            }

            return urunAgaciSonucus;
        }


        #region Eski hesaplama
        //private List<UrunAgaciSonucu> UrunAgaciIslemleri(string YariMamulKodu, string UretimSiparisNo, string UrunKodu, string UrunAdi)
        //{
        //    string bilesenKodu = "";
        //    string AltbilesenKodu = "";
        //    double miktar = 0;
        //    double baslikmiktar = 0;
        //    double xDegeri = 0;
        //    string sql = "Select T1.\"Code\" as \"BilesenKodu\",T1.\"Quantity\" as Miktar,T0.\"Qauntity\" as BaslikMiktar from OITT as T0 INNER JOIN ITT1 as T1 ON T0.[Code] = T1.[Father] INNER JOIN OITM as T2 ON T2.ItemCode = T1.Father where T0.\"Code\" = '" + YariMamulKodu + "' and T2.\"ItmsGrpCod\" = '106' and T1.Code like 'YRM%'";

        //    ConstVariables.oRecordset = (Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
        //    ConstVariables.oRecordset.DoQuery(sql);

        //    string sql2 = "";
        //    while (!ConstVariables.oRecordset.EoF)
        //    {
        //        miktar = parseNumber.parservalues<double>(ConstVariables.oRecordset.Fields.Item("Miktar").Value.ToString());
        //        bilesenKodu = ConstVariables.oRecordset.Fields.Item("BilesenKodu").Value.ToString();
        //        if (bilesenKodu == "YRM-10003" || bilesenKodu == "YRM-10002")
        //        {

        //            #region İlk bileşende bulunursa.
        //            string xmll = ConstVariables.oRecordset1.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
        //            XDocument xDoc = XDocument.Parse(xmll);
        //            XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
        //            var rows = (from t in xDoc.Descendants(ns + "Row")
        //                        select new
        //                        {
        //                            BilesenKodu = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "BilesenKodu" select new XElement(y.Element(ns + "Value"))).First().Value,
        //                            Miktar = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "Miktar" select new XElement(y.Element(ns + "Value"))).First().Value),
        //                            BaslikMiktar = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "BaslikMiktar" select new XElement(y.Element(ns + "Value"))).First().Value),
        //                        }).ToList();


        //            if (rows.Where(x => x.BilesenKodu == "YRM-10003").Count() > 0 && bilesenKodu == "YRM-10003")
        //            {
        //                double yagliSutMiktari = 0;
        //                double kayitedilismisYagliSutMiktari = 0;
        //                double sonuc = 0;
        //                if (urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Count() > 0)
        //                //if (urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Count() > 0)
        //                {

        //                    xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());

        //                    xDegeri = Math.Round(xDegeri, 6);

        //                    yagliSutMiktari = miktar;
        //                    //kayitedilismisYagliSutMiktari = parseNumber.parservalues<double>(urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Select(y => y.YagliSut).FirstOrDefault().ToString());
        //                    //sonuc = kayitedilismisYagliSutMiktari + miktar;

        //                    sonuc = yagliSutMiktari / xDegeri;


        //                    urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).ToList().ForEach(y => y.YagliSut = sonuc);

        //                    //yagliSutMiktari = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());

        //                    //urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).ToList().ForEach(y => y.YagliSut = y.YagliSut + yagliSutMiktari);
        //                }
        //                else
        //                {
        //                    xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());

        //                    xDegeri = Math.Round(xDegeri, 6);

        //                    yagliSutMiktari = miktar;
        //                    sonuc = yagliSutMiktari / xDegeri;

        //                    urunAgaciSonucus.Add(new UrunAgaciSonucu { UretimSiparisNo = Convert.ToInt32(UretimSiparisNo), YagliSut = sonuc, UrunKodu = UrunKodu, UrunAdi = UrunAdi, YariMamul = YariMamulKodu });
        //                }
        //            }
        //            else if (rows.Where(x => x.BilesenKodu == "YRM-10002").Count() > 0 && bilesenKodu == "YRM-10002")
        //            {
        //                double yagsizSutMiktari = 0;
        //                double kayitedilismisYagsizSutMiktari = 0;
        //                double sonuc = 0;
        //                if (urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Count() > 0)
        //                //if (urunAgaciSonucus.Count > 0)
        //                {
        //                    yagsizSutMiktari = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.Miktar).FirstOrDefault().ToString());
        //                    //kayitedilismisYagsizSutMiktari = parseNumber.parservalues<double>(urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Select(y => y.YagsizSut).FirstOrDefault().ToString());
        //                    //sonuc = kayitedilismisYagsizSutMiktari + miktar;


        //                    xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());

        //                    xDegeri = Math.Round(xDegeri, 6);

        //                    yagsizSutMiktari = miktar;

        //                    sonuc = yagsizSutMiktari / xDegeri;


        //                    urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).ToList().ForEach(y => y.YagsizSut = sonuc);

        //                }
        //                else
        //                {

        //                    xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.Miktar).FirstOrDefault().ToString());

        //                    xDegeri = Math.Round(xDegeri, 6);

        //                    yagsizSutMiktari = miktar;
        //                    sonuc = yagsizSutMiktari / xDegeri;

        //                    sonuc = Math.Round(sonuc, 6);

        //                    urunAgaciSonucus.Add(new UrunAgaciSonucu { UretimSiparisNo = Convert.ToInt32(UretimSiparisNo), YagsizSut = sonuc, UrunKodu = UrunKodu, UrunAdi = UrunAdi, YariMamul = YariMamulKodu });
        //                }
        //            } 
        //            #endregion

        //        }
        //        else
        //        {
        //            #region Alt ürün ağacı varsa
        //            tekrarsorgula:
        //            if (AltbilesenKodu == "")
        //            {
        //                sql2 = "Select T1.\"Code\" as \"BilesenKodu\",T1.\"Quantity\" as Miktar,T0.\"Qauntity\" as BaslikMiktar  from OITT as T0 INNER JOIN ITT1 as T1 ON T0.[Code] = T1.[Father] INNER JOIN OITM as T2 ON T2.ItemCode = T1.Father where T0.\"Code\" = '" + bilesenKodu + "' and T2.\"ItmsGrpCod\" = '106'";

        //            }
        //            else
        //            {
        //                sql2 = "Select T1.\"Code\" as \"BilesenKodu\",T1.\"Quantity\" as Miktar,T0.\"Qauntity\" as BaslikMiktar  from OITT as T0 INNER JOIN ITT1 as T1 ON T0.[Code] = T1.[Father] INNER JOIN OITM as T2 ON T2.ItemCode = T1.Father where T0.\"Code\" = '" + AltbilesenKodu + "' and T2.\"ItmsGrpCod\" = '106'";
        //            }

        //            ConstVariables.oRecordset1 = (Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
        //            ConstVariables.oRecordset1.DoQuery(sql2);

        //            AltbilesenKodu = ConstVariables.oRecordset1.Fields.Item("BilesenKodu").Value.ToString();
        //            miktar = parseNumber.parservalues<double>(ConstVariables.oRecordset1.Fields.Item("Miktar").Value.ToString());

        //            if (AltbilesenKodu == "YRM-10003" || AltbilesenKodu == "YRM-10002")
        //            {
        //                string xmll = ConstVariables.oRecordset1.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
        //                XDocument xDoc = XDocument.Parse(xmll);
        //                XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
        //                var rows = (from t in xDoc.Descendants(ns + "Row")
        //                            select new
        //                            {
        //                                BilesenKodu = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "BilesenKodu" select new XElement(y.Element(ns + "Value"))).First().Value,
        //                                Miktar = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "Miktar" select new XElement(y.Element(ns + "Value"))).First().Value),
        //                                BaslikMiktar = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "BaslikMiktar" select new XElement(y.Element(ns + "Value"))).First().Value),
        //                            }).ToList();


        //                if (rows.Where(x => x.BilesenKodu == "YRM-10003").Count() > 0)
        //                {
        //                    double yagliSutMiktari = 0;
        //                    double kayitedilismisYagliSutMiktari = 0;
        //                    double sonuc = 0;
        //                    if (urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Count() > 0)
        //                    //if (urunAgaciSonucus.Count > 0)
        //                    {
        //                        xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());

        //                        yagliSutMiktari = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());
        //                        //kayitedilismisYagliSutMiktari = parseNumber.parservalues<double>(urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Select(y => y.YagliSut).FirstOrDefault().ToString());
        //                        //sonuc = kayitedilismisYagliSutMiktari + yagliSutMiktari;
        //                        xDegeri = Math.Round(xDegeri, 6);

        //                        sonuc = yagliSutMiktari / xDegeri;
        //                        urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).ToList().ForEach(y => y.YagliSut = sonuc);
        //                    }
        //                    else
        //                    {
        //                        xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());

        //                        xDegeri = Math.Round(xDegeri, 6);

        //                        yagliSutMiktari = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10003").Select(y => y.Miktar).FirstOrDefault().ToString());
        //                        sonuc = yagliSutMiktari / xDegeri;
        //                        urunAgaciSonucus.Add(new UrunAgaciSonucu { UretimSiparisNo = Convert.ToInt32(UretimSiparisNo), YagliSut = sonuc, UrunKodu = UrunKodu, UrunAdi = UrunAdi, YariMamul = YariMamulKodu });
        //                    }
        //                }

        //                if (rows.Where(x => x.BilesenKodu == "YRM-10002").Count() > 0)
        //                {
        //                    double yagsizSutMiktari = 0;
        //                    double kayitedilismisYagsizSutMiktari = 0;
        //                    double sonuc = 0;
        //                    if (urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Count() > 0)
        //                    //if (urunAgaciSonucus.Count > 0)
        //                    {

        //                        xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.Miktar).FirstOrDefault().ToString());

        //                        yagsizSutMiktari = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.Miktar).FirstOrDefault().ToString());
        //                        //kayitedilismisYagsizSutMiktari = parseNumber.parservalues<double>(urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).Select(y => y.YagsizSut).FirstOrDefault().ToString());
        //                        //sonuc = kayitedilismisYagsizSutMiktari + yagsizSutMiktari;
        //                        xDegeri = Math.Round(xDegeri, 6);

        //                        sonuc = yagsizSutMiktari / xDegeri;

        //                        urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).ToList().ForEach(y => y.YagsizSut = sonuc);

        //                        //yagsizSutMiktari = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.Miktar).FirstOrDefault().ToString());

        //                        //urunAgaciSonucus.Where(x => x.UretimSiparisNo == Convert.ToInt32(UretimSiparisNo)).ToList().ForEach(y => y.YagsizSut = y.YagsizSut + yagsizSutMiktari);

        //                    }
        //                    else
        //                    {
        //                        xDegeri = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.BaslikMiktar).FirstOrDefault().ToString()) / parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.Miktar).FirstOrDefault().ToString());

        //                        yagsizSutMiktari = parseNumber.parservalues<double>(rows.Where(x => x.BilesenKodu == "YRM-10002").Select(y => y.Miktar).FirstOrDefault().ToString());
        //                        sonuc = yagsizSutMiktari / xDegeri;
        //                        urunAgaciSonucus.Add(new UrunAgaciSonucu { UretimSiparisNo = Convert.ToInt32(UretimSiparisNo), YagsizSut = sonuc, UrunKodu = UrunKodu, UrunAdi = UrunAdi, YariMamul = YariMamulKodu });
        //                    }
        //                }

        //            }
        //            else
        //            {
        //                ConstVariables.oRecordset1.MoveNext();
        //                goto tekrarsorgula;
        //            }

        //            #endregion
        //        }

        //        ConstVariables.oRecordset.MoveNext();
        //    }

        //    return urunAgaciSonucus;
        //} 
        #endregion

        private void Listele(string baslangicTarihi, string bitisTarihi)
        {
            urunAgaciSonucus = new List<UrunAgaciSonucu>();
            string sql = "Select MAX(tbl.[Üretim Sipariş No]) as [Üretim Sipariş No],MAX(tbl.[Ana Ürün]) as [Ana Ürün],MAX(tbl.[Ana Ürün Adı]) as [Ana Ürün Adı],cast(0 as decimal(15, 2)) as \"Yağlı Süt\", cast(0 as decimal(15, 2)) as \"Yağsız Süt\",tbl.\"Yarı Mamül\",SUM(T4.Quantity) as \"Düzenlenen Miktar\" ";
            sql += " from (SELECT T0.DocEntry AS \"Üretim Sipariş No\",T0.ItemCode AS \"Ana Ürün\",T0.ProdName as \"Ana Ürün Adı\",T2.\"ItemCode\" as \"Yarı Mamül\" FROM OWOR AS T0 INNER JOIN OITM AS T1 ON T0.ItemCode = T1.ItemCode ";
            sql += "INNER JOIN WOR1 as T2 ON T0.DocEntry = T2.DocEntry";


            sql += " WHERE 1=1";

            if (baslangicTarihi != "")
            {

                sql += " and Convert(varchar, T0.StartDate, 112) >= '" + baslangicTarihi + "' ";
            }

            if (bitisTarihi != "")
            {
                sql += " and Convert(varchar, T0.StartDate, 112) <= '" + bitisTarihi + "' ";
            }
            sql += " AND T1.ItmsGrpCod = '107') as tbl INNER JOIN OITM as T3 ON tbl.[Yarı Mamül] = T3.ItemCode ";
            sql += "LEFT JOIN IGE1 AS T4 ON tbl.[Üretim Sipariş No] = T4.BaseRef and tbl.[Yarı Mamül] = T4.ItemCode";
            sql += " where T3.ItmsGrpCod = '106' ";
            sql += " group by Tbl.[Yarı Mamül]";

            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRS.DoQuery(sql);


            while (!oRS.EoF)
            {

                UrunAgaciIslemleri(oRS.Fields.Item("Yarı Mamül").Value.ToString(), oRS.Fields.Item("Üretim Sipariş No").Value.ToString(), oRS.Fields.Item("Ana Ürün").Value.ToString(), oRS.Fields.Item("Ana Ürün Adı").Value.ToString(), parseNumber.parservalues<double>(oRS.Fields.Item("Düzenlenen Miktar").Value.ToString()));
                oRS.MoveNext();
            }





            string data2 = string.Join("", urunAgaciSonucus.Select(s => string.Format(row, s.UretimSiparisNo, s.UrunKodu, s.UrunAdi, s.YagliSut, s.YagsizSut, s.YariMamul)));


            var final = string.Format(header, data2);

            oDataTable.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_All, final);



            //oDataTable.Clear();
            //oDataTable.ExecuteQuery(sql);

            oMatrix.Clear();

            oMatrix.Columns.Item("Col_4").DataBind.Bind("DATA", "SipNo");
            oMatrix.Columns.Item("Col_0").DataBind.Bind("DATA", "UrunKodu");
            oMatrix.Columns.Item("Col_1").DataBind.Bind("DATA", "UrunAdi");
            oMatrix.Columns.Item("Col_2").DataBind.Bind("DATA", "YagliSut");
            oMatrix.Columns.Item("Col_3").DataBind.Bind("DATA", "YagsizSut");
            oMatrix.Columns.Item("Col_5").DataBind.Bind("DATA", "YariMamul");

            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();
        }
    }
}
