using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
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
    public class SatinalmaIskontoUrunEkle
    {
        [ItemAtt(AIFConn.SatinalmaIskontoluUrunEkleUID)]
        public static SAPbouiCOM.Form frmSatinalmaIskontoUrunEkle;

        //[ItemAtt("Item_1")]
        //public SAPbouiCOM.ComboBox oComboCalisan;

        [ItemAtt("Item_3")]
        public SAPbouiCOM.Matrix oMatrix;

        [ItemAtt("Item_9")]
        public SAPbouiCOM.EditText oEditAranacak;

        [ItemAtt("Item_5")]
        public SAPbouiCOM.EditText oEditMiktarOlcegi;

        [ItemAtt("Item_12")]
        public SAPbouiCOM.EditText oEditBirinciIsk;

        [ItemAtt("Item_14")]
        public SAPbouiCOM.EditText oEditIkinciIsk;

        [ItemAtt("Item_16")]
        public SAPbouiCOM.EditText oEditUcuncuIsk;

        [ItemAtt("Item_18")]
        public SAPbouiCOM.EditText oEditDorduncuIsk;

        [ItemAtt("Item_20")]
        public SAPbouiCOM.EditText oEditBesinciIsk;

        [ItemAtt("Item_22")]
        public SAPbouiCOM.EditText oEditToplamIsk;

        [ItemAtt("Item_0")]
        public SAPbouiCOM.Button oBtnTumunuSec;

        [ItemAtt("Item_1")]
        public SAPbouiCOM.Button oBtnTumunuKaldir;

        [ItemAtt("Item_10")]
        public SAPbouiCOM.Button oBtnAra;

        private string datatableHeader = "<?xml version=\"1.0\" encoding=\"UTF-16\" ?><DataTable Uid=\"DATA\"><Columns><Column Uid=\"Sec\" Type=\"1\" MaxLength=\"1\"/><Column Uid=\"UrunKodu\" Type=\"1\" MaxLength=\"30\"/><Column Uid=\"UrunTanimi\" Type=\"1\" MaxLength=\"150\"/><Column Uid=\"UrunGrupKod\" Type=\"1\" MaxLength=\"100\"/><Column Uid=\"UrunGrupAdi\" Type=\"1\" MaxLength=\"100\"/><Column Uid=\"KalemGrubu2Kodu\" Type=\"1\" MaxLength=\"10\"/><Column Uid=\"KalemGrubu2Adi\" Type=\"1\" MaxLength=\"100\"/><Column Uid=\"KalemGrubu3Kodu\" Type=\"1\" MaxLength=\"10\"/><Column Uid=\"KalemGrubu3Adi\" Type=\"1\" MaxLength=\"100\"/><Column Uid=\"KalemGrubu4Kodu\" Type=\"1\" MaxLength=\"10\"/><Column Uid=\"KalemGrubu4Adi\" Type=\"1\" MaxLength=\"100\"/></Columns><Rows>{0}</Rows></DataTable>";


        private string row = "<Row><Cells><Cell><ColumnUid>Sec</ColumnUid><Value>{0}</Value></Cell><Cell><ColumnUid>UrunKodu</ColumnUid><Value>{1}</Value></Cell><Cell><ColumnUid>UrunTanimi</ColumnUid><Value>{2}</Value></Cell><Cell><ColumnUid>UrunGrupKod</ColumnUid><Value>{3}</Value></Cell><Cell><ColumnUid>UrunGrupAdi</ColumnUid><Value>{4}</Value></Cell><Cell><ColumnUid>KalemGrubu2Kodu</ColumnUid><Value>{5}</Value></Cell><Cell><ColumnUid>KalemGrubu2Adi</ColumnUid><Value>{6}</Value></Cell><Cell><ColumnUid>KalemGrubu3Kodu</ColumnUid><Value>{7}</Value></Cell><Cell><ColumnUid>KalemGrubu3Adi</ColumnUid><Value>{8}</Value></Cell><Cell><ColumnUid>KalemGrubu4Kodu</ColumnUid><Value>{9}</Value></Cell><Cell><ColumnUid>KalemGrubu4Adi</ColumnUid><Value>{10}</Value></Cell></Cells></Row>";

        string guncellemeModu = "";
        public void LoadForms(string _guncellemeModu = "")
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.SatinalmaIskontoluUrunEkleFrmXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.SatinalmaIskontoluUrunEkleFrmXML));
            Functions.CreateUserOrSystemFormComponent<SatinalmaIskontoUrunEkle>(AIFConn.StnAlmUrn);
            guncellemeModu = _guncellemeModu;

            if (guncellemeModu != "")
            {
                frmSatinalmaIskontoUrunEkle.Title = "Satınalma Iskontolu Ürün Güncelle";
            }
            InitForms();
        }
        string sql = "";
        private SAPbouiCOM.DataTable oDataTable = null;
        public string StringReplace(string text)
        {
            text = text.Replace("İ", "I");
            text = text.Replace("ı", "i");
            text = text.Replace("Ğ", "G");
            text = text.Replace("ğ", "g");
            text = text.Replace("Ö", "O");
            text = text.Replace("ö", "o");
            text = text.Replace("Ü", "U");
            text = text.Replace("ü", "u");
            text = text.Replace("Ş", "S");
            text = text.Replace("ş", "s");
            text = text.Replace("Ç", "C");
            text = text.Replace("ç", "c");
            return text;
        }
        public void InitForms()
        {
            try
            {
                frmSatinalmaIskontoUrunEkle.Freeze(true);
                oDataTable = frmSatinalmaIskontoUrunEkle.DataSources.DataTables.Add("DATA");
                //oMatrix.AddRow();

                frmSatinalmaIskontoUrunEkle.EnableMenu("1283", false);
                frmSatinalmaIskontoUrunEkle.EnableMenu("1284", false);
                frmSatinalmaIskontoUrunEkle.EnableMenu("1286", false);


                if (guncellemeModu == "")
                {
                    #region Sistemdeki tüm kalemler getirilir.
                    string sql = "Select T0.\"ItemCode\",T0.\"ItemName\",T0.\"ItmsGrpCod\",T1.\"ItmsGrpNam\"";

                    sql += ",T0.\"U_ItemGroup2\",(select T98.\"Descr\" from CUFD as T99 INNER JOIN UFD1 as T98 ON T99.\"FieldID\" = T98.\"FieldID\" where T99.\"TableID\" = 'OITM' and T99.\"AliasID\" = 'ItemGroup2' and T98.\"TableID\" = 'OITM' and T98.\"FldValue\" = T0.\"U_ItemGroup2\") as \"KalemGrubu2Adi\" ";
                    sql += ",T0.\"U_ItemGroup3\",(select T98.\"Descr\" from CUFD as T99 INNER JOIN UFD1 as T98 ON T99.\"FieldID\" = T98.\"FieldID\" where T99.\"TableID\" = 'OITM' and T99.\"AliasID\" = 'ItemGroup2' and T98.\"TableID\" = 'OITM' and T98.\"FldValue\" = T0.\"U_ItemGroup3\") as \"KalemGrubu3Adi\" ";
                    sql += ",T0.\"U_ItemGroup4\",(select T98.\"Descr\" from CUFD as T99 INNER JOIN UFD1 as T98 ON T99.\"FieldID\" = T98.\"FieldID\" where T99.\"TableID\" = 'OITM' and T99.\"AliasID\" = 'ItemGroup2' and T98.\"TableID\" = 'OITM' and T98.\"FldValue\" = T0.\"U_ItemGroup4\") as \"KalemGrubu4Adi\"";

                    sql += " from \"OITM\" as T0 INNER JOIN OITB T1 ON T0.\"ItmsGrpCod\" = T1.\"ItmsGrpCod\" order by T0.\"ItemCode\"";
                    ConstVariables.oRecordset.DoQuery(sql);


                    string xmll = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                    XDocument xDoc = XDocument.Parse(xmll);
                    XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                    urunEkles = (from t in xDoc.Descendants(ns + "Row")
                                 select new SatinalmaIskontolu_UrunEkle()
                                 {
                                     Sec = "N",
                                     UrunKodu = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "ItemCode" select new XElement(y.Element(ns + "Value"))).First().Value),
                                     UrunTanimi = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "ItemName" select new XElement(y.Element(ns + "Value"))).First().Value),
                                     UrunGrupKod = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "ItmsGrpCod" select new XElement(y.Element(ns + "Value"))).First().Value),
                                     UrunGrupAdi = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "ItmsGrpNam" select new XElement(y.Element(ns + "Value"))).First().Value),
                                     KalemGrubu2Kodu = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_ItemGroup2" select new XElement(y.Element(ns + "Value"))).First().Value),
                                     KalemGrubu2Adi = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "KalemGrubu2Adi" select new XElement(y.Element(ns + "Value"))).First().Value),
                                     KalemGrubu3Kodu = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_ItemGroup3" select new XElement(y.Element(ns + "Value"))).First().Value),
                                     KalemGrubu3Adi = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "KalemGrubu3Adi" select new XElement(y.Element(ns + "Value"))).First().Value),
                                     KalemGrubu4Kodu = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_ItemGroup4" select new XElement(y.Element(ns + "Value"))).First().Value),
                                     KalemGrubu4Adi = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "KalemGrubu4Adi" select new XElement(y.Element(ns + "Value"))).First().Value),
                                 }).ToList();
                    #endregion
                }
                else if (guncellemeModu != "")
                {
                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)SatinalmaIskontoGiris.frmSatinalmaIskontoGiris.Items.Item("Item_8").Specific;


                    #region Sistemdeki tüm kalemler alınır.
                    string sql = "Select T0.\"ItemCode\",T0.\"ItemName\",T0.\"ItmsGrpCod\",T1.\"ItmsGrpNam\"";

                    sql += ",T0.\"U_ItemGroup2\",(select T98.\"Descr\" from CUFD as T99 INNER JOIN UFD1 as T98 ON T99.\"FieldID\" = T98.\"FieldID\" where T99.\"TableID\" = 'OITM' and T99.\"AliasID\" = 'ItemGroup2' and T98.\"TableID\" = 'OITM' and T98.\"FldValue\" = T0.\"U_ItemGroup2\") as \"KalemGrubu2Adi\" ";
                    sql += ",T0.\"U_ItemGroup3\",(select T98.\"Descr\" from CUFD as T99 INNER JOIN UFD1 as T98 ON T99.\"FieldID\" = T98.\"FieldID\" where T99.\"TableID\" = 'OITM' and T99.\"AliasID\" = 'ItemGroup2' and T98.\"TableID\" = 'OITM' and T98.\"FldValue\" = T0.\"U_ItemGroup3\") as \"KalemGrubu3Adi\" ";
                    sql += ",T0.\"U_ItemGroup4\",(select T98.\"Descr\" from CUFD as T99 INNER JOIN UFD1 as T98 ON T99.\"FieldID\" = T98.\"FieldID\" where T99.\"TableID\" = 'OITM' and T99.\"AliasID\" = 'ItemGroup2' and T98.\"TableID\" = 'OITM' and T98.\"FldValue\" = T0.\"U_ItemGroup4\") as \"KalemGrubu4Adi\"";

                    sql += " from \"OITM\" as T0 INNER JOIN OITB T1 ON T0.\"ItmsGrpCod\" = T1.\"ItmsGrpCod\" order by T0.\"ItemCode\"";
                    ConstVariables.oRecordset.DoQuery(sql);


                    string xmll = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                    XDocument xDoc = XDocument.Parse(xmll);
                    XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                    var oitmRows = (from t in xDoc.Descendants(ns + "Row")
                                    select new UrunEkle()
                                    {
                                        Sec = "N",
                                        UrunKodu = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "ItemCode" select new XElement(y.Element(ns + "Value"))).First().Value),
                                        UrunTanimi = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "ItemName" select new XElement(y.Element(ns + "Value"))).First().Value),
                                        UrunGrupKod = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "ItmsGrpCod" select new XElement(y.Element(ns + "Value"))).First().Value),
                                        UrunGrupAdi = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "ItmsGrpNam" select new XElement(y.Element(ns + "Value"))).First().Value),
                                        KalemGrubu2Kodu = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_ItemGroup2" select new XElement(y.Element(ns + "Value"))).First().Value),
                                        KalemGrubu2Adi = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "KalemGrubu2Adi" select new XElement(y.Element(ns + "Value"))).First().Value),
                                        KalemGrubu3Kodu = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_ItemGroup3" select new XElement(y.Element(ns + "Value"))).First().Value),
                                        KalemGrubu3Adi = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "KalemGrubu3Adi" select new XElement(y.Element(ns + "Value"))).First().Value),
                                        KalemGrubu4Kodu = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_ItemGroup4" select new XElement(y.Element(ns + "Value"))).First().Value),
                                        KalemGrubu4Adi = StringReplace((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "KalemGrubu4Adi" select new XElement(y.Element(ns + "Value"))).First().Value),
                                    }).ToList();
                    #endregion


                    #region Sistemdeki çekilen tüm ürünlerden alınan verilerle birlikte ortak listeye eklenir.
                    var xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                    var rowsUrunler = (from x in XDocument.Parse(xml).Descendants("Row")
                                       select new SatinalmaIskontolu_UrunEkle()
                                       {
                                           UrunKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                           UrunTanimi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                                           birinciIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value,
                                           ikinciIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value,
                                           ucuncuIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,
                                           dorduncuIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                                           besinciIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value,
                                           toplamIskonto = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value,
                                           KalemGrubu2Kodu = oitmRows.Where(z => z.UrunKodu == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.KalemGrubu2Kodu).FirstOrDefault(),
                                           KalemGrubu2Adi = oitmRows.Where(z => z.UrunKodu == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.KalemGrubu2Adi).FirstOrDefault(),
                                           KalemGrubu3Kodu = oitmRows.Where(z => z.UrunKodu == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.KalemGrubu3Kodu).FirstOrDefault(),
                                           KalemGrubu3Adi = oitmRows.Where(z => z.UrunKodu == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.KalemGrubu3Adi).FirstOrDefault(),
                                           KalemGrubu4Kodu = oitmRows.Where(z => z.UrunKodu == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.KalemGrubu4Kodu).FirstOrDefault(),
                                           KalemGrubu4Adi = oitmRows.Where(z => z.UrunKodu == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.KalemGrubu4Adi).FirstOrDefault(),
                                           UrunGrupKod = oitmRows.Where(z => z.UrunKodu == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.UrunGrupKod).FirstOrDefault(),
                                           UrunGrupAdi = oitmRows.Where(z => z.UrunKodu == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.UrunGrupAdi).FirstOrDefault(),

                                       }).ToList();

                    urunEkles.Clear();

                    urunEkles.AddRange(rowsUrunler);
                    #endregion


                }

                string rowsFinal = string.Join("", urunEkles.Select(y => string.Format(row, y.Sec, y.UrunKodu, y.UrunTanimi, y.UrunGrupKod, y.UrunGrupAdi, y.KalemGrubu2Kodu, y.KalemGrubu2Adi, y.KalemGrubu3Kodu, y.KalemGrubu3Adi, y.KalemGrubu4Kodu, y.KalemGrubu4Adi)));

                string data = string.Format(datatableHeader, rowsFinal);


                oDataTable.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_All, data);

                oMatrix.Columns.Item("Col_4").DataBind.Bind("DATA", "Sec");
                oMatrix.Columns.Item("Col_0").DataBind.Bind("DATA", "UrunKodu");
                oMatrix.Columns.Item("Col_1").DataBind.Bind("DATA", "UrunTanimi");
                oMatrix.Columns.Item("Col_2").DataBind.Bind("DATA", "UrunGrupKod");
                oMatrix.Columns.Item("Col_3").DataBind.Bind("DATA", "UrunGrupAdi");
                oMatrix.Columns.Item("Col_5").DataBind.Bind("DATA", "KalemGrubu2Kodu");
                oMatrix.Columns.Item("Col_6").DataBind.Bind("DATA", "KalemGrubu2Adi");
                oMatrix.Columns.Item("Col_7").DataBind.Bind("DATA", "KalemGrubu3Kodu");
                oMatrix.Columns.Item("Col_8").DataBind.Bind("DATA", "KalemGrubu3Adi");
                oMatrix.Columns.Item("Col_9").DataBind.Bind("DATA", "KalemGrubu4Kodu");
                oMatrix.Columns.Item("Col_10").DataBind.Bind("DATA", "KalemGrubu4Adi");

                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();


            }
            catch (Exception ex)
            {
            }
            finally
            {
                frmSatinalmaIskontoUrunEkle.Freeze(false);
            }
        }
        List<SatinalmaIskontolu_UrunEkle> urunEkles = new List<SatinalmaIskontolu_UrunEkle>();
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
                    if (pVal.ColUID == "Col_4" && !pVal.BeforeAction)
                    {
                        string itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value;
                        string secim = ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific).Checked == true ? "Y" : "N";

                        urunEkles.Where(x => x.UrunKodu == itemCode).ToList().ForEach(y => y.Sec = secim);
                    }
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
                    if (!pVal.BeforeAction && pVal.ItemUID == "Item_10")
                    {
                        //AIFConn.IndUrunGrp.LoadForms((SAPbouiCOM.Form)frmIndirimUrunEkle);
                        try
                        {
                            frmSatinalmaIskontoUrunEkle.Freeze(true);
                            string aranacakkelime = StringReplace(oEditAranacak.Value.ToUpper());
                            List<SatinalmaIskontolu_UrunEkle> urunEkleArama = null;

                            if (aranacakkelime == "")
                            {
                                urunEkleArama = urunEkles;
                            }
                            else
                            {
                                urunEkleArama = urunEkles.Where(x => x.UrunKodu.Contains(aranacakkelime) || x.UrunTanimi.Contains(aranacakkelime) || x.UrunKodu.Contains(aranacakkelime) || x.UrunGrupAdi.Contains(aranacakkelime) || x.KalemGrubu2Adi.Contains(aranacakkelime) || x.KalemGrubu2Kodu.Contains(aranacakkelime) || x.KalemGrubu3Kodu.Contains(aranacakkelime) || x.KalemGrubu3Adi.Contains(aranacakkelime) || x.KalemGrubu4Kodu.Contains(aranacakkelime) || x.KalemGrubu4Adi.Contains(aranacakkelime)).ToList();
                            }

                            string rowsFinal = string.Join("", urunEkleArama.Select(y => string.Format(row, y.Sec, y.UrunKodu, y.UrunTanimi, y.UrunGrupKod, y.UrunGrupAdi, y.KalemGrubu2Kodu, y.KalemGrubu2Adi, y.KalemGrubu3Kodu, y.KalemGrubu3Adi, y.KalemGrubu4Kodu, y.KalemGrubu4Adi)));

                            string data = string.Format(datatableHeader, rowsFinal);

                            oDataTable.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_All, data);

                            var xml = oDataTable.SerializeAsXML(BoDataTableXmlSelect.dxs_All);


                            oMatrix.Columns.Item("Col_4").DataBind.Bind("DATA", "Sec");
                            oMatrix.Columns.Item("Col_0").DataBind.Bind("DATA", "UrunKodu");
                            oMatrix.Columns.Item("Col_1").DataBind.Bind("DATA", "UrunTanimi");
                            oMatrix.Columns.Item("Col_2").DataBind.Bind("DATA", "UrunGrupKod");
                            oMatrix.Columns.Item("Col_3").DataBind.Bind("DATA", "UrunGrupAdi");
                            oMatrix.Columns.Item("Col_5").DataBind.Bind("DATA", "KalemGrubu2Kodu");
                            oMatrix.Columns.Item("Col_6").DataBind.Bind("DATA", "KalemGrubu2Adi");
                            oMatrix.Columns.Item("Col_7").DataBind.Bind("DATA", "KalemGrubu3Kodu");
                            oMatrix.Columns.Item("Col_8").DataBind.Bind("DATA", "KalemGrubu3Adi");
                            oMatrix.Columns.Item("Col_9").DataBind.Bind("DATA", "KalemGrubu4Kodu");
                            oMatrix.Columns.Item("Col_10").DataBind.Bind("DATA", "KalemGrubu4Adi");

                            oMatrix.LoadFromDataSource();
                            oMatrix.AutoResizeColumns();

                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmSatinalmaIskontoUrunEkle.Freeze(false);
                        }

                    }
                    else if (pVal.ItemUID == "Item_3" && pVal.ColUID == "Col_4" && pVal.BeforeAction && pVal.Modifiers == BoModifiersEnum.mt_SHIFT)
                    {
                        BubbleEvent = false;
                        //var xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                        //var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                        //            where string.Equals((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value, "Y")
                        //            select new
                        //            {
                        //                UrunKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,

                        //            }).ToList();


                        //foreach (var item in rows)
                        //{
                        //    urunEkles.Where(x => x.UrunKodu == item.UrunKodu).ToList().ForEach(y => y.Sec = "Y");
                        //}
                    }
                    else if (!pVal.BeforeAction && pVal.ItemUID == "Item_4")
                    {
                        List<SatinalmaIskontolu_UrunEkle> listeGonder = urunEkles.Where(c => c.Sec == "Y").ToList();

                        //string _MiktarOlcgei = Convert.ToDouble(oEditMiktarOlcegi.Value, System.Globalization.CultureInfo.InvariantCulture);
                        string _MiktarOlcgei = oEditMiktarOlcegi.Value == null ? Convert.ToDouble("0").ToString() : oEditMiktarOlcegi.Value;
                        string _birinciIsk = oEditBirinciIsk.Value == null ? Convert.ToDouble("0").ToString() : oEditBirinciIsk.Value;
                        string _ikinciIsk = oEditIkinciIsk.Value == null ? Convert.ToDouble("0").ToString() : oEditIkinciIsk.Value;
                        string _ucuncuIsk = oEditUcuncuIsk.Value == null ? Convert.ToDouble("0").ToString() : oEditUcuncuIsk.Value;
                        string _dorduncuIsk = oEditDorduncuIsk.Value == null ? Convert.ToDouble("0").ToString() : oEditDorduncuIsk.Value;
                        string _besinciIsk = oEditBesinciIsk.Value == null ? Convert.ToDouble("0").ToString() : oEditBesinciIsk.Value;
                        string _toplamIsk = oEditToplamIsk.Value == null ? Convert.ToDouble("0").ToString() : oEditToplamIsk.Value;

                        if (guncellemeModu == "")
                        {
                            AIFConn.StnAlmIsk.UrunleriEkle(listeGonder, _birinciIsk, _ikinciIsk, _ucuncuIsk, _dorduncuIsk, _besinciIsk, _toplamIsk, _MiktarOlcgei);
                        }
                        else if (guncellemeModu != "")
                        {
                            AIFConn.StnAlmIsk.UrunleriGuncelle(listeGonder, _birinciIsk, _ikinciIsk, _ucuncuIsk, _dorduncuIsk, _besinciIsk, _toplamIsk, _MiktarOlcgei);
                        }
                        try
                        {
                            frmSatinalmaIskontoUrunEkle.Close();
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (!pVal.BeforeAction && pVal.ItemUID == "Item_0")
                    {
                        try
                        {
                            frmSatinalmaIskontoUrunEkle.Freeze(true);
                            var xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                            var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                        select new UrunEkle()
                                        {
                                            UrunKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                        }).ToList();

                            foreach (var item in rows)
                            {
                                urunEkles.Where(x => x.UrunKodu == item.UrunKodu).ToList().ForEach(y => y.Sec = "Y");
                            }

                            string rowsFinal = string.Join("", urunEkles.Select(y => string.Format(row, y.Sec, y.UrunKodu, y.UrunTanimi, y.UrunGrupKod, y.UrunGrupAdi, y.KalemGrubu2Kodu, y.KalemGrubu2Adi, y.KalemGrubu3Kodu, y.KalemGrubu3Adi, y.KalemGrubu4Kodu, y.KalemGrubu4Adi)));

                            string data = string.Format(datatableHeader, rowsFinal);

                            oDataTable.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_All, data);

                            oMatrix.Columns.Item("Col_4").DataBind.Bind("DATA", "Sec");
                            oMatrix.Columns.Item("Col_0").DataBind.Bind("DATA", "UrunKodu");
                            oMatrix.Columns.Item("Col_1").DataBind.Bind("DATA", "UrunTanimi");
                            oMatrix.Columns.Item("Col_2").DataBind.Bind("DATA", "UrunGrupKod");
                            oMatrix.Columns.Item("Col_3").DataBind.Bind("DATA", "UrunGrupAdi");
                            oMatrix.Columns.Item("Col_5").DataBind.Bind("DATA", "KalemGrubu2Kodu");
                            oMatrix.Columns.Item("Col_6").DataBind.Bind("DATA", "KalemGrubu2Adi");
                            oMatrix.Columns.Item("Col_7").DataBind.Bind("DATA", "KalemGrubu3Kodu");
                            oMatrix.Columns.Item("Col_8").DataBind.Bind("DATA", "KalemGrubu3Adi");
                            oMatrix.Columns.Item("Col_9").DataBind.Bind("DATA", "KalemGrubu4Kodu");
                            oMatrix.Columns.Item("Col_10").DataBind.Bind("DATA", "KalemGrubu4Adi");

                            oMatrix.LoadFromDataSource();
                            oMatrix.AutoResizeColumns();

                            oBtnAra.Item.Click();
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmSatinalmaIskontoUrunEkle.Freeze(false);
                        }
                    }

                    else if (!pVal.BeforeAction && pVal.ItemUID == "Item_1")
                    {
                        try
                        {
                            frmSatinalmaIskontoUrunEkle.Freeze(true);
                            var xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                            var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                        select new UrunEkle()
                                        {
                                            UrunKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                        }).ToList();

                            foreach (var item in rows)
                            {
                                urunEkles.Where(x => x.UrunKodu == item.UrunKodu).ToList().ForEach(y => y.Sec = "N");
                            }


                            string rowsFinal = string.Join("", urunEkles.Select(y => string.Format(row, y.Sec, y.UrunKodu, y.UrunTanimi, y.UrunGrupKod, y.UrunGrupAdi, y.KalemGrubu2Kodu, y.KalemGrubu2Adi, y.KalemGrubu3Kodu, y.KalemGrubu3Adi, y.KalemGrubu4Kodu, y.KalemGrubu4Adi)));

                            string data = string.Format(datatableHeader, rowsFinal);

                            oDataTable.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_All, data);

                            oMatrix.Columns.Item("Col_4").DataBind.Bind("DATA", "Sec");
                            oMatrix.Columns.Item("Col_0").DataBind.Bind("DATA", "UrunKodu");
                            oMatrix.Columns.Item("Col_1").DataBind.Bind("DATA", "UrunTanimi");
                            oMatrix.Columns.Item("Col_2").DataBind.Bind("DATA", "UrunGrupKod");
                            oMatrix.Columns.Item("Col_3").DataBind.Bind("DATA", "UrunGrupAdi");
                            oMatrix.Columns.Item("Col_5").DataBind.Bind("DATA", "KalemGrubu2Kodu");
                            oMatrix.Columns.Item("Col_6").DataBind.Bind("DATA", "KalemGrubu2Adi");
                            oMatrix.Columns.Item("Col_7").DataBind.Bind("DATA", "KalemGrubu3Kodu");
                            oMatrix.Columns.Item("Col_8").DataBind.Bind("DATA", "KalemGrubu3Adi");
                            oMatrix.Columns.Item("Col_9").DataBind.Bind("DATA", "KalemGrubu4Kodu");
                            oMatrix.Columns.Item("Col_10").DataBind.Bind("DATA", "KalemGrubu4Adi");

                            oMatrix.LoadFromDataSource();
                            oMatrix.AutoResizeColumns();

                            oBtnAra.Item.Click();

                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmSatinalmaIskontoUrunEkle.Freeze(false);
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
                    if (!pVal.BeforeAction && pVal.ItemChanged && (pVal.ItemUID == "Item_12" || pVal.ItemUID == "Item_14" || pVal.ItemUID == "Item_16" || pVal.ItemUID == "Item_18" || pVal.ItemUID == "Item_20"))
                    {
                        try
                        {
                            double _birinciIsk = oEditBirinciIsk.Value == "" ? 0 : Convert.ToDouble(oEditBirinciIsk.Value, System.Globalization.CultureInfo.InvariantCulture);
                            double _ikinciIsk = oEditIkinciIsk.Value == "" ? 0 : Convert.ToDouble(oEditIkinciIsk.Value, System.Globalization.CultureInfo.InvariantCulture);
                            double _ucuncuIsk = oEditUcuncuIsk.Value == "" ? 0 : Convert.ToDouble(oEditUcuncuIsk.Value, System.Globalization.CultureInfo.InvariantCulture);
                            double _dorduncuIsk = oEditDorduncuIsk.Value == "" ? 0 : Convert.ToDouble(oEditDorduncuIsk.Value, System.Globalization.CultureInfo.InvariantCulture);
                            double _besinciIsk = oEditBesinciIsk.Value == "" ? 0 : Convert.ToDouble(oEditBesinciIsk.Value, System.Globalization.CultureInfo.InvariantCulture);

                            double _toplamIsk = 100 - (100 * (1 - (_birinciIsk / 100)) * (1 - (_ikinciIsk / 100)) * (1 - (_ucuncuIsk / 100)) * (1 - (_dorduncuIsk / 100)) * (1 - (_besinciIsk / 100)));

                            oEditToplamIsk.Value = _toplamIsk.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        }
                        catch (Exception)
                        {
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
                    //int row = oMatrix.GetNextSelectedRow();
                    //if (row != -1)
                    //{
                    //    //var val1 = ((SAPbouiCOM.ComboBox)oMatrixDetay.Columns.Item("Col_0").Cells.Item(row).Specific).Value.ToString().Trim();
                    //    //DahaOnceSecilenKisiler.Remove(Convert.ToInt32(val1));

                    //    oMatrix.DeleteRow(row);
                    //    if (frmIndirimGiris.Mode == BoFormMode.fm_OK_MODE)
                    //    {
                    //        frmIndirimGiris.Mode = BoFormMode.fm_UPDATE_MODE;
                    //    }


                    //}
                }
                else if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && pVal.BeforeAction)
                {
                    //frmIndirimGiris.DataSources.DBDataSources.Item(1).Clear();
                    //oMatrix.AddRow();

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
            return;
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

    }
}
