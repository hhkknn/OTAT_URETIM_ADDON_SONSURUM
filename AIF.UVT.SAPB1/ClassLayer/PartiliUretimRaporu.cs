using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using AIF.UVT.SAPB1.Models;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class PartiliUretimRaporu
    {
        [ItemAtt(AIFConn.PartiliUretimRaporuUID)]
        public SAPbouiCOM.Form frmPartiliUretimRaporu;

        [ItemAtt("Item_10")]
        public SAPbouiCOM.Matrix oMatrix;

        [ItemAtt("Item_5")]
        public SAPbouiCOM.Button oBtnListele;

        [ItemAtt("Item_1")]
        public SAPbouiCOM.EditText oEditKalemKodu;

        [ItemAtt("Item_3")]
        public SAPbouiCOM.EditText oEditKalemAdi;

        [ItemAtt("Item_4")]
        public SAPbouiCOM.EditText oEditPartiNo;

        [ItemAtt("Item_11")]
        public SAPbouiCOM.Button oBtnSecim;


        private SAPbouiCOM.DataTable oDataTable = null;
        private List<PartiliUretimDetay> partiliUretimDetayOrj = new List<PartiliUretimDetay>();
        private List<PartiliUretimDetay> partiliUretimDetays = new List<PartiliUretimDetay>();

        string kalemKodu = "";
        string kalemAdi = "";
        string partiNo = "";

        public void LoadForms()
        {

            ConstVariables.oFnc.LoadSAPXML(AIFConn.PartiliUretimRaporuXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.PartiliUretimRaporuXML));
            Functions.CreateUserOrSystemFormComponent<PartiliUretimRaporu>(AIFConn.PartiUretRap);

            InitForms();
        }

        string header = @"<?xml version=""1.0"" encoding=""UTF-16"" ?><DataTable Uid=""DT_0""><Columns><Column Uid=""BelgeNo"" Type=""2"" MaxLength=""0""/><Column Uid=""KayitTarihi"" Type=""4"" MaxLength=""0""/><Column Uid=""UstKalemParti"" Type=""1"" MaxLength=""36""/><Column Uid=""UstKalem"" Type=""1"" MaxLength=""50""/><Column Uid=""UrunTanimi"" Type=""1"" MaxLength=""100""/><Column Uid=""GirisDeposu"" Type=""1"" MaxLength=""8""/><Column Uid=""Giren"" Type=""7"" MaxLength=""0""/><Column Uid=""GirenOB"" Type=""1"" MaxLength=""20""/><Column Uid=""AltKalem"" Type=""1"" MaxLength=""50""/><Column Uid=""AltKalemAdi"" Type=""1"" MaxLength=""100""/><Column Uid=""AltKalemParti"" Type=""1"" MaxLength=""36""/><Column Uid=""CikanOB"" Type=""1"" MaxLength=""20""/><Column Uid=""CikisDeposu"" Type=""1"" MaxLength=""8""/><Column Uid=""Cikan"" Type=""7"" MaxLength=""0""/><Column Uid=""Guid1"" Type=""1"" MaxLength=""50""/><Column Uid=""Guid2"" Type=""1"" MaxLength=""50""/><Column Uid=""Guid3"" Type=""1"" MaxLength=""50""/><Column Uid=""Tip"" Type=""1"" MaxLength=""1""/></Columns><Rows>{0}</Rows></DataTable>";

        string row = "<Row><Cells><Cell><ColumnUid>BelgeNo</ColumnUid><Value>{0}</Value></Cell><Cell><ColumnUid>KayitTarihi</ColumnUid><Value>{1}</Value></Cell><Cell><ColumnUid>UstKalemParti</ColumnUid><Value>{2}</Value></Cell><Cell><ColumnUid>UstKalem</ColumnUid><Value>{3}</Value></Cell><Cell><ColumnUid>UrunTanimi</ColumnUid><Value>{4}</Value></Cell><Cell><ColumnUid>GirisDeposu</ColumnUid><Value>{5}</Value></Cell><Cell><ColumnUid>Giren</ColumnUid><Value>{6}</Value></Cell><Cell><ColumnUid>GirenOB</ColumnUid><Value>{7}</Value></Cell><Cell><ColumnUid>AltKalem</ColumnUid><Value>{8}</Value></Cell><Cell><ColumnUid>AltKalemAdi</ColumnUid><Value>{9}</Value></Cell><Cell><ColumnUid>AltKalemParti</ColumnUid><Value>{10}</Value></Cell><Cell><ColumnUid>CikanOB</ColumnUid><Value>{11}</Value></Cell><Cell><ColumnUid>CikisDeposu</ColumnUid><Value>{12}</Value></Cell><Cell><ColumnUid>Cikan</ColumnUid><Value>{13}</Value></Cell><Cell><ColumnUid>Guid1</ColumnUid><Value>{14}</Value></Cell><Cell><ColumnUid>Guid2</ColumnUid><Value>{15}</Value></Cell><Cell><ColumnUid>Guid3</ColumnUid><Value>{16}</Value></Cell><Cell><ColumnUid>Tip</ColumnUid><Value>{17}</Value></Cell></Cells></Row>";
        public void InitForms()
        {
            try
            {
                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                frmPartiliUretimRaporu.EnableMenu("1283", false);
                frmPartiliUretimRaporu.EnableMenu("1284", false);
                frmPartiliUretimRaporu.EnableMenu("1286", false);

                oDataTable = frmPartiliUretimRaporu.DataSources.DataTables.Add("DATA");

                oMatrix.AutoResizeColumns();


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
                    if (BusinessObjectInfo.BeforeAction)
                    {
                        //kaydet();
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

        private string val = "";

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
                    if (pVal.ItemUID == "Item_5" && !pVal.BeforeAction)
                    {

                        try
                        {
                            string sql = "";

                            sql = "SELECT T0.\"BaseEntry\" as \"BelgeNo\",cast(T1.DocDate as date) as \"KayitTarihi\", T1.\"BatchNum\" as \"UstKalemParti\", T0.\"ItemCode\" as \"UstKalem\",T0.\"Dscription\" as \"UrunTanimi\",T0.\"WhsCode\" \"GirisDeposu\",t0.\"Quantity\" as \"Giren\",T0.\"UomCode\" as \"GirenOB\",T3.\"ItemCode\" as \"AltKalem\",T3.\"Dscription\" as \"AltKalemAdi\",T4.\"BatchNum\" as \"AltKalemParti\",t3.\"UomCode\" as \"CikanOB\",t3.\"WhsCode\" as \"CikisDeposu\",T3.\"Quantity\" as \"Cikan\",NEWID() as \"Guid1\" FROM \"IGN1\" T0 ";
                            sql += "LEFT JOIN \"IBT1\" T1 ON t1.\"BaseEntry\" = T0.\"DocEntry\" AND t1.\"ItemCode\" = T0.\"ItemCode\" AND t1.\"BaseType\" = T0.\"ObjType\" LEFT JOIN \"OIGE\" t2 ON t2.\"U_BatchNumber\" = t1.\"BatchNum\" LEFT JOIN \"IGE1\" t3 ON t3.\"DocEntry\" = T2.\"DocEntry\" LEFT JOIN \"IBT1\" T4 ON T4.\"BaseEntry\" = T3.\"DocEntry\" AND T4.\"BsDocLine\" = T3.\"LineNum\" AND T4.\"ItemCode\" = T3.\"ItemCode\" ";
                            //sql += " WHERE t1.\"BatchNum\" = '" + oEditPartiNo.Value.ToString() + "' and T1.\"ItemCode\" = '" + oEditKalemKodu.Value.ToString() + "' ";
                            sql += " WHERE t1.\"BatchNum\" = '20210707-71-1' and T1.\"ItemCode\" = 'YRM-30000' ";

                            oDataTable.Clear();
                            oDataTable.ExecuteQuery(sql);
                            oMatrix.Clear();

                            oMatrix.Columns.Item("Col_0").DataBind.Bind("DATA", "BelgeNo");
                            oMatrix.Columns.Item("Col_1").DataBind.Bind("DATA", "KayitTarihi");
                            oMatrix.Columns.Item("Col_2").DataBind.Bind("DATA", "UstKalemParti");
                            oMatrix.Columns.Item("Col_3").DataBind.Bind("DATA", "UstKalem");
                            oMatrix.Columns.Item("Col_4").DataBind.Bind("DATA", "UrunTanimi");
                            oMatrix.Columns.Item("Col_5").DataBind.Bind("DATA", "GirisDeposu");
                            oMatrix.Columns.Item("Col_6").DataBind.Bind("DATA", "Giren");
                            oMatrix.Columns.Item("Col_7").DataBind.Bind("DATA", "GirenOB");
                            oMatrix.Columns.Item("Col_8").DataBind.Bind("DATA", "AltKalem");
                            oMatrix.Columns.Item("Col_9").DataBind.Bind("DATA", "AltKalemAdi");
                            oMatrix.Columns.Item("Col_10").DataBind.Bind("DATA", "AltKalemParti");
                            oMatrix.Columns.Item("Col_11").DataBind.Bind("DATA", "CikanOB");
                            oMatrix.Columns.Item("Col_12").DataBind.Bind("DATA", "CikisDeposu");
                            oMatrix.Columns.Item("Col_13").DataBind.Bind("DATA", "Cikan");
                            oMatrix.Columns.Item("Col_14").DataBind.Bind("DATA", "Guid1");

                            oMatrix.LoadFromDataSource();
                            oMatrix.AutoResizeColumns();


                            string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                            partiliUretimDetayOrj = (from x in XDocument.Parse(xml).Descendants("Row")
                                                     select new PartiliUretimDetay()
                                                     {
                                                         //sec = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "#" select new XElement(y.Element("Value"))).First().Value,
                                                         BelgeNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                                         KayitTarihi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                                                         UstKalemParti = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value,
                                                         UstKalem = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value,
                                                         UrunTanimi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,
                                                         GirisDeposu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                                                         Giren = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value,
                                                         GirenOB = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value,
                                                         AltKalem = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value,
                                                         AltKalemAdi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_9" select new XElement(y.Element("Value"))).First().Value,
                                                         AltKalemParti = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_10" select new XElement(y.Element("Value"))).First().Value,
                                                         CikanOB = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_11" select new XElement(y.Element("Value"))).First().Value,
                                                         CikisDeposu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_12" select new XElement(y.Element("Value"))).First().Value,
                                                         Cikan = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_13" select new XElement(y.Element("Value"))).First().Value,
                                                         SatirNo = x.ElementsBeforeSelf().Count() + 1,
                                                         guidId = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_14" select new XElement(y.Element("Value"))).First().Value,
                                                         guidId2 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_15" select new XElement(y.Element("Value"))).First().Value,
                                                         guidId3 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_16" select new XElement(y.Element("Value"))).First().Value,
                                                         tip = "A"
                                                     }).ToList();

                        }
                        catch (Exception ex)
                        {
                            Handler.SAPApplication.MessageBox("Hata oluştu." + ex.Message);
                            return false;
                        }
                    }
                    else if (pVal.ItemUID == "Item_8" && !pVal.BeforeAction)
                    {
                        try
                        {
                            this.frmPartiliUretimRaporu.Close();
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_11" && !pVal.BeforeAction)
                    {
                        AIFConn.UretimSecm.LoadForms();
                    }
                    break;

                case BoEventTypes.et_DOUBLE_CLICK:
                    if (pVal.ColUID == "#" && !pVal.BeforeAction)
                    {
                        try
                        {
                            frmPartiliUretimRaporu.Freeze(true);

                            string altKalemParti = "";
                            string guidid = "";

                            altKalemParti = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_10").Cells.Item(pVal.Row).Specific).Value;
                            guidid = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_14").Cells.Item(pVal.Row).Specific).Value;

                            if (altKalemParti != "")
                            {

                                string sql = "";

                                sql = "SELECT T0.\"BaseEntry\" as \"BelgeNo\",cast(T1.DocDate as date) as \"KayitTarihi\", T1.\"BatchNum\" as \"UstKalemParti\", T0.\"ItemCode\" as \"UstKalem\",T0.\"Dscription\" as \"UrunTanimi\",T0.\"WhsCode\" \"GirisDeposu\",t0.\"Quantity\" as \"Giren\",T0.\"UomCode\" as \"GirenOB\",T3.\"ItemCode\" as \"AltKalem\",T3.\"Dscription\" as \"AltKalemAdi\",T4.\"BatchNum\" as \"AltKalemParti\",t3.\"UomCode\" as \"CikanOB\",t3.\"WhsCode\" as \"CikisDeposu\",T3.\"Quantity\" as \"Cikan\" FROM \"IGN1\" T0 ";
                                sql += "LEFT JOIN \"IBT1\" T1 ON t1.\"BaseEntry\" = T0.\"DocEntry\" AND t1.\"ItemCode\" = T0.\"ItemCode\" AND t1.\"BaseType\" = T0.\"ObjType\" LEFT JOIN \"OIGE\" t2 ON t2.\"U_BatchNumber\" = t1.\"BatchNum\" LEFT JOIN \"IGE1\" t3 ON t3.\"DocEntry\" = T2.\"DocEntry\" LEFT JOIN \"IBT1\" T4 ON T4.\"BaseEntry\" = T3.\"DocEntry\" AND T4.\"BsDocLine\" = T3.\"LineNum\" AND T4.\"ItemCode\" = T3.\"ItemCode\" ";
                                sql += " WHERE T4.\"BatchNum\" = '" + altKalemParti + "'";

                                //ConstVariables.oRecordset.DoQuery(sql);

                                //if (ConstVariables.oRecordset.RecordCount > 0)
                                //{

                                //}

                                //oDataTable.Clear();
                                oDataTable.ExecuteQuery(sql);
                                var xml = oDataTable.SerializeAsXML(BoDataTableXmlSelect.dxs_All);
                                var rows2 = getData(xml);

                                //rows2.ToList().ForEach(y=>)



                                partiliUretimDetayOrj.AddRange(rows2);



                                var data2 = setDataXML(rows2);

                                var final = string.Format(header, data2);

                                oDataTable.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_All, final);

                                oMatrix.LoadFromDataSource();
                                oMatrix.AutoResizeColumns();

                            }

                        }
                        catch (Exception ex)
                        {
                            Handler.SAPApplication.MessageBox("Hata oluştu." + ex.Message);
                        }
                        finally
                        {
                            frmPartiliUretimRaporu.Freeze(false);
                        }
                    }
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
                    //if (pVal.ItemUID == "Item_1" && pVal.BeforeAction)
                    //{
                    //SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                    //oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                    //SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                    //oCFL = frmPartiliUretimRaporu.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                    //SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                    //SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                    //SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                    //oCFL.SetConditions(oEmptyConts);
                    //oCons = oCFL.GetConditions();

                    //oCon = oCons.Add();
                    //oCon.Alias = "validFor";
                    //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    //oCon.CondVal = "Y";
                    //oCFL.SetConditions(oCons);
                    //}
                    //else 
                    if (pVal.ItemUID == "Item_1" && !pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;

                            if (oDataTable == null) //datatable boş oldugunda işleme devam etmesin
                            {
                                break;
                            }
                            string Val = "";
                            Val = oDataTable.GetValue("ItemCode", 0).ToString();

                            try
                            {
                                oEditKalemKodu.Value = Val;
                            }
                            catch (Exception)
                            {
                            }

                            Val = oDataTable.GetValue("ItemName", 0).ToString();
                            try
                            {
                                oEditKalemAdi.Value = Val;
                            }
                            catch (Exception)
                            {
                            }

                            Val = oDataTable.GetValue("BatchNum", 0).ToString();
                            try
                            {
                                oEditPartiNo.Value = Val;
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
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        public List<PartiliUretimDetay> getData(string xml)
        {
            var rows2 = (from x in XDocument.Parse(xml).Descendants("Row")
                         select new PartiliUretimDetay()
                         {
                             BelgeNo = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "BelgeNo" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "BelgeNo" select new XElement(y.Element("Value"))).First().Value : "",
                             KayitTarihi = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "KayitTarihi" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "KayitTarihi" select new XElement(y.Element("Value"))).First().Value : "",
                             UstKalemParti = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "UstKalemParti" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "UstKalemParti" select new XElement(y.Element("Value"))).First().Value : "",
                             UstKalem = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "UstKalem" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "UstKalem" select new XElement(y.Element("Value"))).First().Value : "",
                             UrunTanimi = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "UrunTanimi" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "UrunTanimi" select new XElement(y.Element("Value"))).First().Value : "",
                             GirisDeposu = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "GirisDeposu" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "GirisDeposu" select new XElement(y.Element("Value"))).First().Value : "",
                             Giren = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "Giren" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "Giren" select new XElement(y.Element("Value"))).First().Value : "",
                             GirenOB = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "GirenOB" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "GirenOB" select new XElement(y.Element("Value"))).First().Value : "",
                             AltKalem = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "AltKalem" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "AltKalem" select new XElement(y.Element("Value"))).First().Value : "",
                             AltKalemAdi = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "AltKalemAdi" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "AltKalemAdi" select new XElement(y.Element("Value"))).First().Value : "",
                             AltKalemParti = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "AltKalemParti" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "AltKalemParti" select new XElement(y.Element("Value"))).First().Value : "",
                             CikanOB = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "CikanOB" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "CikanOB" select new XElement(y.Element("Value"))).First().Value : "",
                             CikisDeposu = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "CikisDeposu" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "CikisDeposu" select new XElement(y.Element("Value"))).First().Value : "",
                             Cikan = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "Cikan" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "Cikan" select new XElement(y.Element("Value"))).First().Value : "",
                             SatirNo = x.ElementsBeforeSelf().Count() + 1,

                         }).ToList();

            return rows2;
        }
        public string setDataXML(List<PartiliUretimDetay> rows2)
        {
            string data2 = "";
            data2 = string.Join("", rows2.Select(s => string.Format(row, s.BelgeNo, s.KayitTarihi, s.UstKalemParti, s.UstKalem, s.UrunTanimi, s.GirisDeposu, s.Giren, s.GirenOB, s.AltKalem, s.AltKalemAdi, s.AltKalemParti, s.CikanOB, s.CikisDeposu, s.Cikan, s.SatirNo + 1)));

            return data2;

        }

    }
}