using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class HaftalikPlan
    {
        [ItemAtt(AIFConn.HaftalikPlanUID)]
        public SAPbouiCOM.Form frmHaftalikPlan;

        //[ItemAtt("Item_2")]
        //public SAPbouiCOM.EditText EdtHaftaSayisi;
        [ItemAtt("Item_8")]
        public SAPbouiCOM.ComboBox cmbHafta;
        //[ItemAtt("1")]
        //public SAPbouiCOM.Button btnAddOrUpdate;
        //[ItemAtt("Item_4")]
        //public SAPbouiCOM.Grid oGrid;
        //SAPbouiCOM.DataTable oGridDataTable = null;


        [ItemAtt("Item_9")]
        public SAPbouiCOM.EditText EdtBaslangicTarihi;

        [ItemAtt("Item_12")]
        public SAPbouiCOM.EditText EdtBitisTarihi;

        [ItemAtt("Item_1")]
        public SAPbouiCOM.Matrix oMatrix;
        SAPbouiCOM.DataTable oDataTable;
        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.HaftalikPlanXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.HaftalikPlanXML));
            Functions.CreateUserOrSystemFormComponent<HaftalikPlan>(AIFConn.HaftaPlan);

            InitForms();
        }

        string header = @"<?xml version=""1.0"" encoding=""UTF-16"" ?><DataTable Uid=""DT_0""><Columns><Column Uid=""Grup"" Type=""1"" MaxLength=""50""/><Column Uid=""GrupSira"" Type=""1"" MaxLength=""50""/><Column Uid=""Tur"" Type=""1"" MaxLength=""50""/><Column Uid=""UrunKodu"" Type=""1"" MaxLength=""100""/><Column Uid=""UrunAdi"" Type=""1"" MaxLength=""100""/><Column Uid=""Seviye1"" Type=""2"" MaxLength=""0""/><Column Uid=""Seviye2"" Type=""2"" MaxLength=""0""/><Column Uid=""Seviye3"" Type=""2"" MaxLength=""0""/><Column Uid=""a1"" Type=""5"" MaxLength=""0""/><Column Uid=""b1"" Type=""5"" MaxLength=""0""/><Column Uid=""c1"" Type=""5"" MaxLength=""0""/><Column Uid=""a2"" Type=""5"" MaxLength=""0""/><Column Uid=""b2"" Type=""5"" MaxLength=""0""/><Column Uid=""c2"" Type=""5"" MaxLength=""0""/><Column Uid=""a3"" Type=""5"" MaxLength=""0""/><Column Uid=""b3"" Type=""5"" MaxLength=""0""/><Column Uid=""c3"" Type=""5"" MaxLength=""0""/><Column Uid=""a4"" Type=""5"" MaxLength=""0""/><Column Uid=""b4"" Type=""5"" MaxLength=""0""/><Column Uid=""c4"" Type=""5"" MaxLength=""0""/><Column Uid=""a5"" Type=""5"" MaxLength=""0""/><Column Uid=""b5"" Type=""5"" MaxLength=""0""/><Column Uid=""c5"" Type=""5"" MaxLength=""0""/><Column Uid=""a6"" Type=""5"" MaxLength=""0""/><Column Uid=""b6"" Type=""5"" MaxLength=""0""/><Column Uid=""c6"" Type=""5"" MaxLength=""0""/><Column Uid=""a7"" Type=""5"" MaxLength=""0""/><Column Uid=""b7"" Type=""5"" MaxLength=""0""/><Column Uid=""c7"" Type=""5"" MaxLength=""0""/></Columns><Rows>{0}</Rows></DataTable>";

        string row = "<Row><Cells><Cell><ColumnUid>Grup</ColumnUid><Value>{0}</Value></Cell><Cell><ColumnUid>GrupSira</ColumnUid><Value>{1}</Value></Cell><Cell><ColumnUid>Tur</ColumnUid><Value>{2}</Value></Cell><Cell><ColumnUid>UrunKodu</ColumnUid><Value>{3}</Value></Cell><Cell><ColumnUid>UrunAdi</ColumnUid><Value>{4}</Value></Cell><Cell><ColumnUid>Seviye1</ColumnUid><Value>{5}</Value></Cell><Cell><ColumnUid>Seviye2</ColumnUid><Value>{6}</Value></Cell><Cell><ColumnUid>Seviye3</ColumnUid><Value>{7}</Value></Cell><Cell><ColumnUid>a1</ColumnUid><Value>{8}</Value></Cell><Cell><ColumnUid>b1</ColumnUid><Value>{9}</Value></Cell><Cell><ColumnUid>c1</ColumnUid><Value>{10}</Value></Cell><Cell><ColumnUid>a2</ColumnUid><Value>{11}</Value></Cell><Cell><ColumnUid>b2</ColumnUid><Value>{12}</Value></Cell><Cell><ColumnUid>c2</ColumnUid><Value>{13}</Value></Cell><Cell><ColumnUid>a3</ColumnUid><Value>{14}</Value></Cell><Cell><ColumnUid>b3</ColumnUid><Value>{15}</Value></Cell><Cell><ColumnUid>c3</ColumnUid><Value>{16}</Value></Cell><Cell><ColumnUid>a4</ColumnUid><Value>{17}</Value></Cell><Cell><ColumnUid>b4</ColumnUid><Value>{18}</Value></Cell><Cell><ColumnUid>c4</ColumnUid><Value>{19}</Value></Cell><Cell><ColumnUid>a5</ColumnUid><Value>{20}</Value></Cell><Cell><ColumnUid>b5</ColumnUid><Value>{21}</Value></Cell><Cell><ColumnUid>c5</ColumnUid><Value>{22}</Value></Cell><Cell><ColumnUid>a6</ColumnUid><Value>{23}</Value></Cell><Cell><ColumnUid>b6</ColumnUid><Value>{24}</Value></Cell><Cell><ColumnUid>c6</ColumnUid><Value>{25}</Value></Cell><Cell><ColumnUid>a7</ColumnUid><Value>{26}</Value></Cell><Cell><ColumnUid>b7</ColumnUid><Value>{27}</Value></Cell><Cell><ColumnUid>c7</ColumnUid><Value>{28}</Value></Cell></Cells></Row>";

        List<Tuple<string, string, string>> dynamicColumns = new List<Tuple<string, string, string>>();
        public void InitForms()
        {
            try
            {
                oDataTable = (SAPbouiCOM.DataTable)frmHaftalikPlan.DataSources.DataTables.Item("DT_0");
                dynamicColumns.Add(Tuple.Create("Col_8", "N", "1"));
                dynamicColumns.Add(Tuple.Create("Col_9", "N", "1"));
                dynamicColumns.Add(Tuple.Create("Col_10", "N", "1"));
                dynamicColumns.Add(Tuple.Create("Col_11", "N", "2"));
                dynamicColumns.Add(Tuple.Create("Col_12", "N", "2"));
                dynamicColumns.Add(Tuple.Create("Col_13", "N", "2"));
                dynamicColumns.Add(Tuple.Create("Col_14", "N", "3"));
                dynamicColumns.Add(Tuple.Create("Col_15", "N", "3"));
                dynamicColumns.Add(Tuple.Create("Col_16", "N", "3"));
                dynamicColumns.Add(Tuple.Create("Col_17", "N", "4"));
                dynamicColumns.Add(Tuple.Create("Col_18", "N", "4"));
                dynamicColumns.Add(Tuple.Create("Col_19", "N", "4"));
                dynamicColumns.Add(Tuple.Create("Col_20", "N", "5"));
                dynamicColumns.Add(Tuple.Create("Col_21", "N", "5"));
                dynamicColumns.Add(Tuple.Create("Col_22", "N", "5"));
                dynamicColumns.Add(Tuple.Create("Col_23", "N", "6"));
                dynamicColumns.Add(Tuple.Create("Col_24", "N", "6"));
                dynamicColumns.Add(Tuple.Create("Col_25", "N", "6"));
                dynamicColumns.Add(Tuple.Create("Col_26", "N", "7"));
                dynamicColumns.Add(Tuple.Create("Col_27", "N", "7"));
                dynamicColumns.Add(Tuple.Create("Col_28", "N", "7"));




                List<Helper.ValidValue> list;
                string sql = "Declare @i int set @i = 1; CREATE TABLE #Haftalar (HaftaSayisi int) WHILE(@i <= 53) BEGIN insert into #Haftalar(HaftaSayisi) Values (@i) set @i = @i + 1 END select HaftaSayisi as value, Convert(varchar(30),HaftaSayisi) + '.Hafta' as description from #Haftalar drop table #Haftalar";
                var aa = ConstVariables.oCompanyObject;
                list = Helper.GetValidValuesFromRS(sql);

                Helper nesne = new Helper();
                nesne.ComboAction(frmHaftalikPlan, "Item_8", Helper.ActionCombo.add, list);

                cmbHafta.ExpandType = BoExpandType.et_DescriptionOnly;

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
        List<Tuple<DateTime, DateTime>> datesData = new List<Tuple<DateTime, DateTime>>();
        public List<Tuple<DateTime, DateTime>> HaftaIslemleri(int GelenYil, int GelenHaftaNumarasi)
        {
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            DateTime YilinIlkGunu = new DateTime(GelenYil, 1, 1);
            int GunDengesi = DayOfWeek.Monday - YilinIlkGunu.DayOfWeek;
            DateTime IlkGun = YilinIlkGunu.AddDays(GunDengesi);
            Calendar Takvim = dfi.Calendar;
            int IlkHafta = Takvim.GetWeekOfYear(IlkGun, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            var HaftaNumarasi = GelenHaftaNumarasi;
            if (IlkHafta <= 1)
            {
                HaftaNumarasi -= 1;
            }
            var Sonuc = IlkGun.AddDays(HaftaNumarasi * 7);
            var sn = Sonuc.AddDays(-7);

            datesData.Add(Tuple.Create(sn, Sonuc));

            return datesData;
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
                    if (pVal.ItemUID == "Item_8" && !pVal.BeforeAction)
                    {
                        datesData = new List<Tuple<DateTime, DateTime>>();
                        datesData = HaftaIslemleri(DateTime.Now.Year, Convert.ToInt32(cmbHafta.Value.Trim()));

                        DateTime dt = datesData[0].Item1;

                        EdtBaslangicTarihi.Value = dt.ToString("yyyyMMdd");

                        dt = dt.AddDays(7);

                        EdtBitisTarihi.Value = dt.ToString("yyyyMMdd");

                    }
                    break;
                case BoEventTypes.et_CLICK:
                    if (pVal.ItemUID == "Item_5" && !pVal.BeforeAction)
                    {
                        try
                        {
                            if (cmbHafta.Value.Trim() == "")
                            {
                                Handler.SAPApplication.MessageBox("Hafta seçimi yapılmalıdır.");
                                BubbleEvent = false;
                                return false;
                            }
                            frmHaftalikPlan.Freeze(true);
                            datesData = new List<Tuple<DateTime, DateTime>>();
                            OriginalData = new List<Datas>();
                            datesData = HaftaIslemleri(DateTime.Now.Year, Convert.ToInt32(cmbHafta.Value.Trim()));
                            string date1 = datesData[0].Item1.ToString("yyyyMMdd");
                            string date2 = datesData[0].Item2.ToString("yyyyMMdd");

                            string qry = "EXEC HaftalikPlan '" + date1 + "', '" + date2 + "'";

                            oDataTable.ExecuteQuery(qry);

                            #region MyRegion
                            //oGridDataTable.Clear();
                            //oGridDataTable.ExecuteQuery(qry);
                            //DateTime dt = DateTime.Now;
                            //string colname = "";
                            //string tip = "";

                            //int color = ColorTranslator.ToOle(System.Drawing.Color.Gray);
                            //int color2 = ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
                            ////int forecolor = ColorTranslator.ToOle(System.Drawing.Color.White);

                            ////foreach (var item in vals)
                            ////{
                            ////    oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                            ////    oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), forecolor);
                            ////}

                            //int ix = 1;
                            //System.Data.DataTable dataTable = new System.Data.DataTable("table");
                            //System.Data.DataColumn dtColumn = new System.Data.DataColumn();
                            //for (int i = 8; i <= oGridDataTable.Columns.Count - 1; i++)
                            //{
                            //    colname = oGridDataTable.Columns.Item(i).Name.Substring(0, 8);
                            //    tip = oGridDataTable.Columns.Item(i).Name.Substring(8, oGridDataTable.Columns.Item(i).Name.Length - 8);
                            //    dt = new DateTime(Convert.ToInt32(colname.Substring(0, 4)), Convert.ToInt32(colname.Substring(4, 2)), Convert.ToInt32(colname.Substring(6, 2)));

                            //    oGrid.Columns.Item(i).TitleObject.Caption = dt.ToString("dd/MM/yyyy") + " " + CultureInfo.GetCultureInfo("tr-TR").DateTimeFormat.DayNames[(int)dt.DayOfWeek] + " " + tip;
                            //    //oGrid.CommonSetting.SetCellFontStyle(1, i, BoFontStyle.fs_Strikeout);
                            //    //oGrid.Columns.Item(i).;
                            //    if (ix == 7)
                            //    {
                            //        ix = 1;
                            //    }
                            //    if (ix <= 3)
                            //    {
                            //        oGrid.Columns.Item(i).BackColor = color;
                            //    }
                            //    else
                            //    {
                            //        oGrid.Columns.Item(i).BackColor = color2;
                            //    }
                            //    ix++;
                            //}

                            //ix = 1; 

                            //oGrid.AutoResizeColumns();
                            //Handler.SAPApplication.StatusBar.SetText("Listeleme tamamlandı.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                            //var xml = oGrid.DataTable.SerializeAsXML(BoDataTableXmlSelect.dxs_All); 
                            #endregion


                            oMatrix.Columns.Item("Col_0").DataBind.Bind("DT_0", "Grup");
                            oMatrix.Columns.Item("Col_1").DataBind.Bind("DT_0", "GrupSira");
                            oMatrix.Columns.Item("Col_2").DataBind.Bind("DT_0", "Tur");
                            oMatrix.Columns.Item("Col_3").DataBind.Bind("DT_0", "UrunKodu");
                            oMatrix.Columns.Item("Col_4").DataBind.Bind("DT_0", "UrunAdi");
                            oMatrix.Columns.Item("Col_5").DataBind.Bind("DT_0", "Seviye1");
                            oMatrix.Columns.Item("Col_6").DataBind.Bind("DT_0", "Seviye2");
                            oMatrix.Columns.Item("Col_7").DataBind.Bind("DT_0", "Seviye3");
                            oMatrix.Columns.Item("Col_8").DataBind.Bind("DT_0", "a1");
                            oMatrix.Columns.Item("Col_9").DataBind.Bind("DT_0", "b1");
                            oMatrix.Columns.Item("Col_10").DataBind.Bind("DT_0", "c1");
                            oMatrix.Columns.Item("Col_11").DataBind.Bind("DT_0", "a2");
                            oMatrix.Columns.Item("Col_12").DataBind.Bind("DT_0", "b2");
                            oMatrix.Columns.Item("Col_13").DataBind.Bind("DT_0", "c2");
                            oMatrix.Columns.Item("Col_14").DataBind.Bind("DT_0", "a3");
                            oMatrix.Columns.Item("Col_15").DataBind.Bind("DT_0", "b3");
                            oMatrix.Columns.Item("Col_16").DataBind.Bind("DT_0", "c3");
                            oMatrix.Columns.Item("Col_17").DataBind.Bind("DT_0", "a4");
                            oMatrix.Columns.Item("Col_18").DataBind.Bind("DT_0", "b4");
                            oMatrix.Columns.Item("Col_19").DataBind.Bind("DT_0", "c4");
                            oMatrix.Columns.Item("Col_20").DataBind.Bind("DT_0", "a5");
                            oMatrix.Columns.Item("Col_21").DataBind.Bind("DT_0", "b5");
                            oMatrix.Columns.Item("Col_22").DataBind.Bind("DT_0", "c5");
                            oMatrix.Columns.Item("Col_23").DataBind.Bind("DT_0", "a6");
                            oMatrix.Columns.Item("Col_24").DataBind.Bind("DT_0", "b6");
                            oMatrix.Columns.Item("Col_25").DataBind.Bind("DT_0", "c6");
                            oMatrix.Columns.Item("Col_26").DataBind.Bind("DT_0", "a7");
                            oMatrix.Columns.Item("Col_27").DataBind.Bind("DT_0", "b7");
                            oMatrix.Columns.Item("Col_28").DataBind.Bind("DT_0", "c7");

                            oMatrix.LoadFromDataSource();
                            oMatrix.AutoResizeColumns();



                            var xml = oDataTable.SerializeAsXML(BoDataTableXmlSelect.dxs_All);


                            var res = getData(xml);


                            OriginalData.AddRange(res);

                            string renk = "Gray";
                            int color = ColorTranslator.ToOle(System.Drawing.Color.Gray);

                            DateTime dt = datesData[0].Item1;
                            int ix = 1;
                            for (int i = 1; i <= oMatrix.Columns.Count - 1; i++)
                            {
                                string coluid = oMatrix.Columns.Item(i).UniqueID;

                                if (dynamicColumns.Where(x => x.Item1 == coluid).Count() > 0)
                                {
                                    string tip = ix == 1 ? "KG" : ix == 2 ? "PALET" : "ADET";
                                    string colname = dt.ToString("dd/MM/yyyy") + " " + CultureInfo.GetCultureInfo("tr-TR").DateTimeFormat.DayNames[(int)dt.DayOfWeek] + " " + tip;
                                    oMatrix.Columns.Item(i).TitleObject.Caption = colname;
                                    oMatrix.Columns.Item(i).BackColor = color;
                                    ix++;

                                    if (ix == 4)
                                    {
                                        if (renk == "Gray")
                                        {
                                            renk = "DarkGray";
                                            color = ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
                                        }
                                        else
                                        {
                                            renk = "Gray";
                                            color = ColorTranslator.ToOle(System.Drawing.Color.Gray);
                                        }
                                        dt = dt.AddDays(1);
                                        ix = 1;
                                    }
                                }
                            }
                            oMatrix.AutoResizeColumns();


                        }
                        catch (Exception ex)
                        {
                        }
                        finally
                        {
                            frmHaftalikPlan.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "Item_7" && !pVal.BeforeAction)
                    {
                        try
                        {
                            frmHaftalikPlan.Freeze(true);
                            string data2 = string.Join("", OriginalData.Where(x => x.Seviye3 != "3").Select(s => string.Format(row, s.Grup, s.GrupSira, s.Tur, s.UrunKodu, s.UrunAdi, s.Seviye1, s.Seviye2, s.Seviye3, s.a1, s.b1, s.c1, s.a2, s.b2, s.c2, s.a3, s.b3, s.c3, s.a4, s.b4, s.c4, s.a5, s.b5, s.c5, s.a6, s.b6, s.c6, s.a7, s.b7, s.c7)));


                            var final = string.Format(header, data2);

                            oDataTable.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_All, final);

                            oMatrix.LoadFromDataSource();
                            oMatrix.AutoResizeColumns();
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmHaftalikPlan.Freeze(false);
                        }

                    }
                    else if (pVal.ItemUID == "Item_6" && !pVal.BeforeAction)
                    {
                        try
                        {
                            frmHaftalikPlan.Freeze(true);
                            string data2 = string.Join("", OriginalData.Select(s => string.Format(row, s.Grup, s.GrupSira, s.Tur, s.UrunKodu, s.UrunAdi, s.Seviye1, s.Seviye2, s.Seviye3, s.a1, s.b1, s.c1, s.a2, s.b2, s.c2, s.a3, s.b3, s.c3, s.a4, s.b4, s.c4, s.a5, s.b5, s.c5, s.a6, s.b6, s.c6, s.a7, s.b7, s.c7)));


                            var final = string.Format(header, data2);

                            oDataTable.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_All, final);

                            oMatrix.LoadFromDataSource();
                            oMatrix.AutoResizeColumns();
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmHaftalikPlan.Freeze(false);
                        }

                    }
                    break;
                case BoEventTypes.et_DOUBLE_CLICK:
                    if (pVal.ColUID == "#" && !pVal.BeforeAction)
                    {
                        try
                        {
                            frmHaftalikPlan.Freeze(true);
                            string Seviye1 = "";
                            string Seviye2 = "";
                            string Seviye3 = "";

                            Seviye1 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_5").Cells.Item(pVal.Row).Specific).Value;
                            Seviye2 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific).Value == "" ? "0" : ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific).Value;
                            Seviye3 = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_7").Cells.Item(pVal.Row).Specific).Value == "" ? "0" : ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_7").Cells.Item(pVal.Row).Specific).Value;


                            var xml = oDataTable.SerializeAsXML(BoDataTableXmlSelect.dxs_All);
                            var rows2 = getData(xml);


                            var data2 = setDataXML(rows2, Seviye1, Seviye2);

                            var final = string.Format(header, data2);

                            oDataTable.LoadSerializedXML(SAPbouiCOM.BoDataTableXmlSelect.dxs_All, final);

                            oMatrix.LoadFromDataSource();
                            oMatrix.AutoResizeColumns();
                        }
                        catch (Exception ex)
                        {

                        }
                        finally
                        {
                            frmHaftalikPlan.Freeze(false);
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

        public string setDataXML(List<Datas> rows2, string Seviye1, string Seviye2)
        {
            string data2 = "";
            data2 = string.Join("", rows2.Select(s => string.Format(row, s.Grup, s.GrupSira, s.Tur, s.UrunKodu, s.UrunAdi, s.Seviye1, s.Seviye2, s.Seviye3, s.a1, s.b1, s.c1, s.a2, s.b2, s.c2, s.a3, s.b3, s.c3, s.a4, s.b4, s.c4, s.a5, s.b5, s.c5, s.a6, s.b6, s.c6, s.a7, s.b7, s.c7)));


            if (Seviye2 == "0")
            {
                var exist = rows2.Where(x => x.Seviye1 == Seviye1 && x.Seviye2 == Seviye2).Count();
                if (exist > 0)
                {
                    var existProduct = rows2.Where(x => x.Seviye1 == Seviye1 && x.Seviye3 == "3").Count();
                    if (existProduct > 0)
                    {
                        var ddd = rows2.Where(x => x.Seviye1 == Seviye1 && x.Seviye3 == "3").ToList();

                        var finalrows = rows2.Except(ddd).ToList();
                        data2 = string.Join("", finalrows.Select(s => string.Format(row, s.Grup, s.GrupSira, s.Tur, s.UrunKodu, s.UrunAdi, s.Seviye1, s.Seviye2, s.Seviye3, s.a1, s.b1, s.c1, s.a2, s.b2, s.c2, s.a3, s.b3, s.c3, s.a4, s.b4, s.c4, s.a5, s.b5, s.c5, s.a6, s.b6, s.c6, s.a7, s.b7, s.c7)));
                    }
                    else
                    {
                        var ddd = OriginalData.Where(x => x.Seviye1 == Seviye1 && x.Seviye3 == "3").ToList();

                        rows2.AddRange(ddd);

                        var finalrows = rows2.OrderBy(x => x.GrupSira).ToList();


                        data2 = string.Join("", finalrows.Select(s => string.Format(row, s.Grup, s.GrupSira, s.Tur, s.UrunKodu, s.UrunAdi, s.Seviye1, s.Seviye2, s.Seviye3, s.a1, s.b1, s.c1, s.a2, s.b2, s.c2, s.a3, s.b3, s.c3, s.a4, s.b4, s.c4, s.a5, s.b5, s.c5, s.a6, s.b6, s.c6, s.a7, s.b7, s.c7)));
                    }
                }

            }
            else
            {
                var exist = rows2.Where(x => x.Seviye1 == Seviye1 && x.Seviye2 == Seviye2).Count();

                if (exist > 0)
                {
                    var existProduct = rows2.Where(x => x.Seviye1 == Seviye1 && x.Seviye2 == Seviye2 && x.Seviye3 == "3").Count();

                    if (existProduct > 0)
                    {
                        var ddd = rows2.Where(x => x.Seviye1 == Seviye1 && x.Seviye2 == Seviye2 && x.Seviye3 == "3").ToList();

                        var finalrows = rows2.Except(ddd).ToList();
                        data2 = string.Join("", finalrows.Select(s => string.Format(row, s.Grup, s.GrupSira, s.Tur, s.UrunKodu, s.UrunAdi, s.Seviye1, s.Seviye2, s.Seviye3, s.a1, s.b1, s.c1, s.a2, s.b2, s.c2, s.a3, s.b3, s.c3, s.a4, s.b4, s.c4, s.a5, s.b5, s.c5, s.a6, s.b6, s.c6, s.a7, s.b7, s.c7)));
                    }
                    else
                    {
                        var ddd = OriginalData.Where(x => x.Seviye1 == Seviye1 && x.Seviye2 == Seviye2 && x.Seviye3 == "3").ToList();

                        rows2.AddRange(ddd);

                        var finalrows = rows2.OrderBy(x => x.GrupSira).ToList();

                        data2 = string.Join("", finalrows.Select(s => string.Format(row, s.Grup, s.GrupSira, s.Tur, s.UrunKodu, s.UrunAdi, s.Seviye1, s.Seviye2, s.Seviye3, s.a1, s.b1, s.c1, s.a2, s.b2, s.c2, s.a3, s.b3, s.c3, s.a4, s.b4, s.c4, s.a5, s.b5, s.c5, s.a6, s.b6, s.c6, s.a7, s.b7, s.c7)));
                    }
                }
            }

            return data2;

        }
        public List<Datas> getData(string xml)
        {
            var rows2 = (from x in XDocument.Parse(xml).Descendants("Row")
                         select new Datas()
                         {
                             Grup = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "Grup" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "Grup" select new XElement(y.Element("Value"))).First().Value : "",
                             GrupSira = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "GrupSira" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "GrupSira" select new XElement(y.Element("Value"))).First().Value : "",
                             Tur = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "Tur" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "Tur" select new XElement(y.Element("Value"))).First().Value : "",
                             UrunKodu = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "UrunKodu" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "UrunKodu" select new XElement(y.Element("Value"))).First().Value : "",
                             UrunAdi = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "UrunAdi" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "UrunAdi" select new XElement(y.Element("Value"))).First().Value : "",
                             Seviye1 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "Seviye1" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "Seviye1" select new XElement(y.Element("Value"))).First().Value : "0",
                             Seviye2 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "Seviye2" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "Seviye2" select new XElement(y.Element("Value"))).First().Value : "0",
                             Seviye3 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "Seviye3" select new XElement(y.Element("Value"))).Any() == true ? (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "Seviye3" select new XElement(y.Element("Value"))).First().Value : "0",
                             a1 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "a1" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "a1" select new XElement(y.Element("Value"))).First().Value) : 0,
                             b1 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "b1" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "b1" select new XElement(y.Element("Value"))).First().Value) : 0,
                             c1 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "c1" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "c1" select new XElement(y.Element("Value"))).First().Value) : 0,
                             a2 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "a2" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "a2" select new XElement(y.Element("Value"))).First().Value) : 0,
                             b2 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "b2" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "b2" select new XElement(y.Element("Value"))).First().Value) : 0,
                             c2 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "c2" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "c2" select new XElement(y.Element("Value"))).First().Value) : 0,
                             a3 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "a3" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "a3" select new XElement(y.Element("Value"))).First().Value) : 0,
                             b3 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "b3" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "b3" select new XElement(y.Element("Value"))).First().Value) : 0,
                             c3 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "c3" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "c3" select new XElement(y.Element("Value"))).First().Value) : 0,
                             a4 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "a4" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "a4" select new XElement(y.Element("Value"))).First().Value) : 0,
                             b4 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "b4" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "b4" select new XElement(y.Element("Value"))).First().Value) : 0,
                             c4 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "c4" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "c4" select new XElement(y.Element("Value"))).First().Value) : 0,
                             a5 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "a5" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "a5" select new XElement(y.Element("Value"))).First().Value) : 0,
                             b5 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "b5" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "b5" select new XElement(y.Element("Value"))).First().Value) : 0,
                             c5 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "c5" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "c5" select new XElement(y.Element("Value"))).First().Value) : 0,
                             a6 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "a6" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "a6" select new XElement(y.Element("Value"))).First().Value) : 0,
                             b6 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "b6" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "b6" select new XElement(y.Element("Value"))).First().Value) : 0,
                             c6 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "c6" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "c6" select new XElement(y.Element("Value"))).First().Value) : 0,
                             a7 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "a7" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "a7" select new XElement(y.Element("Value"))).First().Value) : 0,
                             b7 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "b7" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "b7" select new XElement(y.Element("Value"))).First().Value) : 0,
                             c7 = (from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "c7" select new XElement(y.Element("Value"))).Any() == true ? parservalues<double>((from y in x.Elements("Cells").Elements("Cell") where y.Element("ColumnUid").Value == "c7" select new XElement(y.Element("Value"))).First().Value) : 0,

                         }).ToList();

            return rows2;
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

        List<Datas> datas = new List<Datas>();
        List<Datas> OriginalData = new List<Datas>();
        public class Datas
        {
            public string Grup { get; set; }
            public string GrupSira { get; set; }
            public string Tur { get; set; }
            public string UrunKodu { get; set; }
            public string UrunAdi { get; set; }
            public string Seviye1 { get; set; }
            public string Seviye2 { get; set; }
            public string Seviye3 { get; set; }
            public double a1 { get; set; }
            public double b1 { get; set; }
            public double c1 { get; set; }
            public double a2 { get; set; }
            public double b2 { get; set; }
            public double c2 { get; set; }
            public double a3 { get; set; }
            public double b3 { get; set; }
            public double c3 { get; set; }
            public double a4 { get; set; }
            public double b4 { get; set; }
            public double c4 { get; set; }
            public double a5 { get; set; }
            public double b5 { get; set; }
            public double c5 { get; set; }
            public double a6 { get; set; }
            public double b6 { get; set; }
            public double c6 { get; set; }
            public double a7 { get; set; }
            public double b7 { get; set; }
            public double c7 { get; set; }
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
