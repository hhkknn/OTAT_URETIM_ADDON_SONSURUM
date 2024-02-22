using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.UVT.SAPB1.Models;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class GirdiKontrolFormu
    {
        [ItemAtt(AIFConn.GirdiKontrolFormuUID)]
        public SAPbouiCOM.Form frmGirdiKntrl;

        [ItemAtt("Item_4")]
        public SAPbouiCOM.Matrix oMatrix;

        [ItemAtt("Item_1")]
        public SAPbouiCOM.EditText oEditTarih;

        [ItemAtt("Item_6")]
        public SAPbouiCOM.EditText oEditGercekSipNo;

        [ItemAtt("Item_3")]
        public SAPbouiCOM.EditText oEditDocEntry;

        [ItemAtt("1")]
        public SAPbouiCOM.Button oBtn;

        [ItemAtt("Item_7")]
        public SAPbouiCOM.EditText oEditTaslakSipNo;

        [ItemAtt("Item_10")]
        public SAPbouiCOM.ComboBox oComboBelgeTipi;

        //[ItemAtt("Item_14")]
        //public SAPbouiCOM.EditText oEditUrunKodu;

        //[ItemAtt("Item_1")]
        //public SAPbouiCOM.EditText oEditUrunAdi;

        public static SAPbouiCOM.DataTable oDataTable = null;
        public static SAPbouiCOM.DBDataSource oDBDSHeader = null;

        //string satinSipNo = "";
        private List<GirdiKontrolParametre> girdiKontrolFormus = new List<GirdiKontrolParametre>();
        public void LoadForms(List<GirdiKontrolParametre> _girdiKontrolFormus)
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.GirdiKontrolFormuXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.GirdiKontrolFormuXML));
            Functions.CreateUserOrSystemFormComponent<GirdiKontrolFormu>(AIFConn.GirdiKntrl);

            //satinSipNo = _satinSipNo;
            girdiKontrolFormus = _girdiKontrolFormus;

            InitForms();
        }

        public void InitForms()
        {
            try
            {

                frmGirdiKntrl.EnableMenu("1283", false);
                frmGirdiKntrl.EnableMenu("1284", false);
                frmGirdiKntrl.EnableMenu("1286", false);

                oDataTable = frmGirdiKntrl.DataSources.DataTables.Add("DATA");

                oEditTarih.Value = DateTime.Now.ToString("yyyyMMdd");
                oMatrix.AddRow();

                if (girdiKontrolFormus != null)
                {
                    frmGirdiKntrl.Freeze(true);

                    string sql = "";

                    if (girdiKontrolFormus[0].BelgeTipi.ToString() == "6") //belge tipi taslak
                    {
                        oEditTaslakSipNo.Value = girdiKontrolFormus[0].SiparisNumarasi.ToString();
                        oComboBelgeTipi.Select("T", BoSearchKey.psk_ByValue);

                        sql = "Select \"DocEntry\" from \"@AIF_GIRDIKONTROL\" where \"U_TaslakSipNo\" = '" + oEditTaslakSipNo.Value + "' AND \"U_BelgeTipi\" = 'T' ";

                    }
                    else
                    {
                        oEditGercekSipNo.Value = girdiKontrolFormus[0].SiparisNumarasi.ToString();
                        oComboBelgeTipi.Select("G", BoSearchKey.psk_ByValue);

                        sql = "Select \"DocEntry\" from \"@AIF_GIRDIKONTROL\" where \"U_GercekSipNo\" = '" + oEditGercekSipNo.Value + "'  AND \"U_BelgeTipi\" = 'G'";

                    }

                    ConstVariables.oRecordset.DoQuery(sql);

                    if (ConstVariables.oRecordset.RecordCount > 0)
                    {
                        oEditDocEntry.Item.Enabled = true;
                        frmGirdiKntrl.Mode = BoFormMode.fm_FIND_MODE;
                        oEditDocEntry.Value = ConstVariables.oRecordset.Fields.Item(0).Value.ToString();
                        oBtn.Item.Click();
                        oEditTarih.Item.Click();
                        oEditDocEntry.Item.Enabled = false;
                    }
                    else
                    {
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_4").Cells.Item(oMatrix.RowCount).Specific).Value = girdiKontrolFormus[0].TedarikciKodu.ToString();
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_5").Cells.Item(oMatrix.RowCount).Specific).Value = girdiKontrolFormus[0].TedarikciAdi.ToString();
                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_6").Cells.Item(oMatrix.RowCount).Specific).Value = girdiKontrolFormus[0].IrsaliyeNo.ToString();
                    }

                }

                oMatrix.SetCellFocus(1, 1);
                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Hata oluştu." + ex.Message);
            }
            finally
            {
                frmGirdiKntrl.Freeze(false);
            }

        }
        public static void CFLFilter(string FormUID, SAPbouiCOM.ItemEvent pVal, string Item_Code, string Cflid, string AliasName, string ikinciFiltre, string ucuncuFiltre, string sql, ref bool calledBefore)
        {
            try
            {
                //if (calledBefore) return;
                SAPbouiCOM.Form oForm = null;
                oForm = Handler.SAPApplication.Forms.Item(FormUID);
                SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                oDataTable = oCFLEvento.SelectedObjects;
                //oDBDSHeader = oForm.DataSources.DBDataSources.Item(Datasource);
                string sorguOnizleme = "";

                if (pVal.BeforeAction)
                {
                    if (pVal.ItemUID == Item_Code || pVal.ColUID == Item_Code)
                    {
                        //SAPbouiCOM.Form oForm = default(SAPbouiCOM.Form);
                        oForm = Handler.SAPApplication.Forms.ActiveForm;
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(Cflid);
                        SAPbouiCOM.Conditions oConds = default(SAPbouiCOM.Conditions);
                        SAPbouiCOM.Condition oCond = default(SAPbouiCOM.Condition);
                        SAPbouiCOM.Conditions oEmptyConds = new SAPbouiCOM.Conditions();

                        var rsetCFL = ConstVariables.oRecordset;
                        oCFL.SetConditions(oEmptyConds);
                        oConds = oCFL.GetConditions();


                        rsetCFL.DoQuery(sql);
                        rsetCFL.MoveFirst();

                        for (int i = 1; i <= rsetCFL.RecordCount; i++)
                        {
                            if (i == (rsetCFL.RecordCount))
                            {
                                if (rsetCFL.RecordCount > 1) //sorgudan gelen satır sayısı 1 den büyük ise or yapar
                                {
                                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                                    sorguOnizleme += " '" + oCond.Relationship + "' ";
                                }

                                oCond = oConds.Add();
                                oCond.Alias = AliasName;
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCond.CondVal = (rsetCFL.Fields.Item(0).Value).ToString().Trim();

                                sorguOnizleme += " '" + oCond.Alias + "' " + " '" + oCond.CondVal + "'";

                                if (rsetCFL.RecordCount > 1)
                                {
                                    oCond.BracketCloseNum = 1; //or parantezini kapatır
                                    sorguOnizleme += " ) ";
                                }
                            }
                            else
                            {
                                if (i > 1)
                                {
                                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                                    sorguOnizleme += " '" + oCond.Relationship + "' ";
                                }

                                oCond = oConds.Add();

                                if (i == 1)
                                {
                                    oCond.BracketOpenNum = 1;
                                    sorguOnizleme += " ( ";

                                }

                                oCond.Alias = AliasName;
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCond.CondVal = (rsetCFL.Fields.Item(0).Value).ToString().Trim();

                                sorguOnizleme += " '" + oCond.Alias + "' " + " '" + oCond.CondVal + "'";

                            }

                            rsetCFL.MoveNext();
                        }

                        if (rsetCFL.RecordCount == 0)
                        {
                            oCond = oConds.Add();
                            oCond.Alias = AliasName;
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                            oCond.CondVal = "-1";

                            sorguOnizleme += " '" + oCond.Alias + "' " + " '" + oCond.Relationship + "' " + " " + " '" + oCond.CondVal + "'";

                        }

                        rsetCFL.MoveFirst(); //son kayıtı başa almaya yarar
                        if (ikinciFiltre != "")
                        {
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                            sorguOnizleme += " '" + oCond.Relationship + "' ";

                            oCond = oConds.Add();
                            oCond.Alias = ikinciFiltre;
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCond.CondVal = (rsetCFL.Fields.Item(1).Value).ToString().Trim();

                            sorguOnizleme += " '" + oCond.Alias + "' " + " '" + oCond.CondVal + "'";

                        }

                        rsetCFL.MoveFirst();
                        if (ucuncuFiltre != "")
                        {
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                            sorguOnizleme += " '" + oCond.Relationship + "' ";

                            oCond = oConds.Add();
                            oCond.Alias = ucuncuFiltre;
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCond.CondVal = (rsetCFL.Fields.Item(2).Value).ToString().Trim();

                            sorguOnizleme += " '" + oCond.Alias + "' " + " '" + oCond.CondVal + "'";
                        }

                        oCFL.SetConditions(oConds);

                    }
                }

                sorguOnizleme = "";

                //if (pVal.BeforeAction == false)
                //{
                //    if (pVal.ItemUID == Item_Code || pVal.ColUID == Item_Code)
                //    {
                //        ((SAPbouiCOM.EditText)oForm.Items.Item(Item_Code).Specific).Value = oDataTable.GetValue(AliasName, 0).ToString();
                //    }
                //}
                //calledBefore = true;
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
                        string sonsatir1 = frmGirdiKntrl.DataSources.DBDataSources.Item(1).GetValue("U_GirdiKodu", frmGirdiKntrl.DataSources.DBDataSources.Item(1).Size - 1);

                        if (sonsatir1 == "")
                        {
                            frmGirdiKntrl.DataSources.DBDataSources.Item(1).RemoveRecord(frmGirdiKntrl.DataSources.DBDataSources.Item(1).Size - 1);
                        }
                    }
                    break;
                case BoEventTypes.et_FORM_DATA_UPDATE:
                    if (BusinessObjectInfo.BeforeAction)
                    {
                        string sonsatir1 = frmGirdiKntrl.DataSources.DBDataSources.Item(1).GetValue("U_GirdiKodu", frmGirdiKntrl.DataSources.DBDataSources.Item(1).Size - 1);

                        if (sonsatir1 == "")
                        {
                            frmGirdiKntrl.DataSources.DBDataSources.Item(1).RemoveRecord(frmGirdiKntrl.DataSources.DBDataSources.Item(1).Size - 1);
                        }
                    }
                    break;
                case BoEventTypes.et_FORM_DATA_DELETE:
                    break;
                case BoEventTypes.et_FORM_DATA_LOAD:
                    //if (!BusinessObjectInfo.BeforeAction)
                    //{
                    //    frmAnaliz.DataSources.DBDataSources.Item(1).Clear();
                    //    frmAnaliz.DataSources.DBDataSources.Item(2).Clear();
                    //    frmAnaliz.DataSources.DBDataSources.Item(3).Clear();

                    //    oMatrixSonUrunKimyasal.AddRow();
                    //    oMatrixDuyusalFonskiyonel.AddRow();
                    //    oMatrixMikrobiyolojik.AddRow();

                    //    oMatrixSonUrunKimyasal.AutoResizeColumns();
                    //    oMatrixDuyusalFonskiyonel.AutoResizeColumns();
                    //    oMatrixMikrobiyolojik.AutoResizeColumns();
                    //}
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
                    if (pVal.ItemUID == "2" && !pVal.BeforeAction)
                    {
                        try
                        {
                            frmGirdiKntrl.Close();
                        }
                        catch (Exception)
                        {

                        }
                    }
                    //if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    //{
                    //    frmAnaliz.DataSources.DBDataSources.Item(1).Clear();
                    //    frmAnaliz.DataSources.DBDataSources.Item(2).Clear();
                    //    frmAnaliz.DataSources.DBDataSources.Item(3).Clear();

                    //    oMatrixSonUrunKimyasal.AddRow();
                    //    oMatrixDuyusalFonskiyonel.AddRow();
                    //    oMatrixMikrobiyolojik.AddRow();

                    //    oMatrixSonUrunKimyasal.AutoResizeColumns();
                    //    oMatrixDuyusalFonskiyonel.AutoResizeColumns();
                    //    oMatrixMikrobiyolojik.AutoResizeColumns();
                    //}
                    //if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    //{
                    //    if (itemUID == "Item" && pVal.ColUID == "")
                    //    {
                    //        if (silinecekler.Count > 0)
                    //        {
                    //            foreach (var item in silinecekler)
                    //            {
                    //                if (item != "")
                    //                {
                    //                    ConstVariables.oRecordset.DoQuery("Delete from \"@AIF_ANALIZGIRIS1\" where \"DocEntry\" = '" + item + "'");
                    //                }
                    //            }
                    //        }

                    //        silinecekler = new List<string>(); 
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
                    break;
                case BoEventTypes.et_DOUBLE_CLICK:
                    if (pVal.ItemUID == "Item_4" && pVal.ColUID == "Col_2" && !pVal.BeforeAction)
                    {
                        //string sql = "SELECT T0.\"ItemCode\" as \"KalemNumarası\",T0.\"ItemName\" as \"KalemTanımı\",T0.\"BatchNum\" as \"PartiNumarası\", T0.\"WhsCode\" as \"DepoKodu\", T2.\"WhsName\" as \"DepoAdı\", T0.\"Quantity\" as \"Miktar\", T1.\"InDate\" as \"ÜretimTarihi\",T1.\"ExpDate\" as \"SonKullanmaTarihi\" FROM IBT1 T0 LEFT JOIN OBTN T1 ON T0.\"BatchNum\" = T1.\"DistNumber\" AND T0.\"ItemCode\" = T1.\"ItemCode\" LEFT JOIN OWHS T2 ON T0.\"WhsCode\" = T2.\"WhsCode\" WHERE T0.\"BaseType\" = '20' ";

                        //sql += " AND T0.\"BaseEntry\" = '" + oEditSatinSipNo.Value.ToString() + "' ";

                        //oDataTable.Clear();
                        //oDataTable.ExecuteQuery(sql);

                        //oMatrix.Clear();

                        //oMatrix.Columns.Item("Col_0").DataBind.Bind("DATA", "KalemNumarası");
                        //oMatrix.Columns.Item("Col_1").DataBind.Bind("DATA", "KalemTanımı");
                        //oMatrix.Columns.Item("Col_2").DataBind.Bind("DATA", "PartiNumarası");
                        //oMatrix.Columns.Item("Col_3").DataBind.Bind("DATA", "Miktar");
                        //oMatrix.Columns.Item("Col_7").DataBind.Bind("DATA", "ÜretimTarihi");
                        //oMatrix.Columns.Item("Col_8").DataBind.Bind("DATA", "SonKullanmaTarihi");
                        //oMatrix.Columns.Item("Col_9").DataBind.Bind("DATA", "DepoKodu");
                        //oMatrix.Columns.Item("Col_10").DataBind.Bind("DATA", "DepoAdı");

                        //oMatrix.LoadFromDataSource();
                        //oMatrix.AutoResizeColumns();

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
                    if (pVal.ItemUID == "Item_4" && pVal.BeforeAction)
                    {
                        if (pVal.ColUID == "Col_0")
                        {
                            try
                            {
                                bool cflCalled = false;
                                string sqlQuery = "";

                                if (girdiKontrolFormus[0].BelgeTipi.ToString() == "6") //belge tipi taslak
                                {
                                    sqlQuery = "select \"ItemCode\",\"Dscription\" from DRF1 where \"DocEntry\"='" + girdiKontrolFormus[0].SiparisNumarasi + "' ";
                                }
                                else
                                {
                                    sqlQuery = "select \"ItemCode\",\"Dscription\" from PDN1 where \"DocEntry\"='" + girdiKontrolFormus[0].SiparisNumarasi + "' ";

                                }

                                CFLFilter(frmGirdiKntrl.UniqueID, pVal, "Item_4", "CFL_0", "ItemCode", "", "", sqlQuery, ref cflCalled);
                            }
                            catch (Exception)
                            {
                            }
                        }
                        if (pVal.ColUID == "Col_2")
                        {
                            //try
                            //{
                            //    bool cflCalled = false;

                            //    //string sqlQuery = "select \"ItemCode\",\"ItemName\" from PDN1 where \"DocEntry\"='" + girdiKontrolFormus[0].SatinalmaSipNo + "' ";

                            //    string itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value;
                            //    string sql = "";

                            //    sql = " Select * from (SELECT T0.\"BatchNum\",T0.\"ItemCode\",T0.\"ItemName\" as \"KalemTanımı\", T0.\"WhsCode\" as \"DepoKodu\", T2.\"WhsName\" as \"DepoAdı\", T0.\"Quantity\" as \"Miktar\", T1.\"InDate\" as \"ÜretimTarihi\",T1.\"ExpDate\" as \"SonKullanmaTarihi\" FROM IBT1 T0 LEFT JOIN OBTN T1 ON T0.\"BatchNum\" = T1.\"DistNumber\" AND T0.\"ItemCode\" = T1.\"ItemCode\" LEFT JOIN OWHS T2 ON T0.\"WhsCode\" = T2.\"WhsCode\" WHERE T0.\"BaseType\" = '20' ";

                            //    if (oEditGercekSipNo.Value.ToString() != "")
                            //    {
                            //        sql += " AND T0.\"BaseEntry\" = '" + oEditGercekSipNo.Value.ToString() + "' ";
                            //    }

                            //    if (oEditTaslakSipNo.Value.ToString() != "")
                            //    {
                            //        sql += " AND T0.\"BaseEntry\" = '" + oEditTaslakSipNo.Value.ToString() + "' ";
                            //    }

                            //    sql += " AND T0.\"ItemCode\" = '" + itemCode + "'";

                            //    sql += "  union all ";

                            //    sql += " select T3.\"DistNumber\" as \"BatchNum\",T3.\"ItemCode\",T3.\"itemName\" AS \"KalemTanımı\",T2.\"WhsCode\" AS \"DepoKodu\",(select T5.\"WhsName\" from \"OWHS\" T5 Where T5.\"WhsCode\" = T2.\"WhsCode\") AS \"DepoAdı\", T2.\"Quantity\" as \"Miktar\",  T3.\"InDate\" as \"UretimTarihi\",T3.\"ExpDate\" as \"SonKullanmaTarihi\"  from \"DRF16\" T2 inner join \"ODBN\" T3 on T2.\"ObjAbs\" = T3.\"AbsEntry\" where T2.\"ItemCode\" = '" + itemCode + "' ";

                            //    if (oEditGercekSipNo.Value.ToString() != "")
                            //    {
                            //        sql += " AND T2.\"AbsEntry\"= '" + oEditGercekSipNo.Value.ToString() + "' ";
                            //    }

                            //    if (oEditTaslakSipNo.Value.ToString() != "")
                            //    {
                            //        sql += " AND T2.\"AbsEntry\" = '" + oEditTaslakSipNo.Value.ToString() + "' ";
                            //    }

                            //    sql += "  ) as tbl";

                            //    //string jj = "Select * from OIBT";

                            //    CFLFilter(frmGirdiKntrl.UniqueID, pVal, "Item_4", "CFL_1", "BatchNum", "", "", jj, ref cflCalled);

                            //    //string sql = "SELECT T0.\"ItemCode\" as \"KalemNumarası\",T0.\"ItemName\" as \"KalemTanımı\",T0.\"BatchNum\" as \"PartiNumarası\", T0.\"WhsCode\" as \"DepoKodu\", T2.\"WhsName\" as \"DepoAdı\", T0.\"Quantity\" as \"Miktar\", T1.\"InDate\" as \"ÜretimTarihi\",T1.\"ExpDate\" as \"SonKullanmaTarihi\" FROM IBT1 T0 LEFT JOIN OBTN T1 ON T0.\"BatchNum\" = T1.\"DistNumber\" AND T0.\"ItemCode\" = T1.\"ItemCode\" LEFT JOIN OWHS T2 ON T0.\"WhsCode\" = T2.\"WhsCode\" WHERE T0.\"BaseType\" = '20' ";

                            //    //sql += " AND T0.\"BaseEntry\" = '" + oEditSatinSipNo.Value.ToString() + "' ";

                            //    //oDataTable.Clear();
                            //    //oDataTable.ExecuteQuery(sql);

                            //    //oMatrix.Clear();

                            //    //oMatrix.Columns.Item("Col_0").DataBind.Bind("DATA", "KalemNumarası");
                            //    //oMatrix.Columns.Item("Col_1").DataBind.Bind("DATA", "KalemTanımı");
                            //    //oMatrix.Columns.Item("Col_2").DataBind.Bind("DATA", "PartiNumarası");
                            //    //oMatrix.Columns.Item("Col_3").DataBind.Bind("DATA", "Miktar");
                            //    //oMatrix.Columns.Item("Col_7").DataBind.Bind("DATA", "ÜretimTarihi");
                            //    //oMatrix.Columns.Item("Col_8").DataBind.Bind("DATA", "SonKullanmaTarihi");
                            //    //oMatrix.Columns.Item("Col_9").DataBind.Bind("DATA", "DepoKodu");
                            //    //oMatrix.Columns.Item("Col_10").DataBind.Bind("DATA", "DepoAdı");

                            //    //oMatrix.LoadFromDataSource();
                            //    //oMatrix.AutoResizeColumns();
                            //}
                            //catch (Exception)
                            //{
                            //}
                        }
                    }
                    else if (pVal.ItemUID == "Item_4" && !pVal.BeforeAction)
                    {
                        if (pVal.ColUID == "Col_0")
                        {

                            SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                            if (oDataTable == null)
                            {
                                break;
                            }
                            string val = "";
                            try
                            {
                                val = oDataTable.GetValue("ItemCode", 0).ToString();
                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value = val;
                            }
                            catch (Exception)
                            {
                            }

                            try
                            {
                                val = oDataTable.GetValue("ItemName", 0).ToString();
                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = val;
                            }
                            catch (Exception)
                            {
                            }

                            oMatrix.AutoResizeColumns();

                        }

                        if (pVal.ColUID == "Col_2")
                        {
                            try
                            {

                                SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                                if (oDataTable == null)
                                {
                                    break;
                                }

                                string val = "";
                                try
                                {
                                    val = oDataTable.GetValue("BatchNum", 0).ToString();
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific).Value = val;
                                }
                                catch (Exception)
                                {
                                }

                                val = "";
                                try
                                {
                                    val = oDataTable.GetValue("WhsCode", 0).ToString();
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific).Value = val;
                                }
                                catch (Exception)
                                {
                                }


                                try
                                {

                                    ConstVariables.oRecordset.DoQuery("Select \"WhsName\" from \"OWHS\" where \"WhsCode\" = '" + val + "'");


                                    val = ConstVariables.oRecordset.Fields.Item(0).Value.ToString();
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_10").Cells.Item(pVal.Row).Specific).Value = val;
                                }
                                catch (Exception)
                                {
                                }


                                val = "";
                                try
                                {


                                    val = oDataTable.GetValue("InDate", 0).ToString();

                                    if (val != "" && Convert.ToDateTime(val).Year.ToString() != "1900")
                                    {
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_7").Cells.Item(pVal.Row).Specific).Value = Convert.ToDateTime(val).ToString("yyyyMMdd");//üretim tarihi
                                    }
                                }
                                catch (Exception)
                                {
                                }

                                val = "";
                                try
                                {
                                    string sql = "";

                                    sql = "SELECT ISNULL(T1.\"ExpDate\",'1900-01-01') as \"SonKullanmaTarihi\" FROM IBT1 T0 LEFT JOIN OBTN T1 ON T0.\"BatchNum\" = T1.\"DistNumber\" AND T0.\"ItemCode\" = T1.\"ItemCode\" LEFT JOIN OWHS T2 ON T0.\"WhsCode\" = T2.\"WhsCode\" WHERE T0.\"BaseType\" = '20' ";

                                    if (oEditGercekSipNo.Value.ToString() != "")
                                    {
                                        sql += " AND T0.\"BaseEntry\" = '" + oEditGercekSipNo.Value.ToString() + "' ";
                                    }

                                    if (oEditTaslakSipNo.Value.ToString() != "")
                                    {
                                        sql += " AND T0.\"BaseEntry\" = '" + oEditTaslakSipNo.Value.ToString() + "' ";
                                    }

                                    sql += " AND T0.\"ItemCode\" = '" + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value + "'";

                                    ConstVariables.oRecordset.DoQuery(sql);


                                    val = ConstVariables.oRecordset.Fields.Item(0).Value.ToString();

                                    //DateTime dtTarih = new DateTime(Convert.ToInt32(val.Substring(0, 4)), Convert.ToInt32(val.Substring(4, 2)), Convert.ToInt32(val.Substring(6, 2)));
                                    if (val != "" && Convert.ToDateTime(val).Year.ToString() != "1900")
                                    {

                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_8").Cells.Item(pVal.Row).Specific).Value = Convert.ToDateTime(val).ToString("yyyyMMdd"); //son kull.tarihi
                                    }
                                }
                                catch (Exception)
                                {
                                }


                                oMatrix.AutoResizeColumns();
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }

                    //else if (pVal.ItemUID == "Item_14" && !pVal.BeforeAction)
                    //{
                    //    SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                    //    if (oDataTable == null)
                    //    {
                    //        break;
                    //    }
                    //    string val = "";
                    //    try
                    //    {
                    //        val = oDataTable.GetValue("ItemCode", 0).ToString();
                    //    }
                    //    catch (Exception)
                    //    {
                    //    }

                    //    try
                    //    {
                    //        oEditUrunKodu.Value = val;
                    //        //((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value = val;
                    //    }
                    //    catch (Exception)
                    //    {
                    //    }

                    //    try
                    //    {
                    //        val = oDataTable.GetValue("ItemName", 0).ToString();
                    //    }
                    //    catch (Exception)
                    //    {
                    //    }

                    //    try
                    //    {
                    //        oEditUrunAdi.Value = val;
                    //        //((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = val;
                    //    }
                    //    catch (Exception)
                    //    {
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
        List<string> silinecekler = new List<string>();
        string itemUID = "";
        public void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.MenuUID == "1282" && !pVal.BeforeAction)
            {
                try
                {
                    oMatrix.AddRow();
                }
                catch (Exception)
                {
                }
            }

            try
            {
                if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && pVal.BeforeAction)
                {
                    int row = oMatrix.GetNextSelectedRow();
                    if (row != -1)
                    {
                        //if (((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_5").Cells.Item(row).Specific).Value != "")
                        //{
                        //    silinecekler.Add(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_5").Cells.Item(row).Specific).Value);
                        //}

                        oMatrix.DeleteRow(row);

                        if (frmGirdiKntrl.Mode == BoFormMode.fm_OK_MODE)
                        {
                            frmGirdiKntrl.Mode = BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                }
                else if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && pVal.BeforeAction)
                {
                    frmGirdiKntrl.DataSources.DBDataSources.Item("@AIF_GIRDIKONTROL1").Clear();
                    oMatrix.AddRow();

                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_4").Cells.Item(oMatrix.RowCount).Specific).Value = girdiKontrolFormus[0].TedarikciKodu.ToString();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_5").Cells.Item(oMatrix.RowCount).Specific).Value = girdiKontrolFormus[0].TedarikciAdi.ToString();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_6").Cells.Item(oMatrix.RowCount).Specific).Value = girdiKontrolFormus[0].IrsaliyeNo.ToString();

                    if (frmGirdiKntrl.Mode == BoFormMode.fm_OK_MODE)
                    {
                        frmGirdiKntrl.Mode = BoFormMode.fm_UPDATE_MODE;
                    }
                }
            }
            catch (Exception)
            {
            }
        }

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
    }
}
