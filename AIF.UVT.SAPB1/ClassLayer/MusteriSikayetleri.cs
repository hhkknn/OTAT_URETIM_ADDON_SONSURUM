using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.UVT.SAPB1.HelperClass;
using AIF.UVT.SAPB1.Models;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class MusteriSikayetleri
    {
        [ItemAtt(AIFConn.MusteriSikayetleriUID)]
        public SAPbouiCOM.Form frmMusteriSikayetleri;

        [ItemAtt("Item_1")]
        public SAPbouiCOM.EditText oEditDocEntry;

        [ItemAtt("Item_4")]
        public SAPbouiCOM.EditText oEditBaslangicTarihi;

        [ItemAtt("Item_6")]
        public SAPbouiCOM.EditText oEditBitisTarihi;

        [ItemAtt("Item_7")]
        public SAPbouiCOM.Button oBtnListele;

        [ItemAtt("Item_16")]
        public SAPbouiCOM.Matrix oMatrixDetay;

        [ItemAtt("Item_11")]
        public SAPbouiCOM.Matrix oMatrixEkler;

        [ItemAtt("Item_12")]
        public SAPbouiCOM.Button oBtnEkle;

        [ItemAtt("Item_13")]
        public SAPbouiCOM.Button oBtnGoruntule;

        [ItemAtt("Item_14")]
        public SAPbouiCOM.Button oBtnSil;

        [ItemAtt("Item_117")]
        public SAPbouiCOM.Folder oFolderMusteriSikayet;

        public static SAPbouiCOM.DataTable oDataTable = null;
        public static SAPbouiCOM.DBDataSource oDBDSHeader = null;
        List<string> silinecekDosya = new List<string>();
        bool chooseFromListSeciliyor = false;
        string secilmisUygunsuzlukNedeni = "";
        bool hesaplamayapiliyorIadeKG = false;

        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.MusteriSikayetleriXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.MusteriSikayetleriXML));
            Functions.CreateUserOrSystemFormComponent<MusteriSikayetleri>(AIFConn.MustSkayet);

            InitForms();
        }

        public void InitForms()
        {
            try
            {
                frmMusteriSikayetleri.EnableMenu("1283", false);
                frmMusteriSikayetleri.EnableMenu("1284", false);
                frmMusteriSikayetleri.EnableMenu("1286", false);
                oDataTable = frmMusteriSikayetleri.DataSources.DataTables.Add("DATA");


                frmMusteriSikayetleri.Items.Item("Item_117").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                #region şikayet grubu comobbox doldurma
                SAPbouiCOM.Column oCol = (SAPbouiCOM.Column)oMatrixDetay.Columns.Item("Col_13");

                string sql = "Select \"U_Id\", \"U_SkytGrubu\"  from \"@AIF_MUSTSKYTGRB\" ";
                ConstVariables.oRecordset.DoQuery(sql);

                if (ConstVariables.oRecordset.RecordCount > 0)
                {
                    while (!ConstVariables.oRecordset.EoF)
                    {
                        oCol.ValidValues.Add(ConstVariables.oRecordset.Fields.Item("U_Id").Value.ToString(), ConstVariables.oRecordset.Fields.Item("U_SkytGrubu").Value.ToString());
                        ConstVariables.oRecordset.MoveNext();
                    }
                }
                #endregion


                oMatrixDetay.AddRow();
                //oMatrixEkler.AddRow();

                oMatrixDetay.AutoResizeColumns();
                oMatrixEkler.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Hata oluştu" + ex.Message);

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
                        SAPbouiCOM.Conditions oConds = new SAPbouiCOM.Conditions();
                        SAPbouiCOM.Condition oCond = oConds.Add();
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

                if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == Item_Code || pVal.ColUID == Item_Code)
                    {
                        ((SAPbouiCOM.EditText)oForm.Items.Item(Item_Code).Specific).Value = oDataTable.GetValue(AliasName, 0).ToString();
                    }
                }
                calledBefore = true;
            }
            catch (Exception)
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
                        string sonsatir1 = frmMusteriSikayetleri.DataSources.DBDataSources.Item(1).GetValue("U_UrunGrubu", frmMusteriSikayetleri.DataSources.DBDataSources.Item(1).Size - 1);
                        string sonsatir2 = frmMusteriSikayetleri.DataSources.DBDataSources.Item(2).GetValue("U_KaynakYol", frmMusteriSikayetleri.DataSources.DBDataSources.Item(2).Size - 1);

                        if (sonsatir1 == "")
                        {
                            frmMusteriSikayetleri.DataSources.DBDataSources.Item(1).RemoveRecord(frmMusteriSikayetleri.DataSources.DBDataSources.Item(1).Size - 1);
                        }

                        if (sonsatir2 == "")
                        {
                            frmMusteriSikayetleri.DataSources.DBDataSources.Item(2).RemoveRecord(frmMusteriSikayetleri.DataSources.DBDataSources.Item(2).Size - 1);
                        }
                    }
                    break;
                case BoEventTypes.et_FORM_DATA_UPDATE:
                    if (BusinessObjectInfo.BeforeAction)
                    {
                        string sonsatir1 = frmMusteriSikayetleri.DataSources.DBDataSources.Item(1).GetValue("U_UrunGrubu", frmMusteriSikayetleri.DataSources.DBDataSources.Item(1).Size - 1);
                        string sonsatir2 = frmMusteriSikayetleri.DataSources.DBDataSources.Item(2).GetValue("U_KaynakYol", frmMusteriSikayetleri.DataSources.DBDataSources.Item(2).Size - 1);

                        if (sonsatir1 == "")
                        {
                            frmMusteriSikayetleri.DataSources.DBDataSources.Item(1).RemoveRecord(frmMusteriSikayetleri.DataSources.DBDataSources.Item(1).Size - 1);
                        }

                        if (sonsatir2 == "")
                        {
                            frmMusteriSikayetleri.DataSources.DBDataSources.Item(2).RemoveRecord(frmMusteriSikayetleri.DataSources.DBDataSources.Item(2).Size - 1);
                        }
                    }
                    break;
                case BoEventTypes.et_FORM_DATA_DELETE:
                    break;
                case BoEventTypes.et_FORM_DATA_LOAD:
                    if (!BusinessObjectInfo.BeforeAction)
                    {
                        frmMusteriSikayetleri.DataSources.DBDataSources.Item(1).Clear();
                        frmMusteriSikayetleri.DataSources.DBDataSources.Item(2).Clear();

                        oMatrixDetay.AddRow();
                        //oMatrixEkler.AddRow();

                        oMatrixDetay.AutoResizeColumns();
                        oMatrixEkler.AutoResizeColumns();
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

                        frmMusteriSikayetleri.DataSources.DBDataSources.Item(1).Clear();
                        frmMusteriSikayetleri.DataSources.DBDataSources.Item(2).Clear();

                        oMatrixDetay.AddRow();
                        //oMatrixEkler.AddRow();

                        oMatrixDetay.AutoResizeColumns();
                        oMatrixEkler.AutoResizeColumns();
                    }
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
                    if (pVal.ItemUID == "Item_16" && pVal.ColUID == "Col_18" && !pVal.BeforeAction && chooseFromListSeciliyor)
                    {
                        try
                        {
                            string cardcode = ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_18").Cells.Item(pVal.Row).Specific).Value;

                            if (cardcode != "")
                            {
                                string sql = "SELECT  T0.\"LineNum\",T0.\"Address\" FROM CRD1 T0 WHERE T0.\"CardCode\" ='" + cardcode + "' ";

                                SAPbouiCOM.ColumnClass oColumn = (SAPbouiCOM.ColumnClass)oMatrixDetay.Columns.Item("Col_2");

                                oColumn.ValidValues.Add("", "");

                                ConstVariables.oRecordset.DoQuery(sql);

                                if (ConstVariables.oRecordset.RecordCount > 0)
                                {
                                    while (!ConstVariables.oRecordset.EoF)
                                    {
                                        try
                                        {
                                            oColumn.ValidValues.Add(ConstVariables.oRecordset.Fields.Item(0).Value.ToString(), ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                                        }
                                        catch (Exception)
                                        {
                                        }

                                        ConstVariables.oRecordset.MoveNext();
                                    }
                                }
                            }
                            oMatrixDetay.AutoResizeColumns();
                            chooseFromListSeciliyor = false;

                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    else if (pVal.ItemUID == "Item_16" && pVal.ColUID == "Col_0" && !pVal.BeforeAction)
                    {
                        int row = oMatrixDetay.GetNextSelectedRow();
                        if (row != -1)
                        {
                            if (((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_0").Cells.Item(row).Specific).Value != "")
                            {
                                //if (((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_0").Cells.Item(oMatrixDetay.RowCount - 1).Specific).Value == "")
                                //{
                                frmMusteriSikayetleri.DataSources.DBDataSources.Item("@AIF_MUSTERISIKAYET1").Clear();
                                oMatrixDetay.AddRow();
                                //}
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Item_16" && pVal.ColUID == "Col_17" && !pVal.BeforeAction && !hesaplamayapiliyorIadeKG)
                    {
                        #region iade kg
                        double sonuc = 0;
                        string sql = "Select \"IWeight1\" from OITM WHERE \"ItemCode\" = '" + ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific).Value + "' ";
                        ConstVariables.oRecordset.DoQuery(sql);

                        if (ConstVariables.oRecordset.RecordCount > 0)
                        {

                            double iadekg = ConstVariables.oRecordset.Fields.Item(0).Value.ToString() == "" ? 0 : Convert.ToDouble(ConstVariables.oRecordset.Fields.Item(0).Value);
                            //double iadekg = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_17").Cells.Item(pVal.Row).Specific).Value);
                            sonuc = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_17").Cells.Item(pVal.Row).Specific).Value) * (iadekg / 1000); //aide adet
                            sonuc = parseNumber.parservalues<double>(sonuc.ToString());

                            try
                            {
                                //oEditUrunAdi.Value = val;
                                hesaplamayapiliyorIadeKG = true;
                                ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_12").Cells.Item(pVal.Row).Specific).Value = sonuc.ToString(System.Globalization.CultureInfo.InvariantCulture); //iade kg
                                hesaplamayapiliyorIadeKG = false;
                            }
                            catch (Exception)
                            {
                            }
                        }
                        #endregion

                    }
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    if (pVal.ItemUID == "Item_16" && pVal.ColUID == "Col_13" && pVal.BeforeAction)
                    {
                        secilmisUygunsuzlukNedeni = ((SAPbouiCOM.ComboBox)oMatrixDetay.Columns.Item("Col_13").Cells.Item(pVal.Row).Specific).Value.ToString();
                    }
                    else if (pVal.ItemUID == "Item_16" && pVal.ColUID == "Col_13" && !pVal.BeforeAction)
                    {
                        string uygunsuzlukNedeni = ((SAPbouiCOM.ComboBox)oMatrixDetay.Columns.Item("Col_13").Cells.Item(pVal.Row).Specific).Value.ToString();

                        if (secilmisUygunsuzlukNedeni != "" && uygunsuzlukNedeni != secilmisUygunsuzlukNedeni)
                        {
                            ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_14").Cells.Item(pVal.Row).Specific).Value = "";
                        }
                    }
                    break;
                case BoEventTypes.et_CLICK:
                    if ((pVal.ItemUID == "Item_12") && pVal.BeforeAction) //ekle btn
                    {
                        if (frmMusteriSikayetleri.Mode == BoFormMode.fm_VIEW_MODE)
                        {
                            //view modda tıklama yapılmasın
                            return false;
                        }
                        try
                        {
                            if (ConstVariables.oCompanyObject.AttachMentPath == "")
                            {
                                Handler.SAPApplication.MessageBox("AttachMentPath klasörü bulunamadığından işleme devam edilemez.");
                                return false;
                            }

                            int rowEkler = oMatrixEkler.GetNextSelectedRow();

                            var fssservice = new AttachmentCreate();
                            string dosya = fssservice.showOpenFileDialog();
                            string kaynakYolEkler = "";
                            if (dosya == "")
                            {
                                return false;
                            }
                            else
                            {
                                //string kaynakYolEkler = file.FileName;
                                //string DosyaAdiEkler = file.SafeFileName;

                                var val = dosya.Split('|');

                                kaynakYolEkler = val[0];
                                string DosyaAdiEkler = val[1];

                                frmMusteriSikayetleri.Freeze(true);

                                //if (rowEkler != -1)
                                //{
                                oMatrixEkler.AddRow();
                                ((SAPbouiCOM.EditText)oMatrixEkler.Columns.Item("Col_0").Cells.Item(oMatrixEkler.RowCount).Specific).Value = kaynakYolEkler;
                                ((SAPbouiCOM.EditText)oMatrixEkler.Columns.Item("Col_1").Cells.Item(oMatrixEkler.RowCount).Specific).Value = ConstVariables.oCompanyObject.AttachMentPath + DosyaAdiEkler;
                                ((SAPbouiCOM.EditText)oMatrixEkler.Columns.Item("Col_2").Cells.Item(oMatrixEkler.RowCount).Specific).Value = DosyaAdiEkler;
                                ((SAPbouiCOM.EditText)oMatrixEkler.Columns.Item("Col_3").Cells.Item(oMatrixEkler.RowCount).Specific).Value = DateTime.Now.ToString("yyyyMMdd");


                                //siranumarasiyazEk();
                                oMatrixEkler.AutoResizeColumns();



                                if (frmMusteriSikayetleri.Mode == BoFormMode.fm_OK_MODE)
                                {
                                    frmMusteriSikayetleri.Mode = BoFormMode.fm_UPDATE_MODE;
                                }
                            }

                            var fi = new FileInfo(kaynakYolEkler);
                        }
                        catch (Exception ex)
                        {
                        }
                        finally
                        {
                            frmMusteriSikayetleri.Freeze(false);
                        }

                    }
                    else if ((pVal.ItemUID == "Item_13") && pVal.BeforeAction)//görüntüle btn
                    {
                        try
                        {
                            //int selrow1 = oMatrixEkler.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder); //seçili satır numarası
                            if (ConstVariables.oCompanyObject.AttachMentPath == "")
                            {
                                Handler.SAPApplication.MessageBox("AttachMentPath klasörü bulunamadığından işleme devam edilemez.");
                                return false;
                            }
                            string hedefYolEkler = "";

                            int rowEkler = oMatrixEkler.GetNextSelectedRow();
                            if (rowEkler != -1)
                            {
                                hedefYolEkler = ((SAPbouiCOM.EditText)oMatrixEkler.Columns.Item("Col_1").Cells.Item(rowEkler).Specific).Value;
                            }
                            else if (rowEkler == -1)
                            {
                                Handler.SAPApplication.MessageBox("Satır seçimi yapılmadı.");
                                return false;
                            }

                            FileInfo fi = new FileInfo(hedefYolEkler);
                            if (fi.Exists)
                            {
                                System.Diagnostics.Process.Start(hedefYolEkler);
                            }
                            else
                            {
                                Handler.SAPApplication.MessageBox("Dosya bulunamadı.");
                            }


                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if ((pVal.ItemUID == "Item_14") && pVal.BeforeAction) //sil btn
                    {
                        if (frmMusteriSikayetleri.Mode == BoFormMode.fm_VIEW_MODE)
                        {
                            //view modda tıklama yapılmasın
                            return false;
                        }

                        try
                        {
                            if (ConstVariables.oCompanyObject.AttachMentPath == "")
                            {
                                Handler.SAPApplication.MessageBox("AttachMentPath klasörü bulunamadığından işleme devam edilemez.");
                                return false;
                            }
                            int rowEkler = oMatrixEkler.GetNextSelectedRow();
                            if (rowEkler != -1)
                            {
                                string hedefYolEkler = ((SAPbouiCOM.EditText)oMatrixEkler.Columns.Item("Col_1").Cells.Item(rowEkler).Specific).Value;

                                silinecekDosya.Add(hedefYolEkler);

                                oMatrixEkler.DeleteRow(rowEkler);
                                //siranumarasiyazEk();
                            }
                            else if (rowEkler == -1)
                            {
                                Handler.SAPApplication.MessageBox("Satır seçimi yapılmadı.");
                                return false;
                            }

                            if (frmMusteriSikayetleri.Mode == BoFormMode.fm_OK_MODE)
                            {
                                frmMusteriSikayetleri.Mode = BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "1" && pVal.BeforeAction) //ekle guncelle btn
                    {
                        if (frmMusteriSikayetleri.Mode == BoFormMode.fm_VIEW_MODE)
                        {
                            //view modda tıklama yapılmasın
                            return false;
                        }

                        #region dosya ekleme
                        try
                        {
                            string kaynakYolEkler = "";
                            string hedefYolEkler = "";
                            string dosyaAdiEkler = "";

                            for (int i = 1; i <= oMatrixEkler.RowCount; i++)
                            {
                                kaynakYolEkler = ((SAPbouiCOM.EditText)oMatrixEkler.Columns.Item("Col_0").Cells.Item(i).Specific).Value;
                                hedefYolEkler = ((SAPbouiCOM.EditText)oMatrixEkler.Columns.Item("Col_1").Cells.Item(i).Specific).Value;
                                dosyaAdiEkler = ((SAPbouiCOM.EditText)oMatrixEkler.Columns.Item("Col_2").Cells.Item(i).Specific).Value;

                                File.Copy(kaynakYolEkler, hedefYolEkler, true); //kaynak dosya,hedef dosya,owwride=üzerine yaz

                                //if (File.Exists(ConstVariables.oCompanyObject.AttachMentPath + "\\" + dosyaAdi))
                                //{
                                //    Handler.SAPApplication.MessageBox("Belirtilen klasörde." + dosyaAdi + " isimli dosya mevcut..");
                                //}
                                //else
                                //{
                                //    File.Copy(dosyaYolu, ConstVariables.oCompanyObject.AttachMentPath + "\\" + dosyaAdi);
                                //    Handler.SAPApplication.MessageBox("Dosya kopyalama başarılı.");
                                //}
                            }

                            if (silinecekDosya.Count > 0)
                            {
                                foreach (var item in silinecekDosya)
                                {
                                    File.Delete(item); //sap ekler dosyasından siler
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                        #endregion

                    }
                    break;
                case BoEventTypes.et_DOUBLE_CLICK:
                    //if (pVal.ItemUID == "Item_6" && pVal.ColUID == "Col_0" && !pVal.BeforeAction)
                    //{

                    //    if (oEditUrunKodu.Value.ToString() == "")
                    //    {
                    //        Handler.SAPApplication.MessageBox("Lütfen Ürün Kodu seçimi yapınız.");
                    //        return false;
                    //    }

                    //    if (oEditUrunKodu.Value.ToString() != "")
                    //    {
                    //        AIFConn.AnlzGrsSec.LoadForms(oEditUrunKodu.Value.ToString(), "KimyasalAnaliz", analizPartilers);
                    //    }
                    //}
                    //else if (pVal.ItemUID == "Item_8" && pVal.ColUID == "Col_0" && !pVal.BeforeAction)
                    //{
                    //    if (oEditUrunKodu.Value.ToString() == "")
                    //    {
                    //        Handler.SAPApplication.MessageBox("Lütfen Ürün Kodu seçimi yapınız.");
                    //        return false;
                    //    }

                    //    if (oEditUrunKodu.Value.ToString() != "")
                    //    {
                    //        AIFConn.AnlzGrsSec.LoadForms(oEditUrunKodu.Value.ToString(), "DuyusalAnaliz", analizPartilers);
                    //    }
                    //}
                    //else if (pVal.ItemUID == "Item_9" && pVal.ColUID == "Col_14" && !pVal.BeforeAction)
                    //{
                    //    if (oEditUrunKodu.Value.ToString() == "")
                    //    {
                    //        Handler.SAPApplication.MessageBox("Lütfen Ürün Kodu seçimi yapınız.");
                    //        return false;
                    //    }
                    //    if (oEditUrunKodu.Value.ToString() != "")
                    //    {
                    //        AIFConn.AnlzGrsSec.LoadForms(oEditUrunKodu.Value.ToString(), "MikrobiyolojikAnaliz", analizPartilers);
                    //    }
                    //}
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
                    if (pVal.ItemUID == "Item_16" && pVal.ColUID == "Col_6" && pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                            oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                            SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                            oCFL = frmMusteriSikayetleri.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                            SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                            SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                            SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                            oCFL.SetConditions(oEmptyConts);
                            oCons = oCFL.GetConditions();

                            oCon = oCons.Add();
                            oCon.Alias = "validFor";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = "Y";

                            oCFL.SetConditions(oCons);
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_16" && pVal.ColUID == "Col_6" && !pVal.BeforeAction)
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
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            //oEditUrunKodu.Value = val;
                            ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific).Value = val;
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            val = oDataTable.GetValue("ItemName", 0).ToString();
                            ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_7").Cells.Item(pVal.Row).Specific).Value = val;
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            //oEditUrunAdi.Value = val;
                            //((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_7").Cells.Item(pVal.Row).Specific).Value = val;
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {

                            string sql = "select T1.Descr from CUFD as T0 INNER JOIN UFD1 as T1 ON T0.FieldID = T1.FieldID where T0.TableID = 'OITM' and T1.TableID = 'OITM' and T0.AliasID='ItemGroup2' and T1.FldValue = '" + oDataTable.GetValue("U_ItemGroup2", 0).ToString() + "'";

                            ConstVariables.oRecordset.DoQuery(sql);

                            ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_5").Cells.Item(pVal.Row).Specific).Value = ConstVariables.oRecordset.Fields.Item(0).Value.ToString();
                        }
                        catch (Exception)
                        {

                        }

                        #region iade birim
                        double sonuc = 0;

                        try
                        {
                            double iadebirim = oDataTable.GetValue("IWeight1", 0).ToString() == "" ? 0 : Convert.ToDouble(oDataTable.GetValue("IWeight1", 0));

                            sonuc = iadebirim / 1000;
                            sonuc = parseNumber.parservalues<double>(sonuc.ToString());
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            //oEditUrunAdi.Value = val;
                            ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_11").Cells.Item(pVal.Row).Specific).Value = sonuc.ToString(System.Globalization.CultureInfo.InvariantCulture); //iade birim
                        }
                        catch (Exception)
                        {
                        }
                        #endregion

                        oMatrixDetay.AutoResizeColumns();

                    }
                    else if (pVal.ItemUID == "Item_16" && pVal.ColUID == "Col_8" && pVal.BeforeAction)
                    {

                        bool cflCalled = false;

                        string kalemkodu = ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific).Value;

                        if (kalemkodu != "")
                        {
                            string sql = "";

                            //sql = "SELECT T0.\"DistNumber\",T0.\"ItemCode\",T0.\"Quantity\",T0.\"CreateDate\"  FROM OBTN T0 WHERE T0.\"ItemCode\" = '" + kalemkodu + "' and T0.\"CreateDate\" >= '" + DateTime.Now.AddDays(-7).ToString("yyyyMMdd") + "' ";

                            #region otat tan gelen sorgu 20231117
                            sql = "SELECT T0.\"DistNumber\",T0.\"ItemCode\",T0.\"ItemName\",T0.\"Quantity\",T0.\"CreateDate\" FROM OBTN T0 ";
                            sql += "WHERE T0.\"ItemCode\" = '" + kalemkodu + "' AND T0.ExpDate > '" + DateTime.Now.ToString("yyyyMMdd") + "' ORDER BY T0.\"DistNumber\"";
                            #endregion

                            CFLFilter(frmMusteriSikayetleri.UniqueID, pVal, "Item_16", "CFL_3", "DistNumber", "ItemCode", "", sql, ref cflCalled);
                        }

                        //try
                        //{
                        //    string kalemkodu = ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific).Value.ToString();
                        //    if (kalemkodu != "")
                        //    {

                        //        SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                        //        oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                        //        SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                        //        oCFL = frmMusteriSikayetleri.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                        //        SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                        //        SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                        //        SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                        //        oCFL.SetConditions(oEmptyConts);
                        //        oCons = oCFL.GetConditions();

                        //        oCon = oCons.Add();
                        //        oCon.Alias = "ItemCode";
                        //        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        //        oCon.CondVal = kalemkodu;

                        //        oCFL.SetConditions(oCons);
                        //    }

                        //}
                        //catch (Exception)
                        //{
                        //}


                    }
                    else if (pVal.ItemUID == "Item_16" && pVal.ColUID == "Col_8" && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                        if (oDataTable == null)
                        {
                            break;
                        }
                        string val = "";
                        try
                        {
                            val = oDataTable.GetValue("DistNumber", 0).ToString();
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            //oEditUrunKodu.Value = val;
                            ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_8").Cells.Item(pVal.Row).Specific).Value = val;
                        }
                        catch (Exception)
                        {
                        }

                        val = "";
                        try
                        {

                            val = oDataTable.GetValue("InDate", 0).ToString();
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            //oEditUrunAdi.Value = val;
                            if (val != "" && Convert.ToDateTime(val).Year.ToString() != "1900")
                            {
                                ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific).Value = Convert.ToDateTime(val).ToString("yyyyMMdd");
                            }
                            //else
                            //{
                            //    ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific).Value = val;
                            //}
                        }
                        catch (Exception)
                        {
                        }

                        val = "";

                        try
                        {
                            val = oDataTable.GetValue("ExpDate", 0).ToString();
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            //oEditUrunAdi.Value = val;
                            if (val != "" && Convert.ToDateTime(val).Year.ToString() != "1900")
                            {
                                ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_10").Cells.Item(pVal.Row).Specific).Value = Convert.ToDateTime(val).ToString("yyyyMMdd");
                            }
                            //else
                            //{
                            //    ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_10").Cells.Item(pVal.Row).Specific).Value = val;
                            //}

                        }
                        catch (Exception)
                        {
                        }

                        val = "";

                        try
                        {
                            val = oDataTable.GetValue("Quantity", 0).ToString();
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            //oEditUrunAdi.Value = val;
                            ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_17").Cells.Item(pVal.Row).Specific).Value = val;
                        }
                        catch (Exception)
                        {
                        }

                        #region iade kg
                        double sonuc = 0;
                        try
                        {
                            string sql = "Select \"IWeight1\" from OITM WHERE \"ItemCode\" = '" + ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_6").Cells.Item(pVal.Row).Specific).Value + "' ";
                            ConstVariables.oRecordset.DoQuery(sql);

                            if (ConstVariables.oRecordset.RecordCount > 0)
                            {

                                double iadekg = ConstVariables.oRecordset.Fields.Item(0).Value.ToString() == "" ? 0 : Convert.ToDouble(ConstVariables.oRecordset.Fields.Item(0).Value);

                                sonuc = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_17").Cells.Item(pVal.Row).Specific).Value) * (iadekg / 1000);
                                sonuc = parseNumber.parservalues<double>(sonuc.ToString());
                            }

                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            //oEditUrunAdi.Value = val;
                            hesaplamayapiliyorIadeKG = true;

                            ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_12").Cells.Item(pVal.Row).Specific).Value = sonuc.ToString(System.Globalization.CultureInfo.InvariantCulture); //iade kg
                            hesaplamayapiliyorIadeKG = false;
                        }
                        catch (Exception)
                        {
                        }
                        #endregion

                        oMatrixDetay.AutoResizeColumns();

                    }
                    else if (pVal.ItemUID == "Item_16" && pVal.ColUID == "Col_3" && pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                            oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                            SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                            oCFL = frmMusteriSikayetleri.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                            SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                            SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                            SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                            oCFL.SetConditions(oEmptyConts);
                            oCons = oCFL.GetConditions();

                            oCon = oCons.Add();
                            oCon.Alias = "validFor";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = "Y";

                            oCFL.SetConditions(oCons);
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_16" && pVal.ColUID == "Col_18" && pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                            oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                            SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                            oCFL = frmMusteriSikayetleri.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                            SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                            SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                            SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                            oCFL.SetConditions(oEmptyConts);
                            oCons = oCFL.GetConditions();

                            oCon = oCons.Add();
                            oCon.Alias = "validFor";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = "Y";

                            oCFL.SetConditions(oCons);
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_16" && pVal.ColUID == "Col_18" && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                        if (oDataTable == null)
                        {
                            break;
                        }
                        string val = "";
                        try
                        {
                            val = oDataTable.GetValue("CardCode", 0).ToString();
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            //oEditUrunKodu.Value = val;
                            ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_18").Cells.Item(pVal.Row).Specific).Value = val;
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            val = oDataTable.GetValue("CardName", 0).ToString();
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            //oEditUrunAdi.Value = val;
                            ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = val;
                        }
                        catch (Exception)
                        {
                        }

                        chooseFromListSeciliyor = true;

                        oMatrixDetay.AutoResizeColumns();

                    }
                    else if (pVal.ItemUID == "Item_16" && pVal.ColUID == "Col_14" && pVal.BeforeAction) //şikeyet nedeni
                    {
                        try
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                            oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                            SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                            oCFL = frmMusteriSikayetleri.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                            SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                            SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                            SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                            oCFL.SetConditions(oEmptyConts);
                            oCons = oCFL.GetConditions();

                            string sikayetGrubu = "";

                            if (((SAPbouiCOM.ComboBox)oMatrixDetay.Columns.Item("Col_13").Cells.Item(pVal.Row).Specific).Value != "")
                            {
                                sikayetGrubu = ((SAPbouiCOM.ComboBox)oMatrixDetay.Columns.Item("Col_13").Cells.Item(pVal.Row).Specific).Value.ToString();
                            }

                            oCon = oCons.Add();
                            oCon.Alias = "U_SkytGrupId";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = sikayetGrubu;

                            //oCon.Relationship = BoConditionRelationship.cr_AND;

                            //oCon = oCons.Add();
                            //oCon.Alias = "U_Aktif";
                            //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            //oCon.CondVal = "1";

                            oCFL.SetConditions(oCons);
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_16" && pVal.ColUID == "Col_14" && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                        if (oDataTable == null)
                        {
                            break;
                        }
                        string val = "";
                        try
                        {
                            val = oDataTable.GetValue("U_UygnszNeden", 0).ToString();
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            //oEditUrunKodu.Value = val;
                            ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_14").Cells.Item(pVal.Row).Specific).Value = val;
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
        List<string> silinecekler = new List<string>();
        string itemUID = "";
        public void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.MenuUID == "1282" && !pVal.BeforeAction)
            {
                try
                {
                    oMatrixDetay.AddRow();
                }
                catch (Exception)
                {
                }
            }

            try
            {
                if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && pVal.BeforeAction)
                {
                    if (itemUID == "Item_16")
                    {
                        int row = oMatrixDetay.GetNextSelectedRow();
                        if (row != -1)
                        {
                            //if (((SAPbouiCOM.EditText)oMatrixSonUrunKimyasal.Columns.Item("Col_0").Cells.Item(row).Specific).Value != "")
                            //{
                            //    silinecekler.Add(((SAPbouiCOM.EditText)oMatrixSonUrunKimyasal.Columns.Item("Col_0").Cells.Item(row).Specific).Value);
                            //}

                            oMatrixDetay.DeleteRow(row);

                            if (frmMusteriSikayetleri.Mode == BoFormMode.fm_OK_MODE)
                            {
                                frmMusteriSikayetleri.Mode = BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                    }

                    if (itemUID == "Item_11")
                    {
                        int row = oMatrixEkler.GetNextSelectedRow();
                        if (row != -1)
                        {
                            //if (((SAPbouiCOM.EditText)oMatrixSonUrunKimyasal.Columns.Item("Col_0").Cells.Item(row).Specific).Value != "")
                            //{
                            //    silinecekler.Add(((SAPbouiCOM.EditText)oMatrixSonUrunKimyasal.Columns.Item("Col_0").Cells.Item(row).Specific).Value);
                            //}

                            oMatrixEkler.DeleteRow(row);

                            if (frmMusteriSikayetleri.Mode == BoFormMode.fm_OK_MODE)
                            {
                                frmMusteriSikayetleri.Mode = BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                    }



                }
                else if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && pVal.BeforeAction)
                {
                    frmMusteriSikayetleri.DataSources.DBDataSources.Item("@AIF_MUSTERISIKAYET1").Clear();
                    oMatrixDetay.AddRow();

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

        public class DialogState
        {
            public DialogResult result;

            public FileDialog dialog;

            public void ThreadProcShowDialog()

            {
                result = dialog.ShowDialog();
            }
        }

        private static DialogResult STAShowDialog(FileDialog dialog)
        {
            DialogState state = new DialogState();

            state.dialog = dialog;

            System.Threading.Thread t = new System.Threading.Thread(state.ThreadProcShowDialog);

            t.SetApartmentState(System.Threading.ApartmentState.STA);

            t.Start();

            t.Join();

            return state.result;
        }
    }
}