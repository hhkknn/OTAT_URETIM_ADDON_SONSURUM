using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.Contracts;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class DolapTayin
    {
        [ItemAtt(AIFConn.DolapTayinUID)]
        public SAPbouiCOM.Form frmDolapTayin;

        [ItemAtt("Item_0")]
        public SAPbouiCOM.Matrix oMatrixDetay;

        [ItemAtt("Item_4")]
        public SAPbouiCOM.EditText oEditBelgeNo;

        [ItemAtt("Item_6")]
        public SAPbouiCOM.EditText oEditKalemKodu;

        public void LoadForms(string _kalemKodu)
        {
            kalemKodu = _kalemKodu;
            ConstVariables.oFnc.LoadSAPXML(AIFConn.frmDolapTayinFrmXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.frmDolapTayinFrmXML));
            Functions.CreateUserOrSystemFormComponent<DolapTayin>(AIFConn.DolapTayin);

            InitForms();
        }
        private SAPbouiCOM.DataTable oDataTable = null;
        string kalemKodu = "";
        bool eklemeGuncelleme = false;
        bool eklemeModu = false;
        bool bulmodu = false;
        string muhatapKodu = "";
        public void InitForms()
        {

            try
            {
                frmDolapTayin.Freeze(true);
                oDataTable = frmDolapTayin.DataSources.DataTables.Add("DATA");
                frmDolapTayin.EnableMenu("1283", false);
                frmDolapTayin.EnableMenu("1284", false);
                frmDolapTayin.EnableMenu("1286", false);

                oEditKalemKodu.Value = kalemKodu.ToString();
                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                //frmDolapTayin.DataSources.DBDataSources.Item(0).Clear();
                //frmDolapTayin.DataSources.DBDataSources.Item(1).Clear();
                string sql = "";

                sql = "Select T0.\"DocEntry\" FROM \"@AIF_DOLAPTAYIN\" as T0 INNER JOIN \"@AIF_DOLAPTAYIN1\" as T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where T0.\"U_KalemKodu\" = '" + kalemKodu + "'";

                ConstVariables.oRecordset.DoQuery(sql);

                if (ConstVariables.oRecordset.RecordCount > 0)
                {
                    try
                    {
                        //oEditBelgeNo.Value = ConstVariables.oRecordset.Fields.Item(0).Value.ToString();
                        //oEditKalemKodu.Value = ConstVariables.oRecordset.Fields.Item(1).Value.ToString();

                        //frmDolapTayin.DataSources.DBDataSources.Item(1).Query();


                        //oMatrixDetay.LoadFromDataSource();


                        if (kalemKodu != "")
                        {
                            frmDolapTayin.Mode = BoFormMode.fm_FIND_MODE;

                            bulmodu = true;
                            frmDolapTayin.Items.Item("Item_4").Enabled = true;
                            ((SAPbouiCOM.EditText)frmDolapTayin.Items.Item("Item_4").Specific).Value = ConstVariables.oRecordset.Fields.Item(0).Value.ToString();
                            frmDolapTayin.Items.Item("Item_1").Click();
                            frmDolapTayin.Items.Item("Item_4").Enabled = false;
                            frmDolapTayin.Items.Item("1").Click();
                            eklemeModu = false;

                            #region combobox lokasyon doldur
                            try
                            {
                                for (int i = 1; i <= oMatrixDetay.RowCount; i++)
                                {
                                    muhatapKodu = ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_0").Cells.Item(i).Specific).Value.ToString();


                                    if (muhatapKodu != "")
                                    {

                                        sql = "SELECT  T0.\"LineNum\",T0.\"Address\" FROM CRD1 T0 WHERE T0.\"CardCode\" = '" + muhatapKodu + "' and T0.\"AdresType\" = 'S' ";

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

                                }
                            }
                            catch (Exception ex)
                            {

                            }
                            #endregion
                        }
                    }
                    catch (Exception)
                    {
                    }

                    frmDolapTayin.Mode = BoFormMode.fm_OK_MODE;

                }
                //frmDolapTayin.DataSources.DBDataSources.Item(0).Clear();

                //oMatrixDetay.AddRow();
                oMatrixDetay.AutoResizeColumns();
            }
            catch (Exception ex)
            {
            }
            finally
            {
                frmDolapTayin.Freeze(false);
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

                    if (!BusinessObjectInfo.BeforeAction)
                    {
                        try
                        {
                            #region ekle tıklandıktan snra kullanıcıya bildirim göndermek için kaydedilecek docentry alınır
                            XmlDocument key = new XmlDocument();
                            key.LoadXml(BusinessObjectInfo.ObjectKey);
                            oEditBelgeNo.Value = key.SelectNodes("AIF_DOLAPTAYINParams").Item(0).SelectNodes("DocEntry").Item(0).InnerXml;
                            #endregion

                            eklemeGuncelleme = true;
                            eklemeModu = true;
                        }
                        catch (Exception)
                        {
                        }
                    }
                    //    string sonsatir = frmDolapTayin.DataSources.DBDataSources.Item(0).GetValue("DocEntry", frmDolapTayin.DataSources.DBDataSources.Item(0).Size - 1);

                    //    if (sonsatir == "")
                    //    {
                    //        frmDolapTayin.DataSources.DBDataSources.Item(0).RemoveRecord(frmDolapTayin.DataSources.DBDataSources.Item(0).Size - 1);
                    //    }
                    //    break;
                    //case BoEventTypes.et_FORM_DATA_UPDATE:
                    //    sonsatir = frmDolapTayin.DataSources.DBDataSources.Item(0).GetValue("DocEntry", frmDolapTayin.DataSources.DBDataSources.Item(0).Size - 1);

                    //    if (sonsatir == "")
                    //    {
                    //        frmDolapTayin.DataSources.DBDataSources.Item(0).RemoveRecord(frmDolapTayin.DataSources.DBDataSources.Item(0).Size - 1);
                    //    }
                    break;
                case BoEventTypes.et_FORM_DATA_UPDATE:

                    if (!BusinessObjectInfo.BeforeAction)
                    {
                        try
                        {
                            #region update tıklandıktan snra kullanıcıya bildirim göndermek için kaydedilecek docentry alınır
                            XmlDocument key = new XmlDocument();
                            key.LoadXml(BusinessObjectInfo.ObjectKey);
                            oEditBelgeNo.Value = key.SelectNodes("AIF_DOLAPTAYINParams").Item(0).SelectNodes("DocEntry").Item(0).InnerXml;
                            #endregion

                            eklemeGuncelleme = true;
                        }
                        catch (Exception)
                        {
                        }
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
        bool chooseFromListSeciliyor = false;
        public bool SAP_ItemEvent(string FormUID, ref ItemEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;

            switch (pVal.EventType)
            {
                case BoEventTypes.et_ALL_EVENTS:
                    break;
                case BoEventTypes.et_ITEM_PRESSED:
                    //if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    //{
                    //    if (silinecekler.Count > 0)
                    //    {
                    //        foreach (var item in silinecekler)
                    //        {
                    //            if (item != "")
                    //            {
                    //                ConstVariables.oRecordset.DoQuery("Delete from \"@AIF_DOLAPTAYIN1\" where \"DocEntry\" = '" + item + "'");
                    //            }
                    //        }
                    //    }

                    //    silinecekler = new List<string>();
                    //}
                    break;
                case BoEventTypes.et_KEY_DOWN:
                    break;
                case BoEventTypes.et_GOT_FOCUS:
                    break;
                case BoEventTypes.et_LOST_FOCUS:
                    if (pVal.ItemUID == "Item_0" && pVal.ColUID == "Col_0" && !pVal.BeforeAction && chooseFromListSeciliyor)
                    {
                        try
                        {
                            muhatapKodu = ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value.ToString();

                            if (muhatapKodu != "")
                            {
                                string sql = "";

                                sql = "SELECT  T0.\"LineNum\",T0.\"Address\" FROM CRD1 T0 WHERE T0.\"CardCode\" = '" + muhatapKodu + "' and T0.\"AdresType\" = 'S' ";

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
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    break;
                case BoEventTypes.et_CLICK:
                    if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    {
                        //formBaslangicDataGetir();
                        #region belge eklendikten sonra belgeyi getir
                        if (eklemeModu)
                        {
                            if (bulmodu)
                            {
                                bulmodu = false;
                                return false;
                            }

                            if (kalemKodu != "")
                            {
                                frmDolapTayin.Mode = BoFormMode.fm_FIND_MODE;

                                bulmodu = true;
                                frmDolapTayin.Items.Item("Item_4").Enabled = true;
                                ((SAPbouiCOM.EditText)frmDolapTayin.Items.Item("Item_10").Specific).Value = oEditBelgeNo.Value;
                                frmDolapTayin.Items.Item("Item_1").Click();
                                frmDolapTayin.Items.Item("Item_4").Enabled = false;
                                frmDolapTayin.Items.Item("1").Click();
                                eklemeModu = false;
                            }
                        }
                        #endregion
                    }
                    //if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    //{
                    //    ConstVariables.oRecordset.DoQuery("Select MAX(\"DocEntry\") from \"@AIF_BNKESLESTIRME\"");
                    //    int maxDocEntry = Convert.ToInt32(ConstVariables.oRecordset.Fields.Item(0).Value);
                    //    maxDocEntry++;

                    //    for (int i = 1; i <= oMatrix.RowCount; i++)
                    //    {
                    //        if (((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value == "")
                    //        {
                    //            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value = maxDocEntry.ToString();
                    //            maxDocEntry++;
                    //        }

                    //    }
                    //}

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
                    if (pVal.ColUID == "Col_0" && pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                            oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                            SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                            oCFL = frmDolapTayin.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

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
                    else if (pVal.ColUID == "Col_0" && !pVal.BeforeAction)
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
                            ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value = val;
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
                            ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = val;
                        }
                        catch (Exception)
                        {
                        }
                        oMatrixDetay.AutoResizeColumns();


                        chooseFromListSeciliyor = true;

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
        public void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && pVal.BeforeAction)
                {
                    int row = oMatrixDetay.GetNextSelectedRow();
                    if (row != -1)
                    {
                        //if (((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(row).Specific).Value != "")
                        //{
                        //    silinecekler.Add(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(row).Specific).Value);
                        //}
                        oMatrixDetay.DeleteRow(row);
                        if (frmDolapTayin.Mode == BoFormMode.fm_OK_MODE)
                        {
                            frmDolapTayin.Mode = BoFormMode.fm_UPDATE_MODE;
                        }


                    }
                }
                else if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && pVal.BeforeAction)
                {
                    frmDolapTayin.DataSources.DBDataSources.Item("@AIF_DOLAPTAYIN1").Clear();
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
                        SAPbouiCOM.Matrix item = (SAPbouiCOM.Matrix)oForm.Items.Item(eventInfo.ItemUID).Specific;
                    }
                    catch (Exception)
                    {
                        Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_AddRow");
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
