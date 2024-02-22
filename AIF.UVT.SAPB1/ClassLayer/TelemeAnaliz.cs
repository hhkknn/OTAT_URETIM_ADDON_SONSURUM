using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
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
    public class TelemeAnaliz
    {
        [ItemAtt(AIFConn.TelemeAnalizTakibiUID)]
        public SAPbouiCOM.Form frmTelemeAnaliz;

        [ItemAtt("Item_48")]
        public SAPbouiCOM.Matrix oMatrixSutOzellikleri;
        [ItemAtt("Item_49")]
        public SAPbouiCOM.Matrix oMatrixProsesOzellikleri;
        [ItemAtt("Item_2")]
        public SAPbouiCOM.Matrix oMatrixSonUrunOzellikleri;
        [ItemAtt("Item_13")]
        public SAPbouiCOM.EditText oEditUrunKodu;
        [ItemAtt("Item_7")]
        public SAPbouiCOM.EditText oEditTanimi;
        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.TelemeAnalizTakibiXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.TelemeAnalizTakibiXML));
            Functions.CreateUserOrSystemFormComponent<TelemeAnaliz>(AIFConn.TlmAnalz);

            InitForms();
        }
        public void InitForms()
        {
            try
            {
                frmTelemeAnaliz.EnableMenu("1283", false);
                frmTelemeAnaliz.EnableMenu("1284", false);
                frmTelemeAnaliz.EnableMenu("1286", false);
                oMatrixSutOzellikleri.AutoResizeColumns();
                oMatrixProsesOzellikleri.AutoResizeColumns();
                oMatrixSonUrunOzellikleri.AutoResizeColumns();

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
        private string matrixuid;

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
                    if (pVal.ColUID == "Col_1" && pVal.ItemUID == "Item_48" && pVal.BeforeAction)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                        oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                        SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                        oCFL = frmTelemeAnaliz.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                        SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                        SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                        SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                        oCFL.SetConditions(oEmptyConts);
                        oCons = oCFL.GetConditions();

                        oCon = oCons.Add();
                        oCon.Alias = "Active";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";

                        oCFL.SetConditions(oCons);
                    }
                    else if (pVal.ColUID == "Col_1" && pVal.ItemUID == "Item_48" && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                        string Val = "";
                        Val = oDataTable.GetValue("empID", 0).ToString();

                        try
                        {
                            ((SAPbouiCOM.EditText)oMatrixSutOzellikleri.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = Val;
                        }
                        catch (Exception)
                        {
                        }


                        Val = oDataTable.GetValue("firstName", 0).ToString() + " " + oDataTable.GetValue("lastName", 0).ToString();

                        try
                        {
                            ((SAPbouiCOM.EditText)oMatrixSutOzellikleri.Columns.Item("Col_11").Cells.Item(pVal.Row).Specific).Value = Val;
                        }
                        catch (Exception)
                        {
                        }
                    }
                    //else if (pVal.ItemUID == "Item_13" && pVal.BeforeAction)
                    //{
                    //    SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                    //    oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                    //    SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                    //    oCFL = frmTelemeAnaliz.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                    //    SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                    //    SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                    //    SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                    //    oCFL.SetConditions(oEmptyConts);
                    //    oCons = oCFL.GetConditions();

                    //    oCon = oCons.Add();
                    //    oCon.Alias = "Active";
                    //    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    //    oCon.CondVal = "Y";

                    //    oCFL.SetConditions(oCons);
                    //}
                    else if (pVal.ItemUID == "Item_13" && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                        string Val = "";
                        Val = oDataTable.GetValue("ItemCode", 0).ToString();

                        try
                        {
                            oEditUrunKodu.Value = Val;
                        }
                        catch (Exception)
                        {
                        }

                        Val = oDataTable.GetValue("ItemName", 0).ToString();

                        try
                        {
                            oEditTanimi.Value = Val;
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
                if (matrixuid == "Item_48")
                {
                    if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && !pVal.BeforeAction)
                    {
                        int row = oMatrixSutOzellikleri.GetNextSelectedRow();
                        if (row != -1)
                        {
                            oMatrixSutOzellikleri.DeleteRow(row);
                        }
                    }
                    if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && !pVal.BeforeAction)
                    {
                        frmTelemeAnaliz.DataSources.DBDataSources.Item(1).Clear();
                        oMatrixSutOzellikleri.AddRow();
                        oMatrixSutOzellikleri.AutoResizeColumns();
                    }
                }
                else if (matrixuid == "Item_49")
                {
                    if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && !pVal.BeforeAction)
                    {
                        int row = oMatrixProsesOzellikleri.GetNextSelectedRow();
                        if (row != -1)
                        {
                            oMatrixProsesOzellikleri.DeleteRow(row);
                        }
                    }
                    if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && !pVal.BeforeAction)
                    {
                        frmTelemeAnaliz.DataSources.DBDataSources.Item(2).Clear();
                        oMatrixProsesOzellikleri.AddRow();
                        oMatrixProsesOzellikleri.AutoResizeColumns();
                    }
                }
                else if (matrixuid == "Item_2")
                {
                    if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && !pVal.BeforeAction)
                    {
                        int row = oMatrixSonUrunOzellikleri.GetNextSelectedRow();
                        if (row != -1)
                        {
                            oMatrixSonUrunOzellikleri.DeleteRow(row);
                        }
                    }
                    if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && !pVal.BeforeAction)
                    {
                        frmTelemeAnaliz.DataSources.DBDataSources.Item(2).Clear();
                        oMatrixSonUrunOzellikleri.AddRow();
                        oMatrixSonUrunOzellikleri.AutoResizeColumns();
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

                matrixuid = eventInfo.ItemUID;
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
