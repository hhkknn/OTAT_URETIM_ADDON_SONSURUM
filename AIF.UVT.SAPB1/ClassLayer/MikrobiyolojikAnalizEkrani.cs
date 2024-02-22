using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class MikrobiyolojikAnalizEkrani
    {
        [ItemAtt(AIFConn.MikrobiyolojikUID)]
        public SAPbouiCOM.Form frmMikrobiyolojik;

        [ItemAtt("Item_2")]
        public SAPbouiCOM.EditText edtBasTarih;

        [ItemAtt("Item_3")]

        public SAPbouiCOM.EditText edtBtsTarih;

        [ItemAtt("Item_8")]
        public SAPbouiCOM.EditText edtAciklamalar;
        [ItemAtt("Item_9")]
        public SAPbouiCOM.Matrix oMatrixMikrobiyolojik;
        string secilmisUygunsuzlukNedeni = "";
        string matrixUID = "";
        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.MikrobiyolojikUID, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.MikrobiyolojikFrmXML));
            Functions.CreateUserOrSystemFormComponent<MikrobiyolojikAnalizEkrani>(AIFConn.Mkrbyljk);

            InitForms();
        }

        public void InitForms()
        {
            try
            {
                frmMikrobiyolojik.EnableMenu("1283", false);
                frmMikrobiyolojik.EnableMenu("1284", false);
                frmMikrobiyolojik.EnableMenu("1286", false);



                SAPbouiCOM.Column oCol = (SAPbouiCOM.Column)oMatrixMikrobiyolojik.Columns.Item("Col_4");

                string sql = "Select \"U_Id\", \"U_Bolum\" from \"@AIF_ANLZGIRISBOLUM\" ";
                ConstVariables.oRecordset.DoQuery(sql);

                if (ConstVariables.oRecordset.RecordCount > 0)
                {
                    while (!ConstVariables.oRecordset.EoF)
                    {
                        oCol.ValidValues.Add(ConstVariables.oRecordset.Fields.Item("U_Id").Value.ToString(), ConstVariables.oRecordset.Fields.Item("U_Bolum").Value.ToString());
                        ConstVariables.oRecordset.MoveNext();
                    }
                }


                oMatrixMikrobiyolojik.AddRow();
                oMatrixMikrobiyolojik.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Hata oluştu" + ex.Message);
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
                        string sonsatir1 = frmMikrobiyolojik.DataSources.DBDataSources.Item(1).GetValue("U_PartiNo", frmMikrobiyolojik.DataSources.DBDataSources.Item(1).Size - 1);


                        if (sonsatir1 == "")
                        {
                            frmMikrobiyolojik.DataSources.DBDataSources.Item(1).RemoveRecord(frmMikrobiyolojik.DataSources.DBDataSources.Item(1).Size - 1);
                        }


                    }
                    break;
                case BoEventTypes.et_FORM_DATA_UPDATE:
                    if (BusinessObjectInfo.BeforeAction)
                    {
                        string sonsatir1 = frmMikrobiyolojik.DataSources.DBDataSources.Item(1).GetValue("U_PartiNo", frmMikrobiyolojik.DataSources.DBDataSources.Item(1).Size - 1);


                        if (sonsatir1 == "")
                        {
                            frmMikrobiyolojik.DataSources.DBDataSources.Item(1).RemoveRecord(frmMikrobiyolojik.DataSources.DBDataSources.Item(1).Size - 1);
                        }


                    }
                    break;
                case BoEventTypes.et_FORM_DATA_DELETE:
                    break;
                case BoEventTypes.et_FORM_DATA_LOAD:
                    if (!BusinessObjectInfo.BeforeAction)
                    {
                        frmMikrobiyolojik.DataSources.DBDataSources.Item(1).Clear();

                        oMatrixMikrobiyolojik.AddRow();


                        oMatrixMikrobiyolojik.AutoResizeColumns();
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
                        frmMikrobiyolojik.DataSources.DBDataSources.Item(1).Clear();

                        oMatrixMikrobiyolojik.AddRow();


                        oMatrixMikrobiyolojik.AutoResizeColumns();
                    }

                    break;
                case BoEventTypes.et_KEY_DOWN:
                    break;
                case BoEventTypes.et_GOT_FOCUS:
                    break;
                case BoEventTypes.et_LOST_FOCUS:
                    if (pVal.ItemUID == "Item_9" && (pVal.ColUID == "Col_0" || pVal.ColUID == "Col_1") && !pVal.BeforeAction)
                    {
                        string deger = ((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).Value.ToString();
                        if (pVal.Row == oMatrixMikrobiyolojik.RowCount && !string.IsNullOrEmpty(deger))
                        {
                            frmMikrobiyolojik.DataSources.DBDataSources.Item(1).Clear();

                            oMatrixMikrobiyolojik.AddRow();
                            oMatrixMikrobiyolojik.AutoResizeColumns();
                        }

                    }
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    if (pVal.ItemUID == "Item_9" && pVal.ColUID == "Col_4" && pVal.BeforeAction)
                    {
                        secilmisUygunsuzlukNedeni = ((SAPbouiCOM.ComboBox)oMatrixMikrobiyolojik.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific).Value.ToString();
                    }
                    else if (pVal.ItemUID == "Item_9" && pVal.ColUID == "Col_4" && !pVal.BeforeAction)
                    {
                        string uygunsuzlukNedeni = ((SAPbouiCOM.ComboBox)oMatrixMikrobiyolojik.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific).Value.ToString();

                        if (secilmisUygunsuzlukNedeni != "" && uygunsuzlukNedeni != secilmisUygunsuzlukNedeni)
                        {
                            ((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item("Col_5").Cells.Item(pVal.Row).Specific).Value = "";
                        }
                    }
                    break;
                case BoEventTypes.et_CLICK:
                    break;
                case BoEventTypes.et_DOUBLE_CLICK:
                    if (pVal.ItemUID == "Item_9" && pVal.ColUID == "Col_0" && !pVal.BeforeAction)
                    {
                        matrixUID = pVal.ItemUID;

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
                    if (pVal.ItemUID == "Item_9" && pVal.ColUID == "Col_5" && pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                            oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                            SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                            oCFL = frmMikrobiyolojik.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                            SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                            SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                            SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                            oCFL.SetConditions(oEmptyConts);
                            oCons = oCFL.GetConditions();

                            string bolum = "";

                            if (((SAPbouiCOM.ComboBox)oMatrixMikrobiyolojik.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific).Value != "")
                            {
                                bolum = ((SAPbouiCOM.ComboBox)oMatrixMikrobiyolojik.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific).Value.ToString();
                            }

                            oCon = oCons.Add();
                            oCon.Alias = "U_BolumId";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = bolum;

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
                    else if (pVal.ItemUID == "Item_9" && pVal.ColUID == "Col_5" && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                        if (oDataTable == null)
                        {
                            break;
                        }
                        string val = "";
                        try
                        {
                            val = oDataTable.GetValue("U_Numune", 0).ToString();
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            //oEditUrunKodu.Value = val;
                            ((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item("Col_5").Cells.Item(pVal.Row).Specific).Value = val;
                        }
                        catch (Exception)
                        {
                        }

                    }
                    else if (pVal.ItemUID == "Item_9" && !pVal.BeforeAction) //son ürün kimyasal analiz - //duyusal analiz- //mikrobiyolojik analz
                    {


                        #region mikrobiyolojik analz matrisi analiz yapan
                        if (pVal.ItemUID == "Item_9" && pVal.ColUID == "Col_13")
                        {
                            try
                            {
                                SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                                string Val = "";
                                if (oDataTable == null)
                                {
                                    return false;
                                }
                                Val = oDataTable.GetValue("empID", 0).ToString();
                                try
                                {
                                    //oEditSorumlu.Value = Val;
                                    ((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item("Col_13").Cells.Item(pVal.Row).Specific).Value = Val;

                                }
                                catch (Exception)
                                {
                                }
                                Val = "";
                                try
                                {
                                    if (oDataTable.GetValue("middleName", 0).ToString() == "")
                                    {
                                        Val = oDataTable.GetValue("firstName", 0).ToString() + " " + oDataTable.GetValue("lastName", 0).ToString();

                                    }
                                    else
                                    {
                                        Val = oDataTable.GetValue("firstName", 0).ToString() + " " + oDataTable.GetValue("middleName", 0).ToString() + " " + oDataTable.GetValue("lastName", 0).ToString();
                                    }

                                    ((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item("Col_14").Cells.Item(pVal.Row).Specific).Value = Val;

                                }
                                catch (Exception ex)
                                {
                                }

                                //}
                            }
                            catch (Exception)
                            {
                            }
                            oMatrixMikrobiyolojik.AutoResizeColumns();
                        }
                        #endregion
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
                    //Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_AddRow");
                    return;
                }


                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_DeleteRow");
                    //Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_AddRow");
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


                }
                catch (Exception ex)
                {
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && pVal.BeforeAction)
                {
                     if (itemUID == "Item_9")
                    {
                        int row = oMatrixMikrobiyolojik.GetNextSelectedRow();
                        if (row != -1)
                        {
                            //if (((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item("Col_0").Cells.Item(row).Specific).Value != "")
                            //{
                            //    silinecekler.Add(((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item("Col_0").Cells.Item(row).Specific).Value);
                            //}

                            oMatrixMikrobiyolojik.DeleteRow(row);

                            if (frmMikrobiyolojik.Mode == BoFormMode.fm_OK_MODE)
                            {
                                frmMikrobiyolojik.Mode = BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

    }
}
