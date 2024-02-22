﻿using AIF.ObjectsDLL;
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
    public class AnalizGirisNumune
    {
        [ItemAtt(AIFConn.AnalizGirisNumuneUID)]
        public SAPbouiCOM.Form frmAnalizGirisNumune;

        [ItemAtt("Item_0")]
        public SAPbouiCOM.Matrix oMatrix;

        [ItemAtt("1")]
        public SAPbouiCOM.Button oBtnIptal;

        [ItemAtt("2")]
        public SAPbouiCOM.Button oBtnSec;

        private SAPbouiCOM.DataTable oDataTable = null;
        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.AnalizGirisNumuneXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.AnalizGirisNumuneXML));
            Functions.CreateUserOrSystemFormComponent<AnalizGirisNumune>(AIFConn.AnlzNumune);

            InitForms();
        }

        public void InitForms()
        {
            try
            {
                frmAnalizGirisNumune.Freeze(true);


                frmAnalizGirisNumune.EnableMenu("1283", false);
                frmAnalizGirisNumune.EnableMenu("1284", false);
                frmAnalizGirisNumune.EnableMenu("1286", false);

                oDataTable = frmAnalizGirisNumune.DataSources.DataTables.Add("DATA");
                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


                string sql = "Select * FROM \"@AIF_ANLZGIRISNUMUNE\" ";

                ConstVariables.oRecordset.DoQuery(sql);

                if (ConstVariables.oRecordset.RecordCount > 0)
                {
                    try
                    {
                        frmAnalizGirisNumune.DataSources.DBDataSources.Item(0).Query();
                        oMatrix.LoadFromDataSource();
                        if (oMatrix.RowCount > 0)
                        {
                            frmAnalizGirisNumune.Mode = BoFormMode.fm_OK_MODE;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
                //frmMusteriSikayetNedeni.DataSources.DBDataSources.Item(0).Clear();

                //oMatrix.AddRow();
                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Hata oluştu" + ex.Message);
            }

            finally
            {
                frmAnalizGirisNumune.Freeze(false);
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
                        string sonsatir1 = frmAnalizGirisNumune.DataSources.DBDataSources.Item(0).GetValue("DocEntry", frmAnalizGirisNumune.DataSources.DBDataSources.Item(0).Size - 1);

                        if (sonsatir1 == "")
                        {
                            frmAnalizGirisNumune.DataSources.DBDataSources.Item(0).RemoveRecord(frmAnalizGirisNumune.DataSources.DBDataSources.Item(0).Size - 1);
                        }
                    }
                    break;

                case BoEventTypes.et_FORM_DATA_UPDATE:
                    if (BusinessObjectInfo.BeforeAction)
                    {
                        string sonsatir1 = frmAnalizGirisNumune.DataSources.DBDataSources.Item(0).GetValue("DocEntry", frmAnalizGirisNumune.DataSources.DBDataSources.Item(0).Size - 1);

                        if (sonsatir1 == "")
                        {
                            frmAnalizGirisNumune.DataSources.DBDataSources.Item(0).RemoveRecord(frmAnalizGirisNumune.DataSources.DBDataSources.Item(0).Size - 1);
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
                        if (silinecekler.Count > 0)
                        {
                            foreach (var item in silinecekler)
                            {
                                if (item != "")
                                {
                                    ConstVariables.oRecordset.DoQuery("Delete from \"@AIF_ANLZGIRISNUMUNE\" where \"DocEntry\" = '" + item + "'");
                                }
                            }
                        }

                        silinecekler = new List<string>();
                    }
                    //if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    //{
                    //    try
                    //    {
                    //        frmAnalizGirisNumune.Freeze(true);

                    //        //frmAnalizGirisBolumler.DataSources.DBDataSources.Item(0).Clear();

                    //        //oMatrix.AddRow();

                    //        //oMatrix.AutoResizeColumns();

                    //        frmAnalizGirisNumune.DataSources.DBDataSources.Item(0).Query();
                    //        oMatrix.LoadFromDataSource();
                    //        oMatrix.AutoResizeColumns();

                    //        //frmAnalizGirisNumune.DataSources.DBDataSources.Item("@AIF_ANLZGIRISNUMUNE").Clear();

                    //        //oMatrix.AddRow();
                    //        //frmAnalizGirisNumune.Mode = BoFormMode.fm_UPDATE_MODE;
                    //    }
                    //    catch (Exception)
                    //    {
                    //    }

                    //    finally
                    //    {
                    //        frmAnalizGirisNumune.Freeze(false);
                    //    }

                    //}
                    break;

                case BoEventTypes.et_KEY_DOWN:
                    break;

                case BoEventTypes.et_GOT_FOCUS:
                    break;

                case BoEventTypes.et_LOST_FOCUS:
                    if (pVal.ItemUID == "Item_0" && pVal.ColUID == "Col_1" && !pVal.BeforeAction)
                    {
                        //int row = oMatrix.GetNextSelectedRow();
                        //if (row != -1)
                        //{
                        try
                        {
                            //patlıyor
                            if (((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value != "")
                            {
                                //if (((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_0").Cells.Item(oMatrixDetay.RowCount - 1).Specific).Value == "")
                                //{
                                frmAnalizGirisNumune.DataSources.DBDataSources.Item("@AIF_ANLZGIRISNUMUNE").Clear();
                                oMatrix.AddRow();
                                //}
                            }
                        }
                        catch (Exception)
                        {

                        }
                        //}
                    }
                    break;

                case BoEventTypes.et_COMBO_SELECT:
                    break;

                case BoEventTypes.et_CLICK:
                    if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    {
                        ConstVariables.oRecordset.DoQuery("Select MAX(\"DocEntry\") from \"@AIF_ANLZGIRISNUMUNE\"");
                        int maxDocEntry = Convert.ToInt32(ConstVariables.oRecordset.Fields.Item(0).Value);
                        maxDocEntry++;

                        for (int i = 1; i <= oMatrix.RowCount; i++)
                        {
                            if (((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific).Value != "")
                            {
                                if (((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value == "")
                                {
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value = maxDocEntry.ToString();
                                    maxDocEntry++;
                                }
                            }
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
        private List<string> silinecekler = new List<string>();

        public void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && pVal.BeforeAction)
                {
                    int row = oMatrix.GetNextSelectedRow();
                    if (row != -1)
                    {
                        if (((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(row).Specific).Value != "")
                        {
                            silinecekler.Add(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(row).Specific).Value);
                        }
                        oMatrix.DeleteRow(row);
                        if (frmAnalizGirisNumune.Mode == BoFormMode.fm_OK_MODE)
                        {
                            frmAnalizGirisNumune.Mode = BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                }
                else if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && pVal.BeforeAction)
                {
                    frmAnalizGirisNumune.DataSources.DBDataSources.Item(0).Clear();
                    oMatrix.AddRow();
                    frmAnalizGirisNumune.Mode = BoFormMode.fm_UPDATE_MODE;

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