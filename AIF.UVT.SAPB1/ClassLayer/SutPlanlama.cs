using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using AIF.UVT.SAPB1.HelperClass;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing.Design;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.XPath;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class SutPlanlama
    {
        [ItemAtt(AIFConn.SutPlanlamaUID)]
        public SAPbouiCOM.Form frmSutPlanlama;

        [ItemAtt("Item_2")]
        public SAPbouiCOM.Matrix oMatrixSatici;
        [ItemAtt("Item_3")]
        public SAPbouiCOM.Matrix oMatrixSaticiTum;
        [ItemAtt("Item_1")]
        public SAPbouiCOM.EditText oEdtTarih;
        [ItemAtt("Item_7")]
        public SAPbouiCOM.EditText oEdtDocEntry;
        [ItemAtt("Item_4")]
        public SAPbouiCOM.EditText oEdtTiklamalik;
        [ItemAtt("Item_8")]
        public SAPbouiCOM.EditText oEdtToplamPlanlananSutMiktari;
        [ItemAtt("1")]
        public SAPbouiCOM.Button btnAddOrUpdate;
        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.SutPlanlamaXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.SutPlanlamaXML));
            Functions.CreateUserOrSystemFormComponent<SutPlanlama>(AIFConn.SutPlan);

            InitForms();
        }
        public void InitForms()
        {
            try
            {
                frmSutPlanlama.EnableMenu("1283", false);
                frmSutPlanlama.EnableMenu("1284", false);
                frmSutPlanlama.EnableMenu("1286", false);
                frmSutPlanlama.EnableMenu("1287", true);

                oMatrixSaticiTum.Columns.Item("Col_2").ColumnSetting.SumType = BoColumnSumType.bst_Auto;
                oMatrixSatici.Columns.Item("Col_2").ColumnSetting.SumType = BoColumnSumType.bst_Auto;

                oMatrixSaticiTum.AutoResizeColumns();
                oMatrixSatici.AutoResizeColumns();

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
                    if (BusinessObjectInfo.BeforeAction)
                    {
                        if (oEdtTarih.Value.ToString() == "")
                        {
                            Handler.SAPApplication.MessageBox("Tarih girilmeden ekleme yapılamaz.");
                            BubbleEvent = false;
                        }
                    }
                    break;
                case BoEventTypes.et_FORM_DATA_UPDATE:
                    if (BusinessObjectInfo.BeforeAction)
                    {
                        if (oEdtTarih.Value.ToString() == "")
                        {
                            Handler.SAPApplication.MessageBox("Tarih girilmeden güncelleme yapılamaz.");
                            BubbleEvent = false;
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

        string val = "";
        bool gotfocus = false;
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
                    if (pVal.ItemUID == "Item_1")
                    {
                        //if (frmSutPlanlama != null)
                        //{
                        //    if (frmSutPlanlama.Mode != BoFormMode.fm_ADD_MODE)
                        //    {
                        //        gotfocus = true;
                        //        oEdtTiklamalik.Item.Click();
                        //        gotfocus = false;
                        //    }
                        //}
                    }
                    break;
                case BoEventTypes.et_LOST_FOCUS:
                    if (pVal.ItemUID == "Item_1" && !gotfocus)
                    {
                        if (frmSutPlanlama != null)
                        {
                            if (frmSutPlanlama.Mode != BoFormMode.fm_FIND_MODE)
                            {
                                if (oEdtTarih != null)
                                {
                                    ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    ConstVariables.oRecordset.DoQuery("Select * from \"@AIF_SUTPLANLAMA\" where \"U_Tarih\" = '" + oEdtTarih.Value.ToString() + "'");

                                    if (ConstVariables.oRecordset.RecordCount > 0)
                                    {
                                        frmSutPlanlama.Mode = BoFormMode.fm_FIND_MODE;
                                        oEdtDocEntry.Value = ConstVariables.oRecordset.Fields.Item("DocEntry").Value.ToString();
                                        btnAddOrUpdate.Item.Click();
                                    }
                                }
                            }
                        }
                    }
                    else if (pVal.ColUID == "Col_2" && !pVal.BeforeAction)
                    {
                        var matrix1xml = oMatrixSaticiTum.SerializeAsXML(BoMatrixXmlSelect.mxs_All);


                        var rows = (from x in XDocument.Parse(matrix1xml).Descendants("Row")
                                    select new
                                    {
                                        Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value)
                                    }).ToList();

                        var matrix2xml = oMatrixSatici.SerializeAsXML(BoMatrixXmlSelect.mxs_All);


                        var rows2 = (from x in XDocument.Parse(matrix2xml).Descendants("Row")
                                     select new
                                     {
                                         Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value)
                                     }).ToList();



                        var value1 = rows.Sum(x => x.Miktar);
                        var value2 = rows2.Sum(x => x.Miktar);

                        var result = value1 + value2;

                        frmSutPlanlama.DataSources.DBDataSources.Item(0).SetValue("U_ToplamSutMiktari", 0, result.ToString());
                        //oEdtToplamPlanlananSutMiktari.Value = result.ToString();
                    }
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
                    if ((pVal.ItemUID == "Item_2" || pVal.ItemUID == "Item_3") && pVal.BeforeAction)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                        oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                        SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                        oCFL = frmSutPlanlama.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                        SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                        SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                        SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                        oCFL.SetConditions(oEmptyConts);
                        oCons = oCFL.GetConditions();

                        oCon = oCons.Add();
                        oCon.Alias = "CardType";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "S";

                        oCon.Relationship = BoConditionRelationship.cr_AND;


                        oCon = oCons.Add();
                        oCon.Alias = "QryGroup1";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";

                        oCFL.SetConditions(oCons);
                    }
                    else if ((pVal.ItemUID == "Item_2" || pVal.ItemUID == "Item_3") && !pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.Matrix oMatrix = null;

                            if (pVal.ItemUID == "Item_2")
                            {
                                oMatrix = oMatrixSatici;
                            }
                            else if (true)
                            {
                                oMatrix = oMatrixSaticiTum;
                            }

                            SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                            string Val = "";
                            Val = oDataTable.GetValue("CardCode", 0).ToString();
                            try
                            {
                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value = Val;
                            }
                            catch (Exception)
                            {
                            }

                            Val = oDataTable.GetValue("CardName", 0).ToString();
                            try
                            {
                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = Val;
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
            if (matrixuid == "Item_2")
            {
                if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && pVal.BeforeAction)
                {
                    int row = oMatrixSatici.GetNextSelectedRow();
                    if (row != -1)
                    {
                        oMatrixSatici.DeleteRow(row);
                    }
                }
                else if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && pVal.BeforeAction)
                {
                    frmSutPlanlama.DataSources.DBDataSources.Item(1).Clear();
                    oMatrixSatici.AddRow();
                }
            }
            else if (matrixuid == "Item_3")
            {
                if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && pVal.BeforeAction)
                {
                    int row = oMatrixSaticiTum.GetNextSelectedRow();
                    if (row != -1)
                    {
                        oMatrixSaticiTum.DeleteRow(row);
                    }
                }
                else if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && pVal.BeforeAction)
                {
                    frmSutPlanlama.DataSources.DBDataSources.Item(2).Clear();
                    oMatrixSaticiTum.AddRow();
                }
            }
            else if (pVal.MenuUID == "1287" && pVal.BeforeAction)
            {
                BubbleEvent = false;
                //frmSutPlanlama.Mode = BoFormMode.fm_ADD_MODE;
                //oEdtTarih.Value = "";
                frmSutPlanlama.DataSources.DBDataSources.Item(0).SetValue("U_Tarih", 0, "");
                frmSutPlanlama.DataSources.DBDataSources.Item(0).SetValue("DocEntry", 0, "");
                frmSutPlanlama.DataSources.DBDataSources.Item(0).SetValue("DocNum", 0, "");
                oEdtTarih.Item.Click();
            }

        }

        string matrixuid = "";

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
