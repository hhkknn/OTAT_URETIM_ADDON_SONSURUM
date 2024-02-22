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
    public class KaliteDetayListesi
    {
        [ItemAtt(AIFConn.KaliteListeDetaylariuUID)]
        public SAPbouiCOM.Form frmKaliteDetaylari;

        [ItemAtt("Item_0")]
        public SAPbouiCOM.Matrix oMatrix;
        [ItemAtt("Item_3")]
        public SAPbouiCOM.EditText oEditDocEntry;

        //[ItemAtt("1")]
        //public SAPbouiCOM.Button btnAddOrUpdate;
        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.KaliteListeDetaylariXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.KaliteListeDetaylariXML));
            Functions.CreateUserOrSystemFormComponent<KaliteDetayListesi>(AIFConn.KltDetay);

            InitForms();
        }
        public void InitForms()
        {
            try
            {
                frmKaliteDetaylari.EnableMenu("1283", false);
                frmKaliteDetaylari.EnableMenu("1284", false);
                frmKaliteDetaylari.EnableMenu("1286", false);

                #region istasyon
                //ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
                string sql = "Select \"U_BolumKodu\",\"U_BolumAdi\" from \"@AIF_BOLUMLER\"";

                ConstVariables.oRecordset.DoQuery(sql);
                SAPbouiCOM.Column oColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item("Col_24");
                while (!ConstVariables.oRecordset.EoF)
                {
                    oColumn.ValidValues.Add(ConstVariables.oRecordset.Fields.Item(0).Value.ToString(), ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                    ConstVariables.oRecordset.MoveNext();
                }
                #endregion

                //frmKaliteDetaylari.DataSources.DBDataSources.Item(1).Query();
                //oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();


                //if (oMatrix.RowCount > 0)
                //{
                //    frmKaliteDetaylari.Mode = BoFormMode.fm_OK_MODE;
                //}
                //else
                //{
                //    oEditRaporTarihi.Value = DateTime.Now.ToString("yyyyMMdd");
                //    oMatrix.AddRow();
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
                    if (!BusinessObjectInfo.BeforeAction)
                    {
                        try
                        {
                            Handler.SAPApplication.StatusBar.SetText("Lütfen Bekleyiz... Form düzenlemeleri gerçekleştiriliyor...", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning);
                            frmKaliteDetaylari.Freeze(true);
                            for (int s = 1; s <= oMatrix.RowCount; s++)
                            {
                                var val = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("Col_2").Cells.Item(s).Specific).Value.Trim();
                                if (val == "2")
                                {
                                    oMatrix.CommonSetting.SetCellEditable(s, 7, true);
                                    for (int i = 8; i <= 27; i++)
                                    {
                                        //string colid = "Col_" + i.ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item(i).Cells.Item(s).Specific).Value = "";
                                        oMatrix.CommonSetting.SetCellEditable(s, i, false);
                                    }
                                }
                                else if (val == "1")
                                {
                                    for (int i = 7; i <= 27; i++)
                                    {
                                        //string colid = "Col_" + i.ToString();
                                        //((SAPbouiCOM.EditText)oMatrix.Columns.Item(colid).Cells.Item(pVal.Row).Specific).Value = "";
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item(i).Cells.Item(s).Specific).Value = "";
                                        oMatrix.CommonSetting.SetCellEditable(s, i, false);
                                    }
                                }
                                else if (val == "3")
                                {
                                    //((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_23").Cells.Item(pVal.Row).Specific).Value = "";

                                    oMatrix.CommonSetting.SetCellEditable(s, 7, false);
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_23").Cells.Item(s).Specific).Value = "";
                                    for (int i = 8; i <= 27; i++)
                                    {
                                        oMatrix.CommonSetting.SetCellEditable(s, i, true);
                                    }
                                }
                            }


                            Handler.SAPApplication.StatusBar.SetText("Form düzenlemeleri tamamlandı...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmKaliteDetaylari.Freeze(false);
                        }
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
                    if (pVal.ItemUID == "Item_0")
                    {
                        try
                        {
                            frmKaliteDetaylari.Freeze(true);
                            oMatrix.AutoResizeColumns();
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmKaliteDetaylari.Freeze(false);
                        }
                    }
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    try
                    {
                        frmKaliteDetaylari.Freeze(true);
                        if (pVal.ColUID == "Col_2" && !pVal.BeforeAction)
                        {
                            //oEditTiklamalik.Item.Click();
                            var val = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific).Value.Trim();
                            if (val == "2")
                            {
                                oMatrix.CommonSetting.SetCellEditable(pVal.Row, 7, true);
                                for (int i = 8; i <= 27; i++)
                                {
                                    //string colid = "Col_" + i.ToString();
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item(i).Cells.Item(pVal.Row).Specific).Value = "";
                                    oMatrix.CommonSetting.SetCellEditable(pVal.Row, i, false);
                                }
                            }
                            else if (val == "1")
                            {
                                for (int i = 7; i <= 27; i++)
                                {
                                    //string colid = "Col_" + i.ToString();
                                    //((SAPbouiCOM.EditText)oMatrix.Columns.Item(colid).Cells.Item(pVal.Row).Specific).Value = "";
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item(i).Cells.Item(pVal.Row).Specific).Value = "";
                                    oMatrix.CommonSetting.SetCellEditable(pVal.Row, i, false);
                                }
                            }
                            else if (val == "3")
                            {
                                //((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_23").Cells.Item(pVal.Row).Specific).Value = "";

                                oMatrix.CommonSetting.SetCellEditable(pVal.Row, 7, false);
                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_23").Cells.Item(pVal.Row).Specific).Value = "";
                                for (int i = 8; i <= 27; i++)
                                {
                                    oMatrix.CommonSetting.SetCellEditable(pVal.Row, i, true);
                                }
                            }
                        }
                        else if (pVal.ColUID == "Col_24" && !pVal.BeforeAction)
                        {
                            var cbx =(SAPbouiCOM.ComboBox)oMatrix.Columns.Item("Col_24").Cells.Item(pVal.Row).Specific; 

                            if (cbx.Selected != null)
                            {
                                string desc = cbx.Selected.Description;
                                string val = cbx.Selected.Value;

                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_25").Cells.Item(pVal.Row).Specific).Value = desc;
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                    finally
                    {

                        frmKaliteDetaylari.Freeze(false);
                    }
                    break;
                case BoEventTypes.et_CLICK:
                    if (pVal.ItemUID == "1" && pVal.BeforeAction)
                    {
                        try
                        {
                            if (frmKaliteDetaylari.Mode == BoFormMode.fm_ADD_MODE || frmKaliteDetaylari.Mode == BoFormMode.fm_UPDATE_MODE)
                            {
                                string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                                var data = (from x in XDocument.Parse(xml).Descendants("Row")
                                            select new
                                            {
                                                saatAraligi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_23" select new XElement(y.Element("Value"))).First().Value,
                                                tur = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value,
                                            }).ToList();


                                if (data.Where(x => x.tur == "Saat" && x.saatAraligi == "").Count() > 0)
                                {
                                    Handler.SAPApplication.MessageBox("Saat Aralığı girişi yapılmadan devam edilemez.");
                                    return true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    #region tek tablo kullanımında yapılmıştı
                    //if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    //{
                    //    try
                    //    {
                    //        if (frmKaliteDetaylari.Mode == BoFormMode.fm_ADD_MODE && frmKaliteDetaylari.Mode == BoFormMode.fm_UPDATE_MODE)
                    //        {
                    //            string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                    //            var data = (from x in XDocument.Parse(xml).Descendants("Row")
                    //                        select new
                    //                        {
                    //                            docentry = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_24" select new XElement(y.Element("Value"))).First().Value,
                    //                            sira = x.ElementsBeforeSelf().Count() + 1
                    //                        }).ToList();

                    //            ConstVariables.oRecordset.DoQuery("Select MAX(ISNULL(\"DocEntry\",0)) + 1 from \"@AIF_KALITEDETAY\" ");
                    //            int sira = Convert.ToInt32(ConstVariables.oRecordset.Fields.Item(0).Value);
                    //            foreach (var item in data.Where(x => x.docentry == ""))
                    //            {

                    //                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_24").Cells.Item(item.sira).Specific).Value = sira.ToString();

                    //                sira++;
                    //            }
                    //        }
                    //    }
                    //    catch (Exception)
                    //    {
                    //    }
                    //} 
                    #endregion

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
                        oMatrix.DeleteRow(row);
                        if (frmKaliteDetaylari.Mode == BoFormMode.fm_OK_MODE)
                        {
                            frmKaliteDetaylari.Mode = BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                }
                else if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && pVal.BeforeAction)
                {
                    frmKaliteDetaylari.DataSources.DBDataSources.Item(1).Clear();
                    oMatrix.AddRow();

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
