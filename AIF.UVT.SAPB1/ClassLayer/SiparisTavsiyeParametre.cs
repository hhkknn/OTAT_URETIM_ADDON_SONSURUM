using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
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
    public class SiparisTavsiyeParametre
    {
        [ItemAtt(AIFConn.SiparisTavsiyeUID)]
        public SAPbouiCOM.Form frmSiparisTavsiyeParametre;

        [ItemAtt("Item_0")]
        public SAPbouiCOM.Matrix oMatrix;

        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.SiparisTavsiyeFrmXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.SiparisTavsiyeFrmXML));
            Functions.CreateUserOrSystemFormComponent<SiparisTavsiyeParametre>(AIFConn.TAVPARAM);

            InitForms();
        }

        public void InitForms()
        {
            try
            {
                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                ConstVariables.oRecordset1 = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                frmSiparisTavsiyeParametre.EnableMenu("1283", false);
                frmSiparisTavsiyeParametre.EnableMenu("1284", false);
                frmSiparisTavsiyeParametre.EnableMenu("1286", false);

                oMatrix.AddRow();
                oMatrix.AutoResizeColumns();

                //SAPbouiCOM.Column oCol = (SAPbouiCOM.Column)oMatrix.Columns.Item("Col_1");
                //ConstVariables.oRecordset.DoQuery("select \"FldValue\",\"Descr\" from \"UFD1\"  where TableID=\"@AIF_TAVSIYEPARAM\"");
                //while (!ConstVariables.oRecordset.EoF)
                //{
                //    oCol.ValidValues.Add(ConstVariables.oRecordset.Fields.Item(0).Value.ToString(), ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                //    ConstVariables.oRecordset.MoveNext();
                //}


                frmSiparisTavsiyeParametre.DataSources.DBDataSources.Item("@AIF_TAVSIYEPARAM").Query();
                oMatrix.LoadFromDataSource();

                if (oMatrix.RowCount == 0)
                {
                    oMatrix.AddRow(5);
                }
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
                    if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    {
                        //kaydet();

                        for (int i = 1; i <= oMatrix.RowCount; i++)
                        {
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(i).Specific).Value = i.ToString();
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

        private void kaydet()
        {
            try
            {
                frmSiparisTavsiyeParametre.Freeze(true);
                string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                            select new
                            {
                                IstasyonKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                IstasyonAdi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                                RotaKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value,
                                RotaAdi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value,
                                AnalizKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,
                                AnalizAdi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                                DocEntry = x.ElementsBeforeSelf().Count() + 1,
                                AktifPasif = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value,
                            }).ToList();

                string xmlss = frmSiparisTavsiyeParametre.DataSources.DBDataSources.Item(0).GetAsXML();

                string header = @"<?xml version=""1.0"" encoding=""UTF-16"" ?><dbDataSources uid=""@AIF_ANALYSISPARAM""><rows>{0}</rows></dbDataSources>";

                string row = "<row><cells><cell><uid>U_StationCode</uid><value>{0}</value></cell><cell><uid>U_StationName</uid><value>{1}</value></cell><cell><uid>U_RotaCode</uid><value>{2}</value></cell><cell><uid>U_RotaName</uid><value>{3}</value></cell><cell><uid>U_AnalysisCode</uid><value>{4}</value></cell><cell><uid>U_AnalysisName</uid><value>{5}</value></cell><cell><uid>DocEntry</uid><value>{6}</value></cell><cell><uid>U_Active</uid><value>{7}</value></cell></cells></row>";

                string data = string.Join("", rows.Select(s => string.Format(row, s.IstasyonKodu, s.IstasyonAdi, s.RotaKodu, s.RotaAdi, s.AnalizKodu, s.AnalizAdi, s.DocEntry, s.AktifPasif)));

                string final = string.Format(header, data);
                frmSiparisTavsiyeParametre.DataSources.DBDataSources.Item("@AIF_ANALYSISPARAM").Clear();
                frmSiparisTavsiyeParametre.DataSources.DBDataSources.Item("@AIF_ANALYSISPARAM").LoadFromXML(final);

                xmlss = frmSiparisTavsiyeParametre.DataSources.DBDataSources.Item(0).GetAsXML();
            }
            catch (Exception ex)
            {
            }
            finally
            {
                frmSiparisTavsiyeParametre.Freeze(false);
            }
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
                    }
                }
                else if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && pVal.BeforeAction)
                {
                    frmSiparisTavsiyeParametre.DataSources.DBDataSources.Item(0).Clear();
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
                    //try
                    //{
                    //    oCreationPackage.UniqueID = "AIFRGHTCLK_DeleteRow";

                    //    oCreationPackage.String = "Satır Sil";

                    //    oCreationPackage.Enabled = true;

                    //    oMenuItem = Handler.SAPApplication.Menus.Item("1280");

                    //    oMenus = oMenuItem.SubMenus;

                    //    oMenus.AddEx(oCreationPackage);

                    //}
                    //catch
                    //{
                    //}

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
