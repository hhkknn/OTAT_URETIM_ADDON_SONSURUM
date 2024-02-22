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
    public class AnalizParametre
    {
        [ItemAtt(AIFConn.AnalizParametreGirisUID)]
        public SAPbouiCOM.Form frmAnalizParametre;

        [ItemAtt("Item_0")]
        public SAPbouiCOM.Matrix oMatrix;

        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.AnalizParametreGirisXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.AnalizParametreGirisXML));
            Functions.CreateUserOrSystemFormComponent<AnalizParametre>(AIFConn.AnalizParam);

            InitForms();
        }

        public void InitForms()
        {
            try
            {
                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                ConstVariables.oRecordset1 = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                frmAnalizParametre.EnableMenu("1283", false);
                frmAnalizParametre.EnableMenu("1284", false);
                frmAnalizParametre.EnableMenu("1286", false);
                oMatrix.AutoResizeColumns();

                SAPbouiCOM.Column oCol = (SAPbouiCOM.Column)oMatrix.Columns.Item("Col_2");
                ConstVariables.oRecordset.DoQuery("Select \"Code\", \"Desc\" from \"ORST\"");
                while (!ConstVariables.oRecordset.EoF)
                {
                    oCol.ValidValues.Add(ConstVariables.oRecordset.Fields.Item(0).Value.ToString(), ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                    ConstVariables.oRecordset.MoveNext();
                }

                string sql = "Select FldValue, Descr from (select FieldID from CUFD as T0 where T0.TableID = 'OWOR' and T0.AliasID = 'ISTASYON') as tbl INNER JOIN UFD1 as T1 ON tbl.FieldID = T1.FieldID and TableID = 'OWOR'";
                ConstVariables.oRecordset.DoQuery(sql);
                oCol = (SAPbouiCOM.Column)oMatrix.Columns.Item("Col_0");

                while (!ConstVariables.oRecordset.EoF)
                {
                    oCol.ValidValues.Add(ConstVariables.oRecordset.Fields.Item(0).Value.ToString(), ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                    ConstVariables.oRecordset.MoveNext();
                }

                //commit.
                IList<Tuple<string, string>> AnalizEkranlari = new List<Tuple<string, string>>();
                if (Program.mKod == "010OTATURVT")
                {
                    AnalizEkranlari.Add(Tuple.Create("1", "Ayran Analiz Ekranı"));
                    AnalizEkranlari.Add(Tuple.Create("2", "TazePeynirProsesTakip_1"));
                    AnalizEkranlari.Add(Tuple.Create("3", "TazePeynirProsesTakip_2"));
                    AnalizEkranlari.Add(Tuple.Create("4", "TelemeAnalizGiris"));
                    AnalizEkranlari.Add(Tuple.Create("5", "TelemeProsesTakip"));
                    AnalizEkranlari.Add(Tuple.Create("6", "TereyagProsesTakip_1"));
                    AnalizEkranlari.Add(Tuple.Create("7", "TereyagProsesTakip_2"));
                    AnalizEkranlari.Add(Tuple.Create("8", "Ayran Paketleme Analiz Ekranı"));
                    AnalizEkranlari.Add(Tuple.Create("9", "TostPeynirProsesTakip_1"));
                    AnalizEkranlari.Add(Tuple.Create("10", "TostPeynirProsesTakip_2"));
                    AnalizEkranlari.Add(Tuple.Create("11", "TelemeProsesTakip_2"));
                    AnalizEkranlari.Add(Tuple.Create("12", "YogurtProsesTakip_1"));
                    AnalizEkranlari.Add(Tuple.Create("13", "LorProsesTakip"));
                    AnalizEkranlari.Add(Tuple.Create("14", "PastorizasyonProsesTakip_1"));
                    AnalizEkranlari.Add(Tuple.Create("15", "TostPeynirProsesTakipSatir"));
                    AnalizEkranlari.Add(Tuple.Create("16", "BulkKulturAnaliz"));
                    AnalizEkranlari.Add(Tuple.Create("17", "TazePeynirSalamuraSuyu"));
                    AnalizEkranlari.Add(Tuple.Create("18", "YogurtMamulAnaliz")); 
                }
                else if (Program.mKod == "20URVT")
                {
                    AnalizEkranlari.Add(Tuple.Create("1", "Ayran Sütü Hazırlık"));
                    AnalizEkranlari.Add(Tuple.Create("2", "Ayran Pastörizasyon Hazırlık"));
                    AnalizEkranlari.Add(Tuple.Create("3", "Paketleme"));
                    AnalizEkranlari.Add(Tuple.Create("4", "Pastörizasyon"));
                    AnalizEkranlari.Add(Tuple.Create("5", "Kaşar Teleme"));
                    AnalizEkranlari.Add(Tuple.Create("6", "Kaşar Haşlama"));
                    AnalizEkranlari.Add(Tuple.Create("7", "Kaşar Paketleme"));
                    AnalizEkranlari.Add(Tuple.Create("8", "Konsantre Ürün UF Analiz"));
                    AnalizEkranlari.Add(Tuple.Create("9", "UF Proses Ür."));
                    AnalizEkranlari.Add(Tuple.Create("10", "Lor Pişirim"));
                    AnalizEkranlari.Add(Tuple.Create("11", "Tereyağı Ürün Analizi"));
                    AnalizEkranlari.Add(Tuple.Create("12", "Kaymak Ürün Analiz"));
                    AnalizEkranlari.Add(Tuple.Create("13", "Krem Peynir Üretim Analizi"));
                    AnalizEkranlari.Add(Tuple.Create("14", "Krema Ürün Analizi"));
                    AnalizEkranlari.Add(Tuple.Create("15", "Beyaz Peynir Paketleme"));
                    //AnalizEkranlari.Add(Tuple.Create("2", "TazePeynirProsesTakip_1"));
                    //AnalizEkranlari.Add(Tuple.Create("3", "TazePeynirProsesTakip_2"));
                    //AnalizEkranlari.Add(Tuple.Create("4", "TelemeAnalizGiris"));
                    //AnalizEkranlari.Add(Tuple.Create("5", "TelemeProsesTakip"));
                    //AnalizEkranlari.Add(Tuple.Create("6", "TereyagProsesTakip_1"));
                    //AnalizEkranlari.Add(Tuple.Create("7", "TereyagProsesTakip_2"));
                    //AnalizEkranlari.Add(Tuple.Create("8", "Ayran Paketleme Analiz Ekranı"));
                    //AnalizEkranlari.Add(Tuple.Create("9", "TostPeynirProsesTakip_1"));
                    //AnalizEkranlari.Add(Tuple.Create("10", "TostPeynirProsesTakip_2"));
                    //AnalizEkranlari.Add(Tuple.Create("11", "TelemeProsesTakip_2"));
                    //AnalizEkranlari.Add(Tuple.Create("12", "YogurtProsesTakip_1"));
                    //AnalizEkranlari.Add(Tuple.Create("13", "LorProsesTakip"));
                    //AnalizEkranlari.Add(Tuple.Create("14", "PastorizasyonProsesTakip_1"));
                    AnalizEkranlari.Add(Tuple.Create("16", "Beyaz Teleme"));
                    AnalizEkranlari.Add(Tuple.Create("17", "Homojen Yoğurt Sütü"));
                    AnalizEkranlari.Add(Tuple.Create("18", "Homojenize Yoğurt Dolum İnkübasyon"));
                    AnalizEkranlari.Add(Tuple.Create("19", "Stand.Kaymaklı Yoğurt Analiz"));
                    AnalizEkranlari.Add(Tuple.Create("20", "Krema Ürün Paketleme"));
                }


                oCol = (SAPbouiCOM.Column)oMatrix.Columns.Item("Col_4");
                foreach (var item in AnalizEkranlari)
                {
                    oCol.ValidValues.Add(item.Item1, item.Item2);
                }

                frmAnalizParametre.DataSources.DBDataSources.Item("@AIF_ANALYSISPARAM").Query();
                oMatrix.LoadFromDataSource();
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
                    try
                    {
                        if (pVal.ColUID == "Col_0" && !pVal.BeforeAction)
                        {
                            if (((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Selected != null)
                            {
                                val = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Selected.Description;
                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = val;
                            }
                        }
                        else if (pVal.ColUID == "Col_2" && !pVal.BeforeAction)
                        {
                            if (((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific).Selected != null)
                            {
                                val = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific).Selected.Description;
                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_3").Cells.Item(pVal.Row).Specific).Value = val;
                            }
                        }
                        else if (pVal.ColUID == "Col_4" && !pVal.BeforeAction)
                        {
                            if (((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific).Selected != null)
                            {
                                val = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific).Selected.Description;
                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_5").Cells.Item(pVal.Row).Specific).Value = val;
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                    break;

                case BoEventTypes.et_CLICK:
                    if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    {
                        //kaydet();

                        for (int i = 1; i <= oMatrix.RowCount; i++)
                        {
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_6").Cells.Item(i).Specific).Value = i.ToString();
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
                frmAnalizParametre.Freeze(true);
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

                string xmlss = frmAnalizParametre.DataSources.DBDataSources.Item(0).GetAsXML();

                string header = @"<?xml version=""1.0"" encoding=""UTF-16"" ?><dbDataSources uid=""@AIF_ANALYSISPARAM""><rows>{0}</rows></dbDataSources>";

                string row = "<row><cells><cell><uid>U_StationCode</uid><value>{0}</value></cell><cell><uid>U_StationName</uid><value>{1}</value></cell><cell><uid>U_RotaCode</uid><value>{2}</value></cell><cell><uid>U_RotaName</uid><value>{3}</value></cell><cell><uid>U_AnalysisCode</uid><value>{4}</value></cell><cell><uid>U_AnalysisName</uid><value>{5}</value></cell><cell><uid>DocEntry</uid><value>{6}</value></cell><cell><uid>U_Active</uid><value>{7}</value></cell></cells></row>";

                string data = string.Join("", rows.Select(s => string.Format(row, s.IstasyonKodu, s.IstasyonAdi, s.RotaKodu, s.RotaAdi, s.AnalizKodu, s.AnalizAdi, s.DocEntry, s.AktifPasif)));

                string final = string.Format(header, data);
                frmAnalizParametre.DataSources.DBDataSources.Item("@AIF_ANALYSISPARAM").Clear();
                frmAnalizParametre.DataSources.DBDataSources.Item("@AIF_ANALYSISPARAM").LoadFromXML(final);

                xmlss = frmAnalizParametre.DataSources.DBDataSources.Item(0).GetAsXML();
            }
            catch (Exception ex)
            {
            }
            finally
            {
                frmAnalizParametre.Freeze(false);
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
                    frmAnalizParametre.DataSources.DBDataSources.Item(0).Clear();
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