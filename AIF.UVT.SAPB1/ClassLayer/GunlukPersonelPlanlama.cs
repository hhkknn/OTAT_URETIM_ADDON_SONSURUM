using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class GunlukPersonelPlanlama
    {
        [ItemAtt(AIFConn.GunlukPersonelPlanUID)]
        public SAPbouiCOM.Form frmGunlukPersonelPlanlama;

        [ItemAtt("Item_2")]
        public SAPbouiCOM.Matrix oMatrix;

        [ItemAtt("1")]
        public SAPbouiCOM.Button btnAddOrUpdate;

        [ItemAtt("Item_8")]
        public SAPbouiCOM.EditText oEditTarih;

        [ItemAtt("Item_1")]
        public SAPbouiCOM.EditText oEditTariheKopyala;

        [ItemAtt("Item_10")]
        public SAPbouiCOM.EditText oEditDocEntry;

        private string header = @"<?xml version=""1.0"" encoding=""UTF-16"" ?><dbDataSources uid=""@AIF_GUNPERSPLAN1""><rows>{0}</rows></dbDataSources>";

        public static string operasyonplaniTarihi = "";

        private string row = "<row><cells><cell><uid>U_PersonelAdi</uid><value>{0}</value></cell><cell><uid>U_PersonelNo</uid><value>{1}</value></cell><cell><uid>U_IST001</uid><value>{2}</value></cell><cell><uid>U_IST002</uid><value>{3}</value></cell><cell><uid>U_IST003</uid><value>{4}</value></cell><cell><uid>U_IST004</uid><value>{5}</value></cell><cell><uid>U_IST005</uid><value>{6}</value></cell><cell><uid>U_IST006</uid><value>{7}</value></cell><cell><uid>U_IST007</uid><value>{8}</value></cell><cell><uid>U_IST008</uid><value>{9}</value></cell><cell><uid>U_Kolileme</uid><value>{10}</value></cell><cell><uid>U_SutAlim</uid><value>{11}</value></cell><cell><uid>U_BulasikHane</uid><value>{12}</value></cell><cell><uid>U_SevkiyatDepo</uid><value>{13}</value></cell><cell><uid>U_AmbalajDepo</uid><value>{14}</value></cell><cell><uid>U_Tarihleme</uid><value>{15}</value></cell><cell><uid>U_GenelTemizlik</uid><value>{16}</value></cell><cell><uid>U_Mutfak</uid><value>{17}</value></cell><cell><uid>U_KazanDairesi</uid><value>{18}</value></cell><cell><uid>U_BakimOnarim</uid><value>{19}</value></cell></cells></row>";

        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.GunlukPersonelPlanXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.GunlukPersonelPlanXML));
            Functions.CreateUserOrSystemFormComponent<GunlukPersonelPlanlama>(AIFConn.GunPersPl);
            InitForms();
        }

        public void InitForms()
        {
            try
            {
                frmGunlukPersonelPlanlama.Freeze(true);
                frmGunlukPersonelPlanlama.EnableMenu("1283", false);
                frmGunlukPersonelPlanlama.EnableMenu("1284", false);
                frmGunlukPersonelPlanlama.EnableMenu("1286", false);
                oMatrix.AutoResizeColumns();
                CalisanListele();
                oEditTariheKopyala.Item.AffectsFormMode = false;
            }
            catch (Exception ex)
            {
            }
            finally
            {
                frmGunlukPersonelPlanlama.Freeze(false);
            }
        }

        private void CalisanListele()
        {
            try
            {
                frmGunlukPersonelPlanlama.Freeze(true);
                if (operasyonplaniTarihi == "")
                {
                    ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    ConstVariables.oRecordset.DoQuery("Select (\"firstName\" + ' ' + \"lastName\") as \"AdSoyad\",\"empID\"  from \"OHEM\" where \"Active\" = 'Y'");

                    string xmll = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                    XDocument xDoc = XDocument.Parse(xmll);
                    XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                    var rows = (from t in xDoc.Descendants(ns + "Row")
                                select new
                                {
                                    AdSoyad = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "AdSoyad" select new XElement(y.Element(ns + "Value"))).First().Value,
                                    Id = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "empID" select new XElement(y.Element(ns + "Value"))).First().Value,
                                }).ToList();

                    string data = string.Join("", rows.Select(s => string.Format(row, s.AdSoyad, s.Id, "N", "N", "N", "N", "N", "N", "N", "N", "N", "N", "N", "N", "N", "N", "N", "N", "N", "N")));

                    frmGunlukPersonelPlanlama.DataSources.DBDataSources.Item("@AIF_GUNPERSPLAN1").LoadFromXML(string.Format(header, data));

                    oMatrix.AutoResizeColumns();
                }
                else
                {
                    ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    ConstVariables.oRecordset.DoQuery("Select * from \"@AIF_GUNPERSPLAN\" where \"U_Tarih\"= '" + operasyonplaniTarihi + "'");
                    string docEntry = ConstVariables.oRecordset.Fields.Item("DocEntry").Value.ToString();

                    if (docEntry != "0")
                    {
                        frmGunlukPersonelPlanlama.Mode = BoFormMode.fm_FIND_MODE;
                        oEditDocEntry.Value = docEntry;
                        btnAddOrUpdate.Item.Click();
                    }
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                frmGunlukPersonelPlanlama.Freeze(false);
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
                        renklendir();
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

        private void TarihKontrol(string tarih)
        {
            string val = tarih;
            ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            ConstVariables.oRecordset.DoQuery("Select * from \"@AIF_GUNPERSPLAN\" where \"U_Tarih\"= '" + val + "'");

            if (ConstVariables.oRecordset.RecordCount > 0 && !question)
            {
                question = true;
                int retval = Handler.SAPApplication.MessageBox("İlgili tarih için daha önce giriş yapılmıştır.", 1, "İlgili Kaydı Getir  ", "İlgili Kaydı Sil", "İptal");

                if (retval == 1)
                {
                    string docEntry = ConstVariables.oRecordset.Fields.Item("DocEntry").Value.ToString();
                    frmGunlukPersonelPlanlama.Mode = BoFormMode.fm_FIND_MODE;
                    oEditDocEntry.Value = docEntry;
                    btnAddOrUpdate.Item.Click();
                }
                else if (retval == 2)
                {
                    string docEntry = ConstVariables.oRecordset.Fields.Item("DocEntry").Value.ToString();
                    SAPbobsCOM.GeneralService oGeneralService = null;

                    SAPbobsCOM.GeneralData oGeneralData = null;

                    SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                    SAPbobsCOM.CompanyService oCompService = null;

                    oCompService = ConstVariables.oCompanyObject.GetCompanyService();

                    oGeneralService = oCompService.GetGeneralService("AIF_GUNPERSPLAN");

                    oGeneralParams = ((SAPbobsCOM.GeneralDataParams)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)));

                    oGeneralParams.SetProperty("DocEntry", docEntry);

                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                    oGeneralService.Delete(oGeneralParams);
                }
            }
            else
                question = false;
        }

        private bool question = false;

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
                        if (oMatrix.RowCount == 0)
                        {
                            CalisanListele();
                            renklendir(true);
                        }
                        else
                        {
                            renklendir();
                        }
                    }
                    break;

                case BoEventTypes.et_KEY_DOWN:
                    break;

                case BoEventTypes.et_GOT_FOCUS:
                    break;

                case BoEventTypes.et_LOST_FOCUS:
                    if (pVal.ItemUID == "Item_8" && !pVal.BeforeAction)
                    {
                        try
                        {
                            if (frmGunlukPersonelPlanlama != null)
                            {
                                if (frmGunlukPersonelPlanlama.Mode != BoFormMode.fm_FIND_MODE && frmGunlukPersonelPlanlama.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    TarihKontrol(oEditTarih.Value.ToString());
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_1" && !pVal.BeforeAction)
                    {
                        try
                        {
                            if (frmGunlukPersonelPlanlama.Mode != BoFormMode.fm_FIND_MODE)
                            {
                                TarihKontrol(oEditTariheKopyala.Value.ToString());
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    break;

                case BoEventTypes.et_COMBO_SELECT:
                    break;

                case BoEventTypes.et_CLICK:
                    if (pVal.ItemUID == "Item_0" && !pVal.BeforeAction)
                    {
                        try
                        {
                            frmGunlukPersonelPlanlama.Freeze(true);
                            string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                            var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                        select new
                                        {
                                            AdSoyad = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "#" select new XElement(y.Element("Value"))).First().Value,
                                            Id = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_18" select new XElement(y.Element("Value"))).First().Value,
                                        }).ToList();

                            ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            ConstVariables.oRecordset.DoQuery("Select (\"firstName\" + ' ' + \"lastName\") as \"AdSoyad\",\"empID\"  from \"OHEM\" where \"Active\" = 'Y'");

                            string xmll = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                            XDocument xDoc = XDocument.Parse(xmll);
                            XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                            var rowsOhem = (from t in xDoc.Descendants(ns + "Row")
                                            select new
                                            {
                                                AdSoyad = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "AdSoyad" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                Id = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "empID" select new XElement(y.Element(ns + "Value"))).First().Value,
                                            }).ToList();

                            var DeleteList = rows.Except(rowsOhem).ToList();
                            var addList = rowsOhem.Except(rows).ToList();

                            var rows2 = (from x in XDocument.Parse(xml).Descendants("Row")
                                         select new
                                         {
                                             PersonelAdi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "#" select new XElement(y.Element("Value"))).First().Value,
                                             PersonelNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_18" select new XElement(y.Element("Value"))).First().Value,
                                             Ayran = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                             Yogurt = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                                             Teleme = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value,
                                             Kasar = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value,
                                             Kaymak = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,
                                             Lor = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                                             Tereyag = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value,
                                             Dondurma = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value,
                                             Kolileme = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value,
                                             SutAlim = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_9" select new XElement(y.Element("Value"))).First().Value,
                                             BulasikHane = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_10" select new XElement(y.Element("Value"))).First().Value,
                                             SevkiyatDepo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_11" select new XElement(y.Element("Value"))).First().Value,
                                             AmbalajDepo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_12" select new XElement(y.Element("Value"))).First().Value,
                                             Tarihleme = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_13" select new XElement(y.Element("Value"))).First().Value,
                                             GenelTemizlik = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_14" select new XElement(y.Element("Value"))).First().Value,
                                             Mutfak = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_15" select new XElement(y.Element("Value"))).First().Value,
                                             KazanDairesi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_16" select new XElement(y.Element("Value"))).First().Value,
                                             BakimOnarim = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_17" select new XElement(y.Element("Value"))).First().Value,
                                         }).ToList();

                            foreach (var item in DeleteList)
                            {
                                rows2.RemoveAll(x => x.PersonelNo == item.Id);
                            }

                            foreach (var item in addList)
                            {
                                rows2.Insert(rows2.Count, new { PersonelAdi = item.AdSoyad, PersonelNo = item.Id, Ayran = "N", Yogurt = "N", Teleme = "N", Kasar = "N", Kaymak = "N", Lor = "N", Tereyag = "N", Dondurma = "N", Kolileme = "N", SutAlim = "N", BulasikHane = "N", SevkiyatDepo = "N", AmbalajDepo = "N", Tarihleme = "N", GenelTemizlik = "N", Mutfak = "N", KazanDairesi = "N", BakimOnarim = "N" });
                            }

                            frmGunlukPersonelPlanlama.Mode = BoFormMode.fm_ADD_MODE;
                            oEditTarih.Value = oEditTariheKopyala.Value;

                            string data = string.Join("", rows2.Select(s => string.Format(row, s.PersonelAdi, s.PersonelNo, s.Ayran, s.Yogurt, s.Teleme, s.Kasar, s.Kaymak, s.Lor, s.Tereyag, s.Dondurma, s.Kolileme, s.SutAlim, s.BulasikHane, s.SevkiyatDepo, s.AmbalajDepo, s.Tarihleme, s.GenelTemizlik, s.Mutfak, s.KazanDairesi, s.BakimOnarim)));

                            frmGunlukPersonelPlanlama.DataSources.DBDataSources.Item("@AIF_GUNPERSPLAN1").LoadFromXML(string.Format(header, data));

                            oMatrix.AutoResizeColumns();

                            renklendir();

                            Handler.SAPApplication.MessageBox("Kopyalama işlemi tamamlandı. Ekleme işlemi ile devam edebilirsiniz.");

                            oEditTariheKopyala.Value = "";
                        }
                        catch (Exception ex)
                        {
                        }
                        finally
                        {
                            frmGunlukPersonelPlanlama.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "Item_0" && pVal.BeforeAction)
                    {
                        if (frmGunlukPersonelPlanlama.Mode == BoFormMode.fm_ADD_MODE || frmGunlukPersonelPlanlama.Mode == BoFormMode.fm_UPDATE_MODE)
                        {
                            Handler.SAPApplication.MessageBox("Ekleme veya güncelleme modunda iken kopyalama yapılamaz.");
                            BubbleEvent = false;
                            return false;
                        }
                        else
                            oEditDocEntry.Item.Click();
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

        private void renklendir(bool addbtn = false)
        {
            try
            {
                frmGunlukPersonelPlanlama.Freeze(true);
                if (!addbtn)
                {
                    string xmlColor = oMatrix.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);
                    XDocument xDoc = XDocument.Parse(xmlColor);
                    var table = (from x in xDoc.Elements("Matrix").Elements("Rows").Elements("Row")
                                 select new
                                 {
                                     Ayran = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                     Yogurt = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                                     Teleme = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value,
                                     Kasar = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value,
                                     Kaymak = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,
                                     Lor = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                                     Tereyag = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value,
                                     Dondurma = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value,
                                     Kolileme = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value,
                                     SutAlim = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_9" select new XElement(y.Element("Value"))).First().Value,
                                     BulasikHane = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_10" select new XElement(y.Element("Value"))).First().Value,
                                     SevkiyatDepo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_11" select new XElement(y.Element("Value"))).First().Value,
                                     AmbalajDepo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_12" select new XElement(y.Element("Value"))).First().Value,
                                     Tarihleme = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_13" select new XElement(y.Element("Value"))).First().Value,
                                     GenelTemizlik = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_14" select new XElement(y.Element("Value"))).First().Value,
                                     Mutfak = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_15" select new XElement(y.Element("Value"))).First().Value,
                                     KazanDairesi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_16" select new XElement(y.Element("Value"))).First().Value,
                                     BakimOnarim = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_17" select new XElement(y.Element("Value"))).First().Value,
                                     INDEX = x.ElementsBeforeSelf().Count() + 1
                                 }).ToList();

                    Color val1 = Color.Orange;
                    int Turuncu = ColorTranslator.ToOle(val1);
                    val1 = Color.LightGreen;
                    int yesil = ColorTranslator.ToOle(val1);

                    var orjinalrenk = table.Where(x => x.Ayran == "Y" || x.Yogurt == "Y" || x.Teleme == "Y" || x.Kasar == "Y" || x.Kaymak == "Y" || x.Lor == "Y" || x.Tereyag == "Y" || x.Dondurma == "Y" || x.Kolileme == "Y" || x.SutAlim == "Y" || x.BulasikHane == "Y" || x.SevkiyatDepo == "Y" || x.AmbalajDepo == "Y" || x.Tarihleme == "Y" || x.GenelTemizlik == "Y" || x.Mutfak == "Y" || x.KazanDairesi == "Y" || x.BakimOnarim == "Y").ToList();

                    var customrenk = table.Where(x => x.Ayran != "Y" && x.Yogurt != "Y" && x.Teleme != "Y" && x.Kasar != "Y" && x.Kaymak != "Y" && x.Lor != "Y" && x.Tereyag != "Y" && x.Dondurma != "Y" && x.Kolileme != "Y" && x.SutAlim != "Y" && x.BulasikHane != "Y" && x.SevkiyatDepo != "Y" && x.AmbalajDepo != "Y" && x.Tarihleme != "Y" && x.GenelTemizlik != "Y" && x.Mutfak != "Y" && x.KazanDairesi != "Y" && x.BakimOnarim != "Y").ToList();

                    foreach (var item in orjinalrenk)
                    {
                        oMatrix.CommonSetting.SetRowBackColor(item.INDEX, yesil);
                    }

                    foreach (var item in customrenk)
                    {
                        oMatrix.CommonSetting.SetRowBackColor(item.INDEX, Turuncu);
                    }
                }
                else
                {
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        oMatrix.CommonSetting.SetRowBackColor(i, -1);
                    }
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                frmGunlukPersonelPlanlama.Freeze(false);
            }
        }

        public bool MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.MenuUID == "1282" && !pVal.BeforeAction)
            {
                try
                {
                    frmGunlukPersonelPlanlama.Close();
                }
                catch (Exception)
                {
                }

                Handler.SAPApplication.ActivateMenuItem("GunPersPl");
                //frmGunlukPersonelPlanlama.Mode = BoFormMode.fm_ADD_MODE;
                //frmGunlukPersonelPlanlama.DataSources.DBDataSources.Item(0).Clear();
                //frmGunlukPersonelPlanlama.DataSources.DBDataSources.Item(1).Clear();
                //oEditDocEntry.Item.Click();
                //CalisanListele();
                renklendir(true);
            }

            return BubbleEvent;
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }
    }
}