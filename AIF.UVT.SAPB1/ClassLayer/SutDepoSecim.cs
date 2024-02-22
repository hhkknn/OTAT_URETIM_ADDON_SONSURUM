using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using AIF.UVT.SAPB1.HelperClass;
using AIF.UVT.SAPB1.Models;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class SutDepoSecim
    {
        [ItemAtt(AIFConn.SutDepoSecimUID)]
        public SAPbouiCOM.Form frmSutDepoSecim;

        [ItemAtt("Item_6")]
        public SAPbouiCOM.EditText oEdtSutKabulNo;

        [ItemAtt("Item_1")]
        public SAPbouiCOM.EditText oEdtSutMiktari;

        [ItemAtt("Item_9")]
        public SAPbouiCOM.EditText oEditTiklamalik;

        [ItemAtt("1")]
        public SAPbouiCOM.Button btnAddOrUpdate;

        [ItemAtt("Item_2")]
        public SAPbouiCOM.Matrix oMatrix;

        private string sutKabulNo = "";
        private double sutMiktari = 0;
        private string urunKodu = "";
        private string partiNo = "";
        private string docDate = "";
        private SAPbouiCOM.Form oFormSutKabul = null;

        public void LoadForms(string _sutKabulNo, double _sutMiktari, string _urunKodu, string _partiNo, string _docDate, SAPbouiCOM.Form _oFormSutKabul)
        {
            try
            {
                ConstVariables.oFnc.LoadSAPXML(AIFConn.SutDepoSecimXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.SutDepoSecimXML));
                Functions.CreateUserOrSystemFormComponent<SutDepoSecim>(AIFConn.SutDepo);
                sutKabulNo = _sutKabulNo;
                sutMiktari = _sutMiktari;
                urunKodu = _urunKodu;
                partiNo = _partiNo;
                docDate = _docDate;
                oFormSutKabul = _oFormSutKabul;
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Süt depo seçimi yapılırken hata oluştu. LoadForms - " + ex.Message);
            }
            InitForms();
        }

        private string header = @"<?xml version=""1.0"" encoding=""UTF-16"" ?><dbDataSources uid=""@AIF_SUT_DEPO1""><rows>{0}</rows></dbDataSources>";

        private string row = "<row><cells><cell><uid>U_Tank</uid><value>{0}</value></cell><cell><uid>U_Miktar</uid><value>{1}</value></cell></cells></row>";

        public void InitForms()
        {
            try
            {
                sutTankSecimiOk = false;
                try
                {
                    frmSutDepoSecim.EnableMenu("1283", false);
                    frmSutDepoSecim.EnableMenu("1284", false);
                    frmSutDepoSecim.EnableMenu("1286", false);
                }
                catch (Exception)
                {
                }

                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                ConstVariables.oRecordset.DoQuery("Select \"DocEntry\" from \"@AIF_SUT_DEPO\" as T0 where T0.\"U_SutKabulNo\" = '" + sutKabulNo + "'");

                if (ConstVariables.oRecordset.RecordCount > 0)
                {
                    frmSutDepoSecim.Mode = BoFormMode.fm_FIND_MODE;
                    oEdtSutKabulNo.Item.Enabled = true;
                    oEdtSutKabulNo.Value = sutKabulNo;
                    btnAddOrUpdate.Item.Click();
                    oEditTiklamalik.Item.Click();
                    oEdtSutKabulNo.Item.Enabled = false;
                    frmSutDepoSecim.Mode = BoFormMode.fm_VIEW_MODE;
                }
                else
                {
                    oEdtSutMiktari.Value = sutMiktari.ToString();
                    oEdtSutKabulNo.Value = sutKabulNo;
                    oMatrix.AddRow();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(1).Specific).Value = sutMiktari.ToString();
                }
                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Süt depo seçimi yapılırken hata oluştu." + ex.Message);
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
                    }
                    else if (BusinessObjectInfo.BeforeAction)
                    {
                        try
                        {
                            ConstVariables.oCompanyObject.StartTransaction();
                            string msg = "";
                            msg = Environment.NewLine;
                            msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " depo seçim ekranında eklenme butonuna basıldı. Ve satınalma mal  girişi oluşturma methoduna depo seçimi kontrollerinden sonra gidilecek.";
                            AIFConn.SutKabul.WriteToFile(msg);

                            for (int i = 1; i <= oMatrix.RowCount; i++)
                            {
                                var DepoKodu = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value;

                                if (DepoKodu == "")
                                {
                                    msg = Environment.NewLine;
                                    msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " Tank seçimi boş olan satır olduğundan dolayı ekleme işlemine devam edilemez hatası alındı.";
                                    AIFConn.SutKabul.WriteToFile(msg);
                                    Handler.SAPApplication.MessageBox("Tank seçimi boş olan satır olduğundan dolayı ekleme işlemine devam edilemez.");
                                    BubbleEvent = false;
                                    return false;
                                }
                            }
                            string uyari = "";

                            var satinalmamalgirisiResp = SatinAlmaMalGirisiOlustur(oFormSutKabul);

                            if (satinalmamalgirisiResp.DocEntry == 0)
                            {
                                Handler.SAPApplication.MessageBox(satinalmamalgirisiResp.Description);

                                try
                                {
                                    ConstVariables.oCompanyObject.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                }
                                catch (Exception)
                                {
                                }
                            }
                            else
                            {
                                uyari += Environment.NewLine + "Satınalma mal girişi oluşturuldu. Satınalma mal girişi numarası : " + satinalmamalgirisiResp.DocEntry;
                            }
                            //bool check = SatinAlmaMalGirisiOlustur(oFormSutKabul);
                            //if (!check)
                            //{
                            //    BubbleEvent = false;
                            //    return false;
                            //}

                            string xml = frmSutDepoSecim.DataSources.DBDataSources.Item(0).GetAsXML();
                            //XmlDocument key = new XmlDocument();
                            //key.LoadXml(BusinessObjectInfo.ObjectKey);
                            //string docentry = key.SelectNodes("AIF_SUT_DEPOParams").Item(0).SelectNodes("DocEntry").Item(0).InnerXml;

                            //if (docentry != "")
                            //{
                            ProductionOrder productionOrder = new ProductionOrder();
                            productionOrder.ItemCode = "YRM-10001";  //urunKodu;
                            productionOrder.Miktar = sutMiktari;
                            var uretimSiparisiOlusturResp = addProductionOrders(productionOrder);
                            int uretimSiparisNo = 0;
                            if (uretimSiparisiOlusturResp.DocEntry > 0)
                            {
                                uyari = "Üretim siparişi başarıyla oluşturuldu. Üretim sipariş numarası : " + uretimSiparisiOlusturResp.DocEntry + "";
                                uretimSiparisNo = uretimSiparisiOlusturResp.DocEntry;
                                msg += Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " Üretim siparişi oluşturma başarılı." + uretimSiparisiOlusturResp.DocEntry;
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " Üretim siparişi oluşturma başarısız hatası aldı." + uretimSiparisiOlusturResp.Description;
                                AIFConn.SutKabul.WriteToFile(msg);
                                BubbleEvent = false;
                                uyari = uretimSiparisiOlusturResp.Description;
                                Handler.SAPApplication.MessageBox(uyari);
                                try
                                {
                                    ConstVariables.oCompanyObject.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                }
                                catch (Exception)
                                {
                                }
                                return false;
                            }

                            //OtatServiceSoapClient otatServiceSoapClient = new OtatServiceSoapClient();

                            //ProductionOrder productionOrder = new ProductionOrder();
                            //productionOrder.ItemCode = "YRM-10001";  //urunKodu;
                            //productionOrder.Miktar = sutMiktari;

                            //var resp = otatServiceSoapClient.AddProductionOrders(productionOrder, ConstVariables.oCompanyObject.CompanyDB);

                            //if (resp.Value == 0)
                            //{
                            //    if (uyari == "")
                            //    {
                            //        uyari = "Üretim siparişi başarıyla oluşturuldu. Üretim sipariş numarası : " + resp.DocEntry + "";
                            //    }

                            //    msg = Environment.NewLine;
                            //    msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " Üretim siparişi oluşturma başarılı." + resp.DocEntry;
                            //    AIFConn.SutKabul.WriteToFile(msg);
                            //}
                            //else
                            //{
                            //    if (uyari == "")
                            //    {
                            //        msg = Environment.NewLine;
                            //        msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " Üretim siparişi oluşturma başarısız hatası aldı." + resp.Description;
                            //        AIFConn.SutKabul.WriteToFile(msg);
                            //        BubbleEvent = false;
                            //        uyari = resp.Description;
                            //        Handler.SAPApplication.MessageBox(uyari);
                            //        break;
                            //    }
                            //}

                            //System.Data.DataTable dataTable = new System.Data.DataTable();
                            //dataTable = otatServiceSoapClient.GetProductionbyDocEntry(resp.DocEntry).List;

                            Parti parti = new Parti();
                            List<Parti> partis = new List<Parti>();
                            InventoryGenExist inventoryGenExist = new InventoryGenExist();
                            List<InventoryGenExist> inventoryGenExists = new List<InventoryGenExist>();
                            inventoryGenExist.UretimSiparisi = uretimSiparisNo;
                            inventoryGenExist.SatirNumarasi = 0;
                            inventoryGenExist.Miktar = sutMiktari;
                            inventoryGenExist.DepoKodu = "200";
                            parti.PartiNo = partiNo;
                            parti.PartikMiktar = sutMiktari;

                            partis.Add(parti);

                            inventoryGenExist.Parti = partis;

                            inventoryGenExists.Add(inventoryGenExist);

                            var uretimdenCikisResp = addInventoryGenExist(inventoryGenExists);

                            if (uretimdenCikisResp.DocEntry > 0)
                            {
                                uyari += Environment.NewLine + "Üretimden çıkış başarıyla oluşturuldu. Üretimden çıkış numarası : " + uretimdenCikisResp.DocEntry;

                                msg += Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " Üretimden çıkış oluşturma başarılı." + uretimdenCikisResp.DocEntry;
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " Üretimden çıkış oluşturma başarısız hatası aldı." + uretimdenCikisResp.Description;
                                AIFConn.SutKabul.WriteToFile(msg);
                                BubbleEvent = false;
                                uyari = Environment.NewLine + uretimdenCikisResp.Description;
                                Handler.SAPApplication.MessageBox(uyari);
                                try
                                {
                                    ConstVariables.oCompanyObject.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                }
                                catch (Exception ex)
                                {
                                }
                                return false;
                            }

                            //Parti parti = new Parti();
                            //List<Parti> partis = new List<Parti>();
                            //InventoryGenExist inventoryGenExist = new InventoryGenExist();
                            //List<InventoryGenExist> inventoryGenExists = new List<InventoryGenExist>();
                            //inventoryGenExist.UretimSiparisi = resp.DocEntry;
                            //inventoryGenExist.SatirNumarasi = 0;
                            //inventoryGenExist.Miktar = sutMiktari;
                            //inventoryGenExist.DepoKodu = "200";
                            //parti.PartiNo = partiNo;
                            //parti.PartikMiktar = sutMiktari;

                            //partis.Add(parti);

                            //inventoryGenExist.Parti = partis.ToArray();

                            //inventoryGenExists.Add(inventoryGenExist);

                            //var resp2 = otatServiceSoapClient.AddInventoryGenExist(inventoryGenExists.ToArray(), ConstVariables.oCompanyObject.CompanyDB);

                            //if (resp2.Value == 0)
                            //{
                            //    uyari += Environment.NewLine + "Üretimden çıkış başarıyla oluşturuldu. Üretimden çıkış numarası : " + resp2.DocEntry;

                            //    msg = Environment.NewLine;
                            //    msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " Üretimden çıkış oluşturma başarılı." + resp2.DocEntry;
                            //    AIFConn.SutKabul.WriteToFile(msg);
                            //}
                            //else
                            //{
                            //    msg = Environment.NewLine;
                            //    msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " Üretimden çıkış oluşturma başarısız hatası aldı." + resp2.Description;
                            //    AIFConn.SutKabul.WriteToFile(msg);
                            //    BubbleEvent = false;
                            //    uyari = Environment.NewLine + resp2.Description;
                            //    Handler.SAPApplication.MessageBox(uyari);
                            //    break;
                            //}

                            InventoryGenEntry inventoryGenEntry = new InventoryGenEntry();
                            List<InventoryGenEntry> inventoryGenEntries = new List<InventoryGenEntry>();

                            for (int i = 1; i <= oMatrix.RowCount; i++)
                            {
                                inventoryGenEntry.UretimSiparisi = uretimSiparisNo;

                                inventoryGenEntry.DepoKodu = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value;
                                inventoryGenEntry.Miktar = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific).Value);

                                inventoryGenEntry.Parti = partiNo;
                                inventoryGenEntry.PartiMiktar = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific).Value);

                                inventoryGenEntries.Add(inventoryGenEntry);
                                inventoryGenEntry = new InventoryGenEntry();
                            }

                            if (inventoryGenEntries.Count > 0)
                            {
                                var uretimdenGirisResp = addInventoryGenEntry(inventoryGenEntries);
                                //var resp3 = otatServiceSoapClient.AddInventoryGenEntry(inventoryGenEntries.ToArray(), ConstVariables.oCompanyObject.CompanyDB);

                                if (uretimdenGirisResp.Value == 0)
                                {
                                    uyari += Environment.NewLine + "Üretimden giriş başarıyla oluşturuldu. Üretimden giriş numarası : " + uretimdenGirisResp.DocEntry;

                                    msg += Environment.NewLine;
                                    msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " Üretimden giriş oluşturma başarılı." + uretimdenGirisResp.DocEntry;
                                    AIFConn.SutKabul.WriteToFile(msg);
                                }
                                else
                                {
                                    msg = Environment.NewLine;
                                    msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " Üretimden giriş oluşturma başarısız hatası aldı." + uretimdenGirisResp.Description;
                                    AIFConn.SutKabul.WriteToFile(msg);
                                    BubbleEvent = false;
                                    uyari = Environment.NewLine + uretimdenGirisResp.Description;
                                    Handler.SAPApplication.MessageBox(uyari);
                                    try
                                    {
                                        ConstVariables.oCompanyObject.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    }
                                    catch (Exception)
                                    {
                                    }
                                    return false;
                                }
                            }

                            //InventoryGenEntry inventoryGenEntry = new InventoryGenEntry();
                            //List<InventoryGenEntry> inventoryGenEntries = new List<InventoryGenEntry>();

                            //for (int i = 1; i <= oMatrix.RowCount; i++)
                            //{
                            //    inventoryGenEntry.UretimSiparisi = resp.DocEntry;

                            //    inventoryGenEntry.DepoKodu = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value;
                            //    inventoryGenEntry.Miktar = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific).Value);

                            //    inventoryGenEntry.Parti = partiNo;
                            //    inventoryGenEntry.PartiMiktar = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific).Value);

                            //    inventoryGenEntries.Add(inventoryGenEntry);
                            //    inventoryGenEntry = new InventoryGenEntry();
                            //}

                            //if (inventoryGenEntries.Count > 0)
                            //{
                            //    var resp3 = otatServiceSoapClient.AddInventoryGenEntry(inventoryGenEntries.ToArray(), ConstVariables.oCompanyObject.CompanyDB);

                            //    if (resp3.Value == 0)
                            //    {
                            //        uyari += Environment.NewLine + "Üretimden giriş başarıyla oluşturuldu. Üretimden giriş numarası : " + resp3.DocEntry;

                            //        msg = Environment.NewLine;
                            //        msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " Üretimden giriş oluşturma başarılı." + resp3.DocEntry;
                            //        AIFConn.SutKabul.WriteToFile(msg);
                            //    }
                            //    else
                            //    {
                            //        msg = Environment.NewLine;
                            //        msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " Üretimden giriş oluşturma başarısız hatası aldı." + resp3.Description;
                            //        AIFConn.SutKabul.WriteToFile(msg);
                            //        BubbleEvent = false;
                            //        uyari = Environment.NewLine + resp3.Description;
                            //        Handler.SAPApplication.MessageBox(uyari);
                            //        break;
                            //    }
                            //}
                            ekleme = true;
                            Handler.SAPApplication.MessageBox(uyari);
                            //}

                            //if (oFormSutKabul.Mode == BoFormMode.fm_UPDATE_MODE)
                            //{
                            sutkabultamamlaniyor = true;
                            oFormSutKabul.Mode = BoFormMode.fm_UPDATE_MODE;
                            ((SAPbouiCOM.ComboBox)oFormSutKabul.Items.Item("Item_166").Specific).Select("C");
                            oFormSutKabul.Items.Item("1").Click();
                            sutkabultamamlaniyor = false;
                            //}

                            try
                            {
                                ConstVariables.oCompanyObject.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                oFormSutKabul.Items.Item("Item_53").Click();
                                AIFConn.SutKabul.statuKapaliIslemleri();
                            }
                            catch (Exception)
                            {
                            }
                        }
                        catch (Exception ex)
                        {
                            try
                            {
                                ConstVariables.oCompanyObject.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }
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

        public static bool sutkabultamamlaniyor = false;
        public static bool sutTankSecimiOk = false;
        private string val = "";
        private bool ekleme = false;

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
                        try
                        {
                            if (ekleme)
                            {
                                frmSutDepoSecim.Close();
                                ekleme = false;
                                sutTankSecimiOk = true;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
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
                    if (pVal.ColUID == "Col_1" && pVal.ItemChanged && !pVal.BeforeAction)
                    {
                        string msg = "";

                        try
                        {
                            msg = Environment.NewLine;
                            msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için tank seçimi yapıldı.";
                            AIFConn.SutKabul.WriteToFile(msg);

                            frmSutDepoSecim.Freeze(true);
                            var deger = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value);
                            if (deger > 0)
                            {
                                string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                            select new
                                            {
                                                DepoKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                                Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value),
                                            }).ToList();

                                double siparisMiktari = sutMiktari;
                                double matrixtoplammiktar = parseNumber.parservalues<double>(rows.Sum(x => x.Miktar).ToString());
                                double sonuc = siparisMiktari - matrixtoplammiktar;
                                if (sonuc > 0)
                                {
                                    rows.Insert(rows.Count, new { DepoKodu = "", Miktar = siparisMiktari - matrixtoplammiktar });

                                    string data = string.Join("", rows.Select(s => string.Format(row, s.DepoKodu, s.Miktar)));

                                    frmSutDepoSecim.DataSources.DBDataSources.Item("@AIF_SUT_DEPO1").LoadFromXML(string.Format(header, data));

                                    oMatrix.AutoResizeColumns();

                                    frmSutDepoSecim.Items.Item("Item_9").Click();
                                    for (int i = 1; i <= oMatrix.RowCount - 1; i++)
                                    {
                                        oMatrix.CommonSetting.SetCellEditable(i, 2, false);
                                    }
                                }
                                else if (sonuc < 0)
                                {
                                    msg = Environment.NewLine;
                                    msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " " + string.Format("Kalan Miktar {0}'dır kalan kadar değer otomatik atanmıştır.", deger - (sonuc * -1)) + " hatası alındı.";
                                    AIFConn.SutKabul.WriteToFile(msg);

                                    Handler.SAPApplication.MessageBox(string.Format("Kalan Miktar {0}'dır kalan kadar değer otomatik atanmıştır.", deger - (sonuc * -1)));
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = parseNumber.parservalues<double>((deger - (sonuc * -1)).ToString()).ToString();
                                }
                            }
                            else if (deger == 0)
                            {
                                if (pVal.Row > 1)
                                {
                                    oMatrix.DeleteRow(pVal.Row);
                                    oMatrix.CommonSetting.SetCellEditable(pVal.Row - 1, 2, true);

                                    string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                    var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                                select new
                                                {
                                                    DepoKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                                    Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value)
                                                }).ToList();

                                    double sonsatir = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(oMatrix.RowCount).Specific).Value);
                                    double siparisMiktari = sutMiktari;
                                    double matrixtoplammiktar = parseNumber.parservalues<double>(rows.Sum(x => x.Miktar).ToString());
                                    double sonuc = siparisMiktari - (matrixtoplammiktar - sonsatir);
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(oMatrix.RowCount).Specific).Value = sonuc.ToString();
                                    //((SAPbouiCOM.EditText)oMatrixUretimSiparisNo.Columns.Item("Col_2").Cells.Item(oMatrixUretimSiparisNo.RowCount).Specific).Value = (siparisMiktari / sonuc).ToString();
                                }
                                else if (pVal.Row == 1)
                                {
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(oMatrix.RowCount).Specific).Value = sutMiktari.ToString();
                                    oMatrix.AutoResizeColumns();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            msg = Environment.NewLine;
                            msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " " + " Hesaplama yapılırken hata oluştu. Hata Kodu A0018 hatası alındı.";
                            AIFConn.SutKabul.WriteToFile(msg);
                            Handler.SAPApplication.StatusBar.SetText("Hesaplama yapılırken hata oluştu. Hata Kodu A0018", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        }
                        finally
                        {
                            frmSutDepoSecim.Freeze(false);
                        }
                    }
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
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                        oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                        SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                        oCFL = frmSutDepoSecim.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                        SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                        SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                        SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                        oCFL.SetConditions(oEmptyConts);
                        oCons = oCFL.GetConditions();

                        oCon = oCons.Add();
                        oCon.Alias = "InActive";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "N";

                        oCon.Relationship = BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "U_DepoTipi";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "01";

                        oCFL.SetConditions(oCons);
                    }
                    else if (pVal.ColUID == "Col_0" && !pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                            string Val = "";
                            Val = oDataTable.GetValue("WhsCode", 0).ToString();
                            try
                            {
                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value = Val.ToString();
                            }
                            catch (Exception)
                            {
                            }
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
                if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && pVal.BeforeAction)
                {
                    int row = oMatrix.GetNextSelectedRow();
                    if (row != -1)
                    {
                        oMatrix.DeleteRow(row);
                        //string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                        //var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                        //            select new AnalizBilgileri
                        //            {
                        //                Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value),
                        //            }).ToList();

                        //double total = rows.Sum(x => x.Miktar);

                        //oedtIrsaliyeNetSutMiktari.Value = total.ToString();
                    }
                }
                else if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && pVal.BeforeAction)
                {
                    frmSutDepoSecim.DataSources.DBDataSources.Item(1).Clear();
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

        public Response SatinAlmaMalGirisiOlustur(SAPbouiCOM.Form _oFormSutKabul) //1 tekil giriş, 2 Çoğul Giriş anlamına gelir.
        {
            SAPbouiCOM.CheckBox oChkTekilGiris = (SAPbouiCOM.CheckBox)_oFormSutKabul.Items.Item("Item_26").Specific;
            string tekilCogul = oChkTekilGiris.Checked == true ? "1" : "2";
            string saticiKodu = ((SAPbouiCOM.EditText)_oFormSutKabul.Items.Item("Item_159").Specific).Value;
            string docEntry = ((SAPbouiCOM.EditText)_oFormSutKabul.Items.Item("Item_128").Specific).Value;
            string kantarNo = ((SAPbouiCOM.EditText)_oFormSutKabul.Items.Item("Item_144").Specific).Value;
            SAPbouiCOM.Matrix oMatrixAnalizBilgileri = (SAPbouiCOM.Matrix)_oFormSutKabul.Items.Item("Item_77").Specific;
            int _docEntry = 0;
            ConstVariables.oRecordset.DoQuery("Select \"U_SutKabulId\",\"DocDate\",\"DocEntry\" from \"OPDN\" where \"U_SutKabulId\" = '" + docEntry + "'");
            string msg = "";
            if (ConstVariables.oRecordset.RecordCount > 0)
            {
                msg = Environment.NewLine;
                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " daha önce süt kabul işlemini tamamla işlemi yapılmıştır. Kontrolüne takıldı.";
                AIFConn.SutKabul.WriteToFile(msg);
                partiNo = Convert.ToDateTime(ConstVariables.oRecordset.Fields.Item("DocDate").Value).ToString("yyyyMMdd") + "-" + ConstVariables.oRecordset.Fields.Item("DocEntry").Value;
                return new Response { Value = 0, Description = "süt kabul işlemini tamamla işlemi yapılmıştır.", List = null, DocEntry = Convert.ToInt32(docEntry) };
            }
            //SAPbobsCOM.Documents oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.oPurchaseDeliveryNotes);

            string xml = oMatrixAnalizBilgileri.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
            var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                        select new AnalizBilgileri
                        {
                            Tedarikci = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                            TedarikciAdi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_22" select new XElement(y.Element("Value"))).First().Value,
                            Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value),
                            Sicaklik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value),
                            Briks = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value),
                            Yag = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value),
                            Yogunluk = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value),
                            pH = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value),
                            SH = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value),
                            SutTuru = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_9" select new XElement(y.Element("Value"))).First().Value,
                            //TemizlikKoku = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_10" select new XElement(y.Element("Value"))).First().Value,
                            Soda = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_11" select new XElement(y.Element("Value"))).First().Value,
                            Antibiyotik = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_12" select new XElement(y.Element("Value"))).First().Value,
                            Ykm = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_13" select new XElement(y.Element("Value"))).First().Value),
                            Alkol = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_14" select new XElement(y.Element("Value"))).First().Value,
                            KaynatSh = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_15" select new XElement(y.Element("Value"))).First().Value),
                            Depo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_16" select new XElement(y.Element("Value"))).First().Value,
                            Protein = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_17" select new XElement(y.Element("Value"))).First().Value),
                            Laktoz = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_18" select new XElement(y.Element("Value"))).First().Value),
                            KatilanSu = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_19" select new XElement(y.Element("Value"))).First().Value),
                            SomatikHucre = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_20" select new XElement(y.Element("Value"))).First().Value),
                            MetilenMavisi = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_21" select new XElement(y.Element("Value"))).First().Value),
                            row = Convert.ToString(x.ElementsBeforeSelf().Count() + 1),
                            SutTemizlik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_28" select new XElement(y.Element("Value"))).First().Value),
                            TankTemizlik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_29" select new XElement(y.Element("Value"))).First().Value),
                            EkipmanTemizlik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_30" select new XElement(y.Element("Value"))).First().Value),
                            AracKasasiTemizlik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_31" select new XElement(y.Element("Value"))).First().Value),
                            AgirlikliPuan = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_32" select new XElement(y.Element("Value"))).First().Value),
                            IrsaliyeNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_10" select new XElement(y.Element("Value"))).First().Value,
                        }).ToList();
            int i = 0;

            //OtatServiceSoapClient otatServiceSoapClient = new OtatServiceSoapClient();
            PurchaseDeliveryNotes purchaseDeliveryNotes = new PurchaseDeliveryNotes();
            //Süt türü 1 ve 2 yani yağlı ve yağsız için toplam 1 kalem üretileceği için gruplamada tek satır gelmesi için süt türü 2 ler 1 olarak güncellendi.
            rows.Where(x => x.SutTuru == "2").ToList().ForEach(x => x.SutTuru = "1");

            var result = rows.GroupBy(x => x.SutTuru).Select(y => new AnalizBilgileri
            {
                SutTuru = y.First().SutTuru,
                Miktar = y.Sum(z => z.Miktar)
            });

            var resultwithGroupBy = rows.GroupBy(x => new { x.Tedarikci });

            purchaseDeliveryNotes.BPLId = 1;

            //string date = Convert.ToDateTime(docDate).ToString("yyyymmdd");//chn
            //purchaseDeliveryNotes.DocDate = docDate;//chn
            purchaseDeliveryNotes.DocDate = DateTime.Now.ToString("yyyyMMdd");
            purchaseDeliveryNotes.userId = ConstVariables.oCompanyObject.UserSignature;
            purchaseDeliveryNotes.sutKabulId = Convert.ToInt32(docEntry);
            PurchaseDeliveryNotesDetail purchaseDeliveryNotesDetail = new PurchaseDeliveryNotesDetail();
            List<PurchaseDeliveryNotesDetail> purchaseDeliveryNotesDetails = new List<PurchaseDeliveryNotesDetail>();
            PurchaseDeliveryNotesDetailParti purchaseDeliveryNotesDetailParti = new PurchaseDeliveryNotesDetailParti();
            List<PurchaseDeliveryNotesDetailParti> purchaseDeliveryNotesDetailPartis = new List<PurchaseDeliveryNotesDetailParti>();

            if (tekilCogul == "1")
            {
                purchaseDeliveryNotes.CardCode = saticiKodu;
                foreach (var itemAif in result)
                {
                    if (itemAif.SutTuru == "1" || itemAif.SutTuru == "2")
                    {
                        purchaseDeliveryNotesDetail.ItemCode = "HAM-00001";
                    }

                    urunKodu = purchaseDeliveryNotesDetail.ItemCode;

                    purchaseDeliveryNotesDetail.WareHouse = "200";
                    purchaseDeliveryNotesDetail.Quantity = itemAif.Miktar;

                    purchaseDeliveryNotesDetailParti.BatchNumber = DateTime.Now.Year + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0') + "-" + docEntry;
                    purchaseDeliveryNotesDetailParti.Quantity = itemAif.Miktar;

                    partiNo = purchaseDeliveryNotesDetailParti.BatchNumber;
                    purchaseDeliveryNotesDetailPartis.Add(purchaseDeliveryNotesDetailParti);

                    purchaseDeliveryNotesDetail.PartiDetails = purchaseDeliveryNotesDetailPartis;

                    purchaseDeliveryNotesDetails.Add(purchaseDeliveryNotesDetail);
                }

                purchaseDeliveryNotes.Lines = purchaseDeliveryNotesDetails;

                if (purchaseDeliveryNotes.Lines.Count() > 0)
                {
                    string irsaliyeNo = "";
                    if (rows.Where(x => x.IrsaliyeNo != null && x.IrsaliyeNo != "").Count() > 0)
                    {
                        //irsaliyeNo = rows.Where(x => x.Tedarikci == purchaseDeliveryNotes.CardCode).Select(y => y.IrsaliyeNo).FirstOrDefault().ToString();
                        irsaliyeNo = rows.Where(x => x.IrsaliyeNo != "").Select(y => y.IrsaliyeNo).First().ToString();
                    }

                    var resp = addPurchaseDeliveryNotes(purchaseDeliveryNotes, irsaliyeNo, kantarNo);
                    //var resp = otatServiceSoapClient.AddOrUpdatePurchaseDeliveryNotes(purchaseDeliveryNotes, ConstVariables.oCompanyObject.CompanyDB);

                    if (resp.Value == 0)
                    {
                        msg = Environment.NewLine;
                        msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " satınalma mal girişi oluşturma başarılı." + resp.DocEntry;
                        AIFConn.SutKabul.WriteToFile(msg);
                        ((SAPbouiCOM.EditText)_oFormSutKabul.Items.Item("Item_164").Specific).Value = resp.DocEntry.ToString();
                        //_oFormSutKabul.Items.Item("1").Click();
                        //oEditBelgeNo.Value = resp.DocEntry.ToString();
                        //oComboBelgeDurumu.Select("C");
                        //oButtonAddOrUpdate.Item.Click();

                        return new Response { Value = 0, Description = "Satınalma mal girişi başarıyla oluşturuldu.", List = null, DocEntry = Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey()) };
                    }
                    else
                    {
                        msg = Environment.NewLine;
                        msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " Satınalma belgesi oluşturulurken hata oluştu." + resp.Description;
                        AIFConn.SutKabul.WriteToFile(msg);

                        return new Response { Value = 0, Description = "Satınalma belgesi oluşturulurken hata oluştu." + resp.Description, List = null, DocEntry = 0 };

                        //Handler.SAPApplication.MessageBox("Satınalma belgesi oluşturulurken hata oluştu. " + resp.Description);
                        //return false;
                    }
                }
            }
            else
            {
                foreach (var item in resultwithGroupBy)
                {
                    purchaseDeliveryNotes.CardCode = item.Key.Tedarikci.ToString();

                    var sumQty = rows.Where(x => x.Tedarikci == item.Key.Tedarikci.ToString()).Select(z => z.Miktar).Sum();

                    purchaseDeliveryNotesDetail.ItemCode = "HAM-00001";

                    purchaseDeliveryNotesDetail.WareHouse = "200";
                    purchaseDeliveryNotesDetail.Quantity = sumQty;

                    purchaseDeliveryNotesDetailParti.BatchNumber = DateTime.Now.Year + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0') + "-" + docEntry;
                    purchaseDeliveryNotesDetailParti.Quantity = sumQty;

                    partiNo = purchaseDeliveryNotesDetailParti.BatchNumber;
                    purchaseDeliveryNotesDetailPartis.Add(purchaseDeliveryNotesDetailParti);

                    purchaseDeliveryNotesDetail.PartiDetails = purchaseDeliveryNotesDetailPartis;

                    purchaseDeliveryNotesDetails.Add(purchaseDeliveryNotesDetail);

                    purchaseDeliveryNotes.Lines = purchaseDeliveryNotesDetails;

                    if (purchaseDeliveryNotes.Lines.Count() > 0)
                    {
                        string irsaliyeNo = "";
                        if (rows.Where(x => x.Tedarikci == purchaseDeliveryNotes.CardCode).Select(y => y.IrsaliyeNo).FirstOrDefault() != null && rows.Where(x => x.Tedarikci == purchaseDeliveryNotes.CardCode).Select(y => y.IrsaliyeNo).FirstOrDefault().ToString() != "")
                        {
                            irsaliyeNo = rows.Where(x => x.Tedarikci == purchaseDeliveryNotes.CardCode).Select(y => y.IrsaliyeNo).FirstOrDefault().ToString();
                        }

                        var resp = addPurchaseDeliveryNotes(purchaseDeliveryNotes, irsaliyeNo, kantarNo);

                        //var resp = otatServiceSoapClient.AddOrUpdatePurchaseDeliveryNotes(purchaseDeliveryNotes, ConstVariables.oCompanyObject.CompanyDB);

                        purchaseDeliveryNotesDetail = new PurchaseDeliveryNotesDetail();
                        purchaseDeliveryNotesDetails = new List<PurchaseDeliveryNotesDetail>();
                        purchaseDeliveryNotesDetailParti = new PurchaseDeliveryNotesDetailParti();
                        purchaseDeliveryNotesDetailPartis = new List<PurchaseDeliveryNotesDetailParti>();

                        if (resp.Value == 0)
                        {
                            string olusanbelgelistesi = "";
                            msg = Environment.NewLine;
                            var siraNo = rows.Where(x => x.Tedarikci == item.Key.Tedarikci.ToString()).ToList();
                            foreach (var itemAif in siraNo)
                            {
                                ((SAPbouiCOM.EditText)oMatrixAnalizBilgileri.Columns.Item("Col_23").Cells.Item(Convert.ToInt32(itemAif.row)).Specific).Value = resp.DocEntry.ToString();
                                olusanbelgelistesi += " " + resp.DocEntry;
                            }

                            msg += "İşlem zamanı:" + DateTime.Now + "--" + oEdtSutKabulNo.Value + " numaralı süt kabul belgesi için " + " satınalma mal girişi oluşturma başarılı. Çoklu giriş belgesi olduğundan oluşan belgeler sırasıyla " + olusanbelgelistesi;
                            AIFConn.SutKabul.WriteToFile(msg);

                            _docEntry = resp.DocEntry;
                            //oEditBelgeNo.Value = resp.DocEntry.ToString();
                            //oComboBelgeDurumu.Select("C");
                            //oButtonAddOrUpdate.Item.Click();
                        }
                        else
                        {
                            return new Response { Value = 0, Description = "Satınalma belgeleri oluşturulurken hata oluştu." + resp.Description, List = null, DocEntry = 0 };
                            //Handler.SAPApplication.MessageBox("Satınalma belgeleri oluşturulurken hata oluştu. " + resp.Description);
                            //return false;
                        }
                    }
                }

                //_oFormSutKabul.Mode = BoFormMode.fm_UPDATE_MODE;
                //_oFormSutKabul.Items.Item("1").Click();

                //if (_oFormSutKabul.Mode == BoFormMode.fm_UPDATE_MODE)
                //{
                //    //oComboBelgeDurumu.Select("C");
                //}
            }

            return new Response { Value = 0, Description = "Satınalma mal girişi/girişleri başarıyla oluşturuldu.", List = null, DocEntry = _docEntry };
        }

        public Response addProductionOrders(ProductionOrder productionOrder)
        {
            try
            {
                ProductionOrders wo = (ProductionOrders)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.oProductionOrders);

                wo.ItemNo = productionOrder.ItemCode;
                wo.DueDate = DateTime.Now;
                wo.ProductionOrderType = BoProductionOrderTypeEnum.bopotStandard;
                wo.PlannedQuantity = productionOrder.Miktar;
                int retVal = wo.Add();

                if (retVal == 0)
                {
                    wo.GetByKey(Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey()));
                    wo.ProductionOrderStatus = BoProductionOrderStatusEnum.boposReleased;
                    wo.Update();
                    return new Response { Value = 0, Description = "Üretim siparişi başarıyla oluşturuldu.", List = null, DocEntry = Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey()) };
                }
                else
                {
                    return new Response { Value = -7200, Description = "Üretim siparişi oluşturulurken hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription(), List = null, DocEntry = 0 };
                }
            }
            catch (Exception ex)
            {
                return new Response { Value = 9000, Description = "Bilinmeyen hata oluştu. " + ex.Message, List = null, DocEntry = 0 };
            }
        }

        public Response addInventoryGenExist(List<InventoryGenExist> goodIssues)
        {
            try
            {
                SAPbobsCOM.Documents oGenExist = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

                oGenExist.DocDate = DateTime.Now;
                oGenExist.BPL_IDAssignedToInvoice = 1;
                oGenExist.UserFields.Fields.Item("U_KullaniciId").Value = ConstVariables.oCompanyObject.UserSignature.ToString();
                oGenExist.UserFields.Fields.Item("U_SutKabulId").Value = oEdtSutKabulNo.Value;
                oGenExist.UserFields.Fields.Item("U_RotaCode").Value = goodIssues[0].RotaKodu == null ? "" : goodIssues[0].RotaKodu;
                oGenExist.UserFields.Fields.Item("U_BatchNumber").Value = goodIssues[0].PartiNo == null ? "" : goodIssues[0].PartiNo;

                int i = 0;
                foreach (var item in goodIssues)
                {
                    oGenExist.Lines.BaseType = 202;
                    oGenExist.Lines.BaseEntry = item.UretimSiparisi;
                    oGenExist.Lines.BaseLine = item.SatirNumarasi;
                    oGenExist.Lines.Quantity = item.Miktar;

                    if (item.DepoKodu != null || item.DepoKodu != "")
                    {
                        oGenExist.Lines.WarehouseCode = item.DepoKodu;
                    }

                    foreach (var itemx in item.Parti)
                    {
                        if (i != 0)
                        {
                            oGenExist.Lines.BatchNumbers.Add();
                        }
                        oGenExist.Lines.BatchNumbers.SetCurrentLine(i);
                        oGenExist.Lines.BatchNumbers.BatchNumber = itemx.PartiNo;
                        oGenExist.Lines.BatchNumbers.Quantity = itemx.PartikMiktar;
                        i++;
                    }

                    oGenExist.Lines.Add();

                    i = 0;
                }

                var aa = oGenExist.Add();

                if (aa == 0)
                {
                    return new Response { Value = 0, Description = "Üretim için çıkış başarıyla oluşturuldu.", List = null, DocEntry = Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey()) };
                }
                else

                {
                    return new Response { Value = -4100, Description = "Hata Kodu - 4100 Üretim için çıkış oluşturulurken hata oluştu. " + ConstVariables.oCompanyObject.GetLastErrorDescription(), List = null, DocEntry = 0 };
                }
            }
            catch (Exception ex)
            {
                return new Response { Value = 9000, Description = "Bilinmeyen hata oluştu. " + ex.Message, List = null, DocEntry = 0 };
            }
        }

        public Response addInventoryGenEntry(List<InventoryGenEntry> inventoryGenEntries)
        {
            try
            {
                SAPbobsCOM.Documents oGenEntry = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

                oGenEntry.DocDate = DateTime.Now;
                oGenEntry.BPL_IDAssignedToInvoice = 1;
                oGenEntry.UserFields.Fields.Item("U_SutKabulId").Value = oEdtSutKabulNo.Value;
                oGenEntry.UserFields.Fields.Item("U_KullaniciId").Value = ConstVariables.oCompanyObject.UserSignature.ToString();
                int i = 0;
                bool check = false;
                foreach (var item in inventoryGenEntries.Where(x => x.Miktar > 0))
                {
                    oGenEntry.Lines.BaseType = 202;
                    oGenEntry.Lines.BaseEntry = item.UretimSiparisi;
                    //oGenEntry.Lines.BaseLine = 0;// item.SatirNumarasi;
                    oGenEntry.Lines.Quantity = item.Miktar;
                    oGenEntry.Lines.TransactionType = SAPbobsCOM.BoTransactionTypeEnum.botrntComplete;

                    if (item.DepoKodu != null || item.DepoKodu != "")
                    {
                        oGenEntry.Lines.WarehouseCode = item.DepoKodu;
                    }

                    //foreach (var itemx in item.Parti)
                    //{
                    //if (i != 0)
                    //{
                    //oGenEntry.Lines.BatchNumbers.Add();
                    //}
                    oGenEntry.Lines.BatchNumbers.SetCurrentLine(i);
                    oGenEntry.Lines.BatchNumbers.BatchNumber = item.Parti;
                    oGenEntry.Lines.BatchNumbers.Quantity = item.PartiMiktar;
                    //i++;
                    //}

                    oGenEntry.Lines.Add();
                    check = true;
                }

                int aa = 0;

                if (check)
                {
                    aa = oGenEntry.Add();
                }
                i = 0;

                if (inventoryGenEntries.Where(z => z.Miktar < 0).Count() > 0)
                {
                    oGenEntry = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

                    oGenEntry.DocDate = DateTime.Now;
                    oGenEntry.BPL_IDAssignedToInvoice = 1;
                    oGenEntry.UserFields.Fields.Item("U_SutKabulId").Value = oEdtSutKabulNo.Value;
                    oGenEntry.UserFields.Fields.Item("U_KullaniciId").Value = ConstVariables.oCompanyObject.UserSignature;

                    oGenEntry.UserFields.Fields.Item("U_RotaCode").Value = inventoryGenEntries[0].RotaKodu;
                    oGenEntry.UserFields.Fields.Item("U_BatchNumber").Value = inventoryGenEntries[0].Parti;
                    foreach (var item in inventoryGenEntries.Where(x => x.Miktar < 0))
                    {
                        oGenEntry.Lines.BaseType = 202;
                        oGenEntry.Lines.BaseEntry = item.UretimSiparisi;

                        if (item.SiraNo != "")
                        {
                            oGenEntry.Lines.BaseLine = Convert.ToInt32(item.SiraNo);
                        }
                        oGenEntry.Lines.Quantity = item.Miktar * -1;
                        //oGenEntry.Lines.TransactionType = SAPbobsCOM.BoTransactionTypeEnum.botrntComplete;

                        //foreach (var itemx in item.Parti)
                        //{
                        //if (i != 0)
                        //{
                        //
                        //}
                        oGenEntry.Lines.BatchNumbers.Add();
                        oGenEntry.Lines.BatchNumbers.SetCurrentLine(i);
                        oGenEntry.Lines.BatchNumbers.BatchNumber = item.Parti;
                        oGenEntry.Lines.BatchNumbers.Quantity = item.PartiMiktar * -1;
                        i++;
                        //}

                        oGenEntry.Lines.Add();
                    }
                    aa = oGenEntry.Add();
                }

                if (aa == 0)
                {
                    return new Response { Value = 0, Description = "Üretim için giriş başarıyla oluşturuldu.", List = null, DocEntry = Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey()) };
                }
                else
                {
                    string error = "";

                    try
                    {
                        error = ConstVariables.oCompanyObject.GetLastErrorDescription();
                    }
                    catch (Exception)
                    {
                    }
                    return new Response { Value = -4100, Description = "Hata Kodu - 4200 Üretim için giriş oluşturulurken hata oluştu. " + error, List = null };
                }
            }
            catch (Exception ex)
            {
                return new Response { Value = 9000, Description = "Bilinmeyen hata oluştu. " + ex.Message, List = null };
            }
        }

        public Response addPurchaseDeliveryNotes(PurchaseDeliveryNotes purchaseOrder, string irsaliyeNo, string kantarNo)
        {
            try
            {
                SAPbobsCOM.Documents oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);

                oDocuments.BPL_IDAssignedToInvoice = purchaseOrder.BPLId;
                oDocuments.CardCode = purchaseOrder.CardCode;
                DateTime date = DateTime.ParseExact(docDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                oDocuments.DocDate = date;
                oDocuments.DocDueDate = new DateTime(Convert.ToInt32(purchaseOrder.DocDate.Substring(0, 4)) + Convert.ToInt32(purchaseOrder.DocDate.Substring(4, 2)) + Convert.ToInt32(purchaseOrder.DocDate.Substring(6, 2)));
                oDocuments.TaxDate = new DateTime(Convert.ToInt32(purchaseOrder.DocDate.Substring(0, 4)) + Convert.ToInt32(purchaseOrder.DocDate.Substring(4, 2)) + Convert.ToInt32(purchaseOrder.DocDate.Substring(6, 2)));
                oDocuments.UserFields.Fields.Item("U_SutKabulId").Value = purchaseOrder.sutKabulId.ToString();
                oDocuments.UserFields.Fields.Item("U_KullaniciId").Value = purchaseOrder.userId.ToString();
                oDocuments.UserFields.Fields.Item("U_SutKabulIrsNo").Value = irsaliyeNo.ToString();
                oDocuments.UserFields.Fields.Item("U_KantarNo").Value = kantarNo.ToString();
                double toplamSutMiktari = 0;
                int i = 0;
                foreach (var item in purchaseOrder.Lines)
                {
                    oDocuments.Lines.ItemCode = item.ItemCode;
                    oDocuments.Lines.Quantity = item.Quantity;
                    toplamSutMiktari += item.Quantity;

                    if (item.WareHouse != null || item.WareHouse != "")
                    {
                        oDocuments.Lines.WarehouseCode = item.WareHouse;
                    }

                    foreach (var itemPrt in item.PartiDetails)
                    {
                        oDocuments.Lines.BatchNumbers.Add();
                        oDocuments.Lines.BatchNumbers.SetCurrentLine(i);
                        oDocuments.Lines.BatchNumbers.BatchNumber = itemPrt.BatchNumber;
                        oDocuments.Lines.BatchNumbers.Quantity = itemPrt.Quantity;
                        i++;
                    }

                    i = 0;

                    oDocuments.Lines.Add();
                }
                oDocuments.UserFields.Fields.Item("U_TopSutMiktari").Value = toplamSutMiktari.ToString();
                int retval = oDocuments.Add();

                if (retval == 0)
                {
                    return new Response { Value = 0, Description = "Satınalma siparişi başarıyla oluşturuldu.", List = null, DocEntry = Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey()) };
                }
                else
                {
                    return new Response { Value = -1200, Description = "Hata Kodu : -1200 Satınalma siparişi oluşturulurken hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription(), List = null, DocEntry = 0 };
                }
            }
            catch (Exception ex)
            {
                return new Response { Value = 9000, Description = "Bilinmeyen hata oluştu. " + ex.Message, List = null, DocEntry = 0 };
            }
        }
    }
}