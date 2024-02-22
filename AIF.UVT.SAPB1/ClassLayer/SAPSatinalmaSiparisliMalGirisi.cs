using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using AIF.UVT.SAPB1.HelperClass;
using AIF.UVT.SAPB1.Models;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class SAPSatinalmaSiparisliMalGirisi
    {
        [ItemAtt(AIFConn.SAPSatinalmaSiparisliMalGirisi_FormUID)]
        public SAPbouiCOM.Form frmSAPSatinalmaSiparisliMalGirisi;

        private static string formuid = "";
        //private SAPbouiCOM.Button oBtnKaliteKnt;
        //public SAPbouiCOM.Button btnIptal;
        private List<GirdiKontrolParametre> girdiKontrolParametres = new List<GirdiKontrolParametre>();

        public void LoadForms()
        {
            if (Program.mKod == "010OTATURVT")
            {
                Functions.CreateUserOrSystemFormComponent<SAPSatinalmaSiparisliMalGirisi>(AIFConn.Sys143, true, formuid);

                System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
                System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("AIF.UVT.SAPB1.FormsView.SAPSatinalmaSiparisliMalGirisi.xml");

                System.IO.StreamReader streamreader = new System.IO.StreamReader(stream, true);
                xmldoc.LoadXml(string.Format(streamreader.ReadToEnd(), formuid));
                Handler.SAPApplication.LoadBatchActions(xmldoc.InnerXml);

                streamreader.Close();

                var cml = frmSAPSatinalmaSiparisliMalGirisi.GetAsXML();
                InitForms();
            }
        }

        public void InitForms()
        {
            try
            {
                //oBtnKaliteKnt = (SAPbouiCOM.Button)frmSAPSatinalmaSiparisliMalGirisi.Items.Item("btnKaliteKnt").Specific;
                //btnIptal = (SAPbouiCOM.Button)frmSAPSatinalmaSiparisliMalGirisi.Items.Item("2").Specific;

                //oBtnKaliteKnt.Item.Top = btnIptal.Item.Top;
                //oBtnKaliteKnt.Item.LinkTo = "2";

                #region kalite kontrol buton yerleşimi
                frmSAPSatinalmaSiparisliMalGirisi.Items.Item("btnKaltKnt").Top = frmSAPSatinalmaSiparisliMalGirisi.Items.Item("2").Top;
                //frmSAPSatinalmaSiparisliMalGirisi.Items.Item("btnKaltKnt").Height = frmSAPSatinalmaSiparisliMalGirisi.Items.Item("2").Height;
                frmSAPSatinalmaSiparisliMalGirisi.Items.Item("btnKaltKnt").Left = frmSAPSatinalmaSiparisliMalGirisi.Items.Item("2").Left + 85;
                frmSAPSatinalmaSiparisliMalGirisi.Items.Item("btnKaltKnt").LinkTo = "2";
                #endregion

                #region gerçek belgede kalite kontrol butonu gösterilmesin.
                string belgeTipi = ((SAPbouiCOM.ComboBox)frmSAPSatinalmaSiparisliMalGirisi.Items.Item("81").Specific).Value.ToString();

                if (belgeTipi != "6") //taslak
                {
                    frmSAPSatinalmaSiparisliMalGirisi.Items.Item("btnKaltKnt").Visible = false;
                }
                #endregion
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Form yüklenirken oluştu." + ex.Message);
            }
        }
        bool ekleme = false;
        string SASGercekDocEntry = "";
        string SASTaslakDocEntry = "";
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
                    if (Program.mKod == "010OTATURVT")
                    {
                        if (!BusinessObjectInfo.BeforeAction)
                        {
                            try
                            {
                                string belgeTipi = ((SAPbouiCOM.ComboBox)frmSAPSatinalmaSiparisliMalGirisi.Items.Item("81").Specific).Value.ToString();

                                if (belgeTipi == "6")
                                {
                                    string xml = frmSAPSatinalmaSiparisliMalGirisi.DataSources.DBDataSources.Item(0).GetAsXML();
                                    XmlDocument key = new XmlDocument();
                                    key.LoadXml(BusinessObjectInfo.ObjectKey);
                                    SASGercekDocEntry = key.SelectNodes("DocumentParams").Item(0).SelectNodes("DocEntry").Item(0).InnerXml;

                                    //string sql = "Select * from ODRF ";

                                    SAPbouiCOM.Form oForm = null;
                                    //SAPbouiCOM.DBDataSource odrf = null;

                                    oForm = Handler.SAPApplication.Forms.ActiveForm;

                                    //odrf = oForm.DataSources.DBDataSources.Item("ODRF");

                                    if (oForm.DataSources.DBDataSources.Item(0).TableName == "ODRF")
                                    {

                                    }
                                    else if (oForm.DataSources.DBDataSources.Item(0).TableName == "OPDN")
                                    {

                                    }

                                    ekleme = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                Handler.SAPApplication.MessageBox("Belge ekleme sırasında hata oluştu." + ex.Message);
                            }
                        }
                    }
                    break;

                case BoEventTypes.et_FORM_DATA_UPDATE:
                    //if (!BusinessObjectInfo.BeforeAction)
                    //{
                    //    eklemeGuncelleme = true;
                    //    SASDocEntry = frmSatinalmaSiparisi.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).ToString();
                    //    eklemeGuncellemeBelgeNo = frmSatinalmaSiparisi.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).ToString();
                    //}
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
                    if (Program.mKod == "010OTATURVT")
                    {
                        if (pVal.ItemUID == "1" && pVal.BeforeAction)
                        {
                            try
                            {
                                SASTaslakDocEntry = frmSAPSatinalmaSiparisliMalGirisi.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).ToString();
                            }
                            catch (Exception ex)
                            {
                                Handler.SAPApplication.MessageBox("Taslak belge numarası alınırken hata oluştu." + ex.Message);
                            }

                        }
                        else if (pVal.ItemUID == "1" && !pVal.BeforeAction && ekleme && SASGercekDocEntry != "")
                        {
                            #region Taslak belgeden gerçek belgeye çevirilme yapılırken GIRDIKONTROL formundaki taslak iken oluşan kayıt güncellenir.
                            try
                            {
                                SAPbobsCOM.GeneralService oGeneralService = null;

                                SAPbobsCOM.GeneralData oGeneralData = null;

                                SAPbobsCOM.GeneralDataParams oGeneralParams = null;

                                SAPbobsCOM.CompanyService sCmp = null;

                                SAPbobsCOM.GeneralData oChild = null;

                                SAPbobsCOM.GeneralDataCollection oChildren = null;

                                sCmp = ConstVariables.oCompanyObject.GetCompanyService();

                                oGeneralService = sCmp.GetGeneralService("AIF_GIRDIKONTROL");

                                oGeneralParams = ((SAPbobsCOM.GeneralDataParams)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)));

                                string sql = "Select \"DocEntry\" from \"@AIF_GIRDIKONTROL\" where \"U_TaslakSipNo\" = '" + SASTaslakDocEntry + "' AND \"U_BelgeTipi\" = 'T' ";
                                ConstVariables.oRecordset.DoQuery(sql);

                                if (ConstVariables.oRecordset.RecordCount > 0)
                                {
                                    oGeneralParams.SetProperty("DocEntry", Convert.ToInt32(ConstVariables.oRecordset.Fields.Item(0).Value));

                                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);


                                    oGeneralData.SetProperty("U_GercekSipNo", SASGercekDocEntry.ToString());

                                    oGeneralData.SetProperty("U_BelgeTipi", "G"); //gerçek

                                    try
                                    {
                                        oGeneralService.Update(oGeneralData);
                                    }
                                    catch (Exception ex)
                                    {
                                        Handler.SAPApplication.MessageBox("Girdi Kontrol tablosu güncellenirken hata oluştu." + ex.Message);
                                    }

                                    #region satır güncelleme
                                    //oChildren = oGeneralData.Child("AIF_EKSTREAKTARIMI1");

                                    ////oChildren.Add();

                                    ////oChildren.Item(_lineId - 1).SetProperty("U_YevmiyeNo", yevmiyeNo);

                                    //oChild = oChildren.Item(_sira - 1);

                                    //oChild.SetProperty("U_YevmiyeNo", yevmiyeNo);
                                    //oChild.SetProperty("U_Durum", "Yevmiye kaydı oluşturulmuştur");
                                    //oGeneralService.Update(oGeneralData); 
                                    #endregion
                                }
                            }
                            catch (Exception ex)
                            {
                                Handler.SAPApplication.MessageBox("Hata oluştu." + ex.Message);
                            }
                            finally
                            {
                                SASGercekDocEntry = "";
                                SASTaslakDocEntry = "";
                                ekleme = false;
                            }
                            #endregion
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
                    if (Program.mKod == "010OTATURVT")
                    {
                        if (pVal.ItemUID == "btnKaltKnt" && !pVal.BeforeAction)
                        {
                            try
                            {
                                string belgeTipi = ((SAPbouiCOM.ComboBox)frmSAPSatinalmaSiparisliMalGirisi.Items.Item("81").Specific).Value.ToString();
                                string docEntry = frmSAPSatinalmaSiparisliMalGirisi.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).ToString();

                                if (belgeTipi != "6") //taslak
                                {
                                    if (frmSAPSatinalmaSiparisliMalGirisi.Mode != BoFormMode.fm_OK_MODE)
                                    {
                                        Handler.SAPApplication.MessageBox("Ekleme, güncelleme veya bul modunda işleme devam edilemez.");
                                        return false;
                                    }

                                }

                                if (belgeTipi != "" && docEntry != "")
                                {
                                    girdiKontrolParametres = new List<GirdiKontrolParametre>();
                                    girdiKontrolParametres.Add(new GirdiKontrolParametre
                                    {
                                        SiparisNumarasi = docEntry,
                                        TedarikciKodu = ((SAPbouiCOM.EditText)frmSAPSatinalmaSiparisliMalGirisi.Items.Item("4").Specific).Value.ToString(),
                                        TedarikciAdi = ((SAPbouiCOM.EditText)frmSAPSatinalmaSiparisliMalGirisi.Items.Item("54").Specific).Value.ToString(),
                                        IrsaliyeNo = ((SAPbouiCOM.EditText)frmSAPSatinalmaSiparisliMalGirisi.Items.Item("14").Specific).Value.ToString(),
                                        BelgeTipi = belgeTipi,
                                    });
                                }

                                if (girdiKontrolParametres != null)
                                {
                                    AIFConn.GirdiKntrl.LoadForms(girdiKontrolParametres);
                                }
                            }
                            catch (Exception ex)
                            {
                                Handler.SAPApplication.MessageBox("Hata oluştu." + ex.Message);
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
                    if (pVal.BeforeAction)
                    {
                        frmSAPSatinalmaSiparisliMalGirisi = Handler.SAPApplication.Forms.Item(pVal.FormUID);
                        //frmInit = false;
                        var xml = frmSAPSatinalmaSiparisliMalGirisi.GetAsXML();
                        formuid = pVal.FormUID;
                        AIFConn.Sys143.LoadForms();
                    }
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
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }
    }
}