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
    public class SAPIade
    {
        [ItemAtt(AIFConn.SAPIade_FormUID)]
        public SAPbouiCOM.Form frmSAPIade;

        private static string formuid = "";
        //private SAPbouiCOM.Button oBtnKaliteKnt;
        //public SAPbouiCOM.Button btnIptal;
        private List<UrunIadeParametre> urunIadeParametres = new List<UrunIadeParametre>();

        public void LoadForms()
        {
            if (Program.mKod == "010OTATURVT")
            {
                Functions.CreateUserOrSystemFormComponent<SAPIade>(AIFConn.Sys180, true, formuid);

                System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
                System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("AIF.UVT.SAPB1.FormsView.SAPIade.xml");

                System.IO.StreamReader streamreader = new System.IO.StreamReader(stream, true);
                xmldoc.LoadXml(string.Format(streamreader.ReadToEnd(), formuid));
                Handler.SAPApplication.LoadBatchActions(xmldoc.InnerXml);

                streamreader.Close();

                var cml = frmSAPIade.GetAsXML();
                InitForms();
            }
        }

        public void InitForms()
        {
            try
            {
                //oBtnKaliteKnt = (SAPbouiCOM.Button)frmSAPIade.Items.Item("btnKltKnt").Specific;
                //btnIptal = (SAPbouiCOM.Button)frmSAPIade.Items.Item("2").Specific;

                //oBtnKaliteKnt.Item.Top = btnIptal.Item.Top;
                //oBtnKaliteKnt.Item.LinkTo = "2";

                #region kalite kontrol buton yerleşimi

                frmSAPIade.Items.Item("btnKltKnt").Top = frmSAPIade.Items.Item("2").Top;
                frmSAPIade.Items.Item("btnKltKnt").Height = frmSAPIade.Items.Item("2").Height;
                frmSAPIade.Items.Item("btnKltKnt").Left = frmSAPIade.Items.Item("2").Left + 85;
                frmSAPIade.Items.Item("btnKltKnt").LinkTo = "2";
                #endregion

                #region taslak belgede kalite kontrol butonu gösterilmesin.
                //string belgeTipi = ((SAPbouiCOM.ComboBox)frmSAPIade.Items.Item("81").Specific).Value.ToString();

                //if (belgeTipi == "6") //taslak
                //{
                //    frmSAPIade.Items.Item("btnKltKnt").Visible = false;
                //} 
                #endregion
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Hata oluştu." + ex.Message);
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
                    //if (!BusinessObjectInfo.BeforeAction)
                    //{
                    //    string xml = frmSatinalmaSiparisi.DataSources.DBDataSources.Item(0).GetAsXML();
                    //    XmlDocument key = new XmlDocument();
                    //    key.LoadXml(BusinessObjectInfo.ObjectKey);
                    //    SASDocEntry = key.SelectNodes("DocumentParams").Item(0).SelectNodes("DocEntry").Item(0).InnerXml;
                    //    eklemeGuncellemeBelgeNo = key.SelectNodes("DocumentParams").Item(0).SelectNodes("DocEntry").Item(0).InnerXml;
                    //    eklemeGuncelleme = true;
                    //}
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
                        if (pVal.ItemUID == "btnKltKnt" && !pVal.BeforeAction)
                        {
                            try
                            {
                                if (frmSAPIade.Mode != BoFormMode.fm_OK_MODE)
                                {
                                    Handler.SAPApplication.MessageBox("Ekleme, güncelleme veya bul modunda işleme devam edilemez.");
                                    return false;
                                }

                                string docEntry = frmSAPIade.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).ToString();
                                string docDate = frmSAPIade.DataSources.DBDataSources.Item(0).GetValue("DocDate", 0).ToString();
                                string cardCode = frmSAPIade.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).ToString();
                                string cardName = frmSAPIade.DataSources.DBDataSources.Item(0).GetValue("CardName", 0).ToString();
                                string shiptoCode = frmSAPIade.DataSources.DBDataSources.Item(0).GetValue("ShipToCode", 0).ToString();

                                if (docEntry != "")
                                {

                                    //DateTime tarih = new DateTime(Convert.ToInt32(docDate.ToString().Substring(0, 4)), Convert.ToInt32(docDate.ToString().Substring(4, 2)), Convert.ToInt32(docDate.ToString().Substring(6, 2)));

                                    urunIadeParametres = new List<UrunIadeParametre>();
                                    urunIadeParametres.Add(new UrunIadeParametre
                                    {
                                        IadeBelgeNo = docEntry,
                                        IadeTarih = docDate,
                                        CariKodu = cardCode,
                                        CariAdi = cardName,
                                        CariDetay = shiptoCode,
                                    });
                                }

                                if (urunIadeParametres != null)
                                {
                                    AIFConn.UrnIadeScm.LoadForms(urunIadeParametres);
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
                        frmSAPIade = Handler.SAPApplication.Forms.Item(pVal.FormUID);
                        //frmInit = false;
                        var xml = frmSAPIade.GetAsXML();
                        formuid = pVal.FormUID;
                        AIFConn.Sys180.LoadForms();
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