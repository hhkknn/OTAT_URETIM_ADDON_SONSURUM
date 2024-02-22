using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using AIF.UVT.SAPB1.HelperClass;
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
    public class SatinalmaSiparisi
    {
        [ItemAtt(AIFConn.SatinalmaSiparisi_FormUID)]
        public SAPbouiCOM.Form frmSatinalmaSiparisi;

        private static string formuid = "";
        private SAPbouiCOM.EditText oEdtCardCode = null;
        private SAPbouiCOM.EditText oEdtDocDate = null;
        private SAPbouiCOM.Matrix oMatrix = null;

        public void LoadForms()
        {
            if (Program.mKod == "010OTATURVT")
            {
                Functions.CreateUserOrSystemFormComponent<SatinalmaSiparisi>(AIFConn.Sys142, true, formuid);

                System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
                System.IO.Stream stream = null;

                stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("AIF.UVT.SAPB1.FormsView.SatinalmaSiparisi.xml");

                #region ÇOKLU EKRAN ÇALIŞMASI İPTAL EDİLDİ
                //if (Program.mKod == "010OTATURVT")
                //{
                //    stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("AIF.UVT.SAPB1.FormsView.SatinalmaSiparisi_Otat.xml");
                //}
                //else
                //{
                //    stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("AIF.UVT.SAPB1.FormsView.SatinalmaSiparisi_Yoruk.xml");
                //} 
                #endregion

                System.IO.StreamReader streamreader = new System.IO.StreamReader(stream, true);
                xmldoc.LoadXml(string.Format(streamreader.ReadToEnd(), formuid));
                Handler.SAPApplication.LoadBatchActions(xmldoc.InnerXml);

                streamreader.Close();

                var cml = frmSatinalmaSiparisi.GetAsXML();
                InitForms();
            }
        }

        public void InitForms()
        {
            try
            {
                frmSatinalmaSiparisi.Freeze(true);

                if (Program.mKod == "010OTATURVT")
                {
                    //frmSatinalmaSiparisi.Items.Item("Item_5").Top = frmSatinalmaSiparisi.Items.Item("2").Top - 20;
                    //frmSatinalmaSiparisi.Items.Item("Item_5").Left = frmSatinalmaSiparisi.Items.Item("16").Left;
                    //frmSatinalmaSiparisi.Items.Item("Item_6").Top = frmSatinalmaSiparisi.Items.Item("Item_5").Top;

                    oEdtCardCode = (SAPbouiCOM.EditText)frmSatinalmaSiparisi.Items.Item("4").Specific;
                    oEdtDocDate = (SAPbouiCOM.EditText)frmSatinalmaSiparisi.Items.Item("10").Specific;
                    oMatrix = (SAPbouiCOM.Matrix)frmSatinalmaSiparisi.Items.Item("38").Specific;

                    frmSatinalmaSiparisi.Items.Item("Item_4").Top = frmSatinalmaSiparisi.Items.Item("2").Top;
                    frmSatinalmaSiparisi.Items.Item("Item_4").Left = frmSatinalmaSiparisi.Items.Item("2").Left + 85;
                    frmSatinalmaSiparisi.Items.Item("Item_4").Width = 150;
                    frmSatinalmaSiparisi.Items.Item("Item_4").LinkTo = "2";


                    frmSatinalmaSiparisi.Items.Item("Item_5").Top = frmSatinalmaSiparisi.Items.Item("Item_4").Top;
                    frmSatinalmaSiparisi.Items.Item("Item_5").Left = frmSatinalmaSiparisi.Items.Item("Item_4").Left + frmSatinalmaSiparisi.Items.Item("Item_4").Width + 5;
                    frmSatinalmaSiparisi.Items.Item("Item_5").Width = 120;
                    frmSatinalmaSiparisi.Items.Item("Item_5").LinkTo = "Item_4";

                    if (SatisSiparisi.docEntrySatisSiparisi != "")
                    {
                        ((SAPbouiCOM.EditText)frmSatinalmaSiparisi.Items.Item("Item_3").Specific).Value = SatisSiparisi.docEntrySatisSiparisi;
                        SAPbobsCOM.Documents oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

                        oDocuments.GetByKey(Convert.ToInt32(SatisSiparisi.docEntrySatisSiparisi));

                        ((SAPbouiCOM.EditText)frmSatinalmaSiparisi.Items.Item("4").Specific).Value = SatisSiparisi.cardCodeSatisSiparisi.ToString();

                        for (int i = 0; i <= oDocuments.Lines.Count - 1; i++)
                        {
                            oDocuments.Lines.SetCurrentLine(i);
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(oMatrix.RowCount).Specific).Value = oDocuments.Lines.ItemCode;
                            //((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(oMatrix.RowCount - 1).Specific).Value = oDocuments.Lines.DiscountPercent.ToString();
                            //((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Iskonto1").Cells.Item(oMatrix.RowCount - 1).Specific).Value = oDocuments.Lines.UserFields.Fields.Item("U_Iskonto1").Value.ToString();
                            //((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Iskonto2").Cells.Item(oMatrix.RowCount - 1).Specific).Value = oDocuments.Lines.UserFields.Fields.Item("U_Iskonto2").Value.ToString();
                            //((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Iskonto3").Cells.Item(oMatrix.RowCount - 1).Specific).Value = oDocuments.Lines.UserFields.Fields.Item("U_Iskonto3").Value.ToString();
                            //((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Iskonto4").Cells.Item(oMatrix.RowCount - 1).Specific).Value = oDocuments.Lines.UserFields.Fields.Item("U_Iskonto4").Value.ToString();
                            //((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Iskonto5").Cells.Item(oMatrix.RowCount - 1).Specific).Value = oDocuments.Lines.UserFields.Fields.Item("U_Iskonto5").Value.ToString();
                            //((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SSDocEntry").Cells.Item(oMatrix.RowCount - 1).Specific).Value = oDocuments.DocEntry.ToString();
                            //((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SSLineNum").Cells.Item(oMatrix.RowCount - 1).Specific).Value = oDocuments.Lines.LineNum.ToString();
                            //((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_SSDocType").Cells.Item(oMatrix.RowCount - 1).Specific).Value = "17";
                        }

                        oMatrix.Columns.Item("1").Cells.Item(1).Click();

                        //SatisSiparisi.docEntrySatisSiparisi = "";
                    }
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
                frmSatinalmaSiparisi.Freeze(false);
            }
        }

        private bool eklemeGuncelleme = false;
        private string eklemeGuncellemeBelgeNo;

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
                        if (Program.mKod == "010OTATURVT")
                        {
                            string xml = frmSatinalmaSiparisi.DataSources.DBDataSources.Item(0).GetAsXML();
                            XmlDocument key = new XmlDocument();
                            key.LoadXml(BusinessObjectInfo.ObjectKey);
                            SASDocEntry = key.SelectNodes("DocumentParams").Item(0).SelectNodes("DocEntry").Item(0).InnerXml;
                            eklemeGuncellemeBelgeNo = key.SelectNodes("DocumentParams").Item(0).SelectNodes("DocEntry").Item(0).InnerXml;
                            eklemeGuncelleme = true;
                        }
                    }
                    break;

                case BoEventTypes.et_FORM_DATA_UPDATE:
                    if (!BusinessObjectInfo.BeforeAction)
                    {
                        if (Program.mKod == "010OTATURVT")
                        {
                            eklemeGuncelleme = true;
                            SASDocEntry = frmSatinalmaSiparisi.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).ToString();
                            eklemeGuncellemeBelgeNo = frmSatinalmaSiparisi.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).ToString();
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

        private string SASDocEntry = "";
        private string oncekiIndirimComboDegeri = "";
        private string val = "";

        private void CreateAttachment(string docNum)
        {
            #region Crystal reports işlemlerinin yapıldığı yer

            if (Program.mKod == "010OTATURVT")
            {
                try
                {
                    string ServerPass = "";

                    string path = "";

                    path = System.AppDomain.CurrentDomain.BaseDirectory + "PDF\\YemCiktisi.rpt";


                    //if (ConstVariables.oCompanyObject.CompanyDB.ToString().ToUpper().Contains("TESTOTAT"))
                    //{
                    //    path= System.AppDomain.CurrentDomain.BaseDirectory + "PDF\\YemCiktisi.rpt"; 
                    //}
                    //else
                    //{
                    //    path = System.AppDomain.CurrentDomain.BaseDirectory + "PDF\\YemCiktisi.rpt";
                    //}

                    //string path = HttpContext.Current.Request.PhysicalApplicationPath + "bin\\PDF\\KargoFisi.rpt";

                    ReportDocument cryRpt = new ReportDocument();
                    cryRpt.Load(path);

                    cryRpt.SetDatabaseLogon("sa", "Qaz1Wsx2", "SAPSRV", ConstVariables.oCompanyObject.CompanyDB.ToString());

                    cryRpt.VerifyDatabase();

                    cryRpt.SetParameterValue(0, docNum);

                    CrystalDecisions.Shared.ExportOptions CrExportOptions;
                    DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                    PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();

                    CrDiskFileDestinationOptions.DiskFileName = System.AppDomain.CurrentDomain.BaseDirectory + "PDF\\" + docNum + ".pdf";
                    //CrDiskFileDestinationOptions.DiskFileName = path + "bin\\PDF\\" + docNum + ".pdf";
                    CrExportOptions = cryRpt.ExportOptions;
                    {
                        CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                        CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                        CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                        CrExportOptions.FormatOptions = CrFormatTypeOptions;
                    }
                    cryRpt.Export();
                    cryRpt.Close();
                }
                catch (Exception x)
                {
                }
            }

            #endregion Crystal reports işlemlerinin yapıldığı yer
        }

        private bool mailGonder(SAPbobsCOM.Documents _oDocuments)
        {

            try
            {
                if (Program.mKod == "010OTATURVT")
                {
                    string body = string.Empty;
                    using (StreamReader reader = new StreamReader(System.Windows.Forms.Application.StartupPath + "\\PDF\\mail.html"))
                    {
                        body = reader.ReadToEnd();
                    }
                    SASDocEntry = frmSatinalmaSiparisi.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).ToString();
                    //string kimden = "support@aifteam.com";
                    //string sifre = "Ncckr1989.";

                    string kimden = "bilgilendirme@otat.com.tr";
                    string sifre = "Otat543654";

                    //string kimden = "fatih.gulen@aifteam.com";
                    //string sifre = "Ncckr1989.";

                    SmtpClient SMTP = new SmtpClient();
                    MailMessage Mail = new MailMessage();

                    var durum = ((SAPbouiCOM.ComboBox)frmSatinalmaSiparisi.Items.Item("81").Specific).Value.Trim();

                    string sql = "";

                    if (durum == "6")
                    {
                        sql = "Select T1.\"DocEntry\" as \"SipNo\",T1.\"CardName\" as \"Musteri\",T0.\"DocDate\" as \"Tarih\" from \"DRF1\" as T0 INNER JOIN \"ORDR\" as T1 ON T0.\"DocNum\" = T1.\"DocEntry\" where T0.\"DocEntry\" = '" + _oDocuments.DocEntry + "' and \"ObjType\" = '17' ";
                    }
                    else
                    {
                        sql = "Select T1.\"DocEntry\" as \"SipNo\",T1.\"CardName\" as \"Musteri\",T0.\"DocDate\" as \"Tarih\" from \"POR1\" as T0 INNER JOIN \"ORDR\" as T1 ON T0.\"BaseEntry\" = T1.\"DocEntry\" where T0.\"DocEntry\" = '" + _oDocuments.DocEntry + "' ";
                    }

                    ConstVariables.oRecordset.DoQuery(sql);
                    Mail.From = new MailAddress(kimden);
                    Mail.Subject = "Yem Sipariş Onayı / Sip.No: " + ConstVariables.oRecordset.Fields.Item("SipNo").Value.ToString() + " / Müşteri: " + ConstVariables.oRecordset.Fields.Item("Musteri").Value.ToString() + "";
                    Mail.IsBodyHtml = true;

                    CreateAttachment(SASDocEntry.ToString());
                    Attachment attachment = new Attachment(System.Windows.Forms.Application.StartupPath + "\\PDF\\" + SASDocEntry + ".pdf");
                    Mail.Attachments.Add(attachment);

                    string anayazi = "" + Convert.ToDateTime(ConstVariables.oRecordset.Fields.Item("Tarih").Value).ToString("dd/MM/yyyy") + " siparişli tarih " + ConstVariables.oRecordset.Fields.Item("SipNo").Value.ToString() + " numaralı " + _oDocuments.DocTotal.ToString("#,#0.00") + " liralık " + ConstVariables.oRecordset.Fields.Item("Musteri").Value.ToString() + " ait siparişi için onayınız beklenmektedir.";

                    string hostname = "http://172.55.10.16:3100" + "//Home//Index?docEntry=" + SASDocEntry.ToString();
                    //ConstVariables.oRecordset.Fields.Item("U_HostAdresi").Value.ToString() + "\\Home\\Index?guid=" + guid;
                    body = body.Replace("{yazi}", anayazi);
                    body = body.Replace("{onaylink}", hostname + "&onayTipi=1");
                    body = body.Replace("{redlink}", hostname + "&onayTipi=2");
                    body = body.Replace("{Yil}", DateTime.Now.Year.ToString());

                    Mail.To.Add("yemsiparis@otat.com.tr");
                    Mail.To.Add("murat.yagmur@otat.com.tr");
                    //Mail.Bcc.Add("support@aifteam.com");

                    //Mail.CC.Add("kadir.degirmenci@aifteam.com");
                    //Mail.CC.Add("fatih.gulen@aifteam.com"); 

                    SMTP.Host = "mail.otat.com.tr";
                    //SMTP.Host = "smtp.outlook.com";
                    SMTP.DeliveryMethod = SmtpDeliveryMethod.Network;
                    SMTP.EnableSsl = true;
                    SMTP.UseDefaultCredentials = false;
                    SMTP.Credentials = new System.Net.NetworkCredential(kimden, sifre);
                    SMTP.Port = 25;
                    //SMTP.Port = 587;

                    Mail.Body = body;

                    SMTP.Send(Mail);

                    Handler.SAPApplication.MessageBox("Onay maili gönderilmiştir.");
                }
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Mail gönderirken hata oluştu." + ex.Message);
                return false;
            }

            return true;
            //var retval = oDocuments.Update();

            //if (retval == 0)
            //{
            //}
            //else
            //{
            //    var tst = ConstVariables.oCompanyObject.GetLastErrorDescription();
            //    Handler.SAPApplication.MessageBox("Hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription());
            //} 

        }

        private bool yenidencomboseciliyor = false;

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
                        if (!pVal.BeforeAction && eklemeGuncelleme)
                        {
                            if (eklemeGuncellemeBelgeNo != "")
                            {
                                SAPbobsCOM.Documents oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

                                var exist = oDocuments.GetByKey(Convert.ToInt32(eklemeGuncellemeBelgeNo));

                                if (exist)
                                {
                                    //if (oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value.ToString() != "" && oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value.ToString() != null && oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value.ToString() != "0" && oDocuments.Lines.UserFields.Fields.Item("U_SSLineNum").Value.ToString() != "" && oDocuments.Lines.UserFields.Fields.Item("U_SSLineNum").Value.ToString() != null && oDocuments.Lines.UserFields.Fields.Item("U_SSLineNum").Value.ToString() != "0" && oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value.ToString() != "" && oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value.ToString() != null && oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value.ToString() != "0")
                                    if (oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value.ToString() != "" && oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value.ToString() != null && oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value.ToString() != "0" && oDocuments.Lines.UserFields.Fields.Item("U_SSLineNum").Value.ToString() != "" && oDocuments.Lines.UserFields.Fields.Item("U_SSLineNum").Value.ToString() != null && oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value.ToString() != "" && oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value.ToString() != null && oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value.ToString() != "0")
                                    {
                                        string updatesql = "";
                                        for (int i = 0; i <= oDocuments.Lines.Count - 1; i++)
                                        {
                                            oDocuments.Lines.SetCurrentLine(i);

                                            if (oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value.ToString() != "")
                                            {
                                                updatesql = "UPDATE POR1 SET \"BaseEntry\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value + "' where \"DocEntry\" = '" + SASDocEntry + "' and \"LineNum\" = '" + i + "' ";

                                                ConstVariables.oRecordset.DoQuery(updatesql);
                                            }

                                            if (oDocuments.Lines.UserFields.Fields.Item("U_SSLineNum").Value.ToString() != "")
                                            {
                                                updatesql = "UPDATE POR1 SET \"BaseLine\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSLineNum").Value + "' where \"DocEntry\" = '" + SASDocEntry + "' and \"LineNum\" = '" + i + "' ";

                                                ConstVariables.oRecordset.DoQuery(updatesql);
                                            }

                                            if (oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value.ToString() != "")
                                            {
                                                updatesql = "UPDATE POR1 SET \"BaseType\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value + "' where \"DocEntry\" = '" + SASDocEntry + "' and \"LineNum\" = '" + i + "' ";

                                                ConstVariables.oRecordset.DoQuery(updatesql);
                                            }



                                            //if (oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value.ToString() != "")
                                            //{
                                            //    updatesql = "UPDATE POR1 SET \"BaseEntry\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value + "' where \"DocEntry\" = '" + SASDocEntry + "' and \"LineNum\" = '" + i + "' ";

                                            //    ConstVariables.oRecordset.DoQuery(updatesql);
                                            //}

                                            //if (oDocuments.Lines.UserFields.Fields.Item("U_SSLineNum").Value.ToString() != "")
                                            //{
                                            //    updatesql = "UPDATE POR1 SET \"BaseLine\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSLineNum").Value + "' where \"DocEntry\" = '" + SASDocEntry + "' and \"LineNum\" = '" + i + "' ";

                                            //    ConstVariables.oRecordset.DoQuery(updatesql);
                                            //}

                                            //if (oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value.ToString() != "")
                                            //{
                                            //    updatesql = "UPDATE POR1 SET \"BaseType\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value + "' where \"DocEntry\" = '" + SASDocEntry + "' and \"LineNum\" = '" + i + "' ";

                                            //    ConstVariables.oRecordset.DoQuery(updatesql);
                                            //}
                                        }

                                        #region for içindeydi.break olduğu için for da dönmeyi bıraktığından ilk satırda basetype, baseentry, baseline ı yazıp diğerlerini yazamıyor
                                        //ConstVariables.oRecordset.DoQuery("Select * from \"POR1\" where \"BaseEntry\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value + "'");

                                        //if (ConstVariables.oRecordset.RecordCount > 0)
                                        //{
                                        //    break;
                                        //}

                                        //ConstVariables.oRecordset.DoQuery("Select * from \"POR1\" where \"BaseLine\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSLineNum").Value + "'");

                                        //if (ConstVariables.oRecordset.RecordCount > 0)
                                        //{
                                        //    break;
                                        //}

                                        //ConstVariables.oRecordset.DoQuery("Select * from \"POR1\" where \"BaseType\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value + "'");

                                        //if (ConstVariables.oRecordset.RecordCount > 0)
                                        //{
                                        //    break;
                                        //}

                                        //ConstVariables.oRecordset.DoQuery("Select * from \"POR1\" where \"U_SSDocEntry\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value + "'");

                                        //if (ConstVariables.oRecordset.RecordCount > 0)
                                        //{
                                        //    break;
                                        //}

                                        //ConstVariables.oRecordset.DoQuery("Select * from \"POR1\" where \"U_SSLineNum\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSLineNum").Value + "'");

                                        //if (ConstVariables.oRecordset.RecordCount > 0)
                                        //{
                                        //    break;
                                        //}

                                        //ConstVariables.oRecordset.DoQuery("Select * from \"POR1\" where \"U_SSDocType\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value + "'");

                                        //if (ConstVariables.oRecordset.RecordCount > 0)
                                        //{
                                        //    break;
                                        //}
                                        #endregion
                                    }

                                    oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
                                    oDocuments.GetByKey(Convert.ToInt32(SASDocEntry));

                                    if (oDocuments.UserFields.Fields.Item("U_IndirimTipi").Value.ToString() == "1")
                                    {
                                        string sql = "Select T2.\"U_UrunKodu\",T2.\"U_BirinciIsk\",T2.\"U_IkinciIsk\",T2.\"U_UcuncuIsk\",T2.\"U_DorduncuIsk\",T2.\"U_BesinciIsk\",T2.\"U_ToplamIsk\" from \"@AIF_SATINALMAISK\" as T0 INNER JOIN \"@AIF_SATINALMAISK1\" as T2 ON T0.\"DocEntry\" = T2.\"DocEntry\" where '" + oDocuments.DocDate.ToString("yyyyMMdd") + "' >= T0.\"U_BaslangicTarihi\" and '" + oDocuments.DocDate.ToString("yyyyMMdd") + "' <= T0.\"U_BitisTarihi\" and T0.\"U_Durum\" = '1' and T0.\"U_SaticiKodu\" = '" + oDocuments.CardCode + "' ";

                                        try
                                        {
                                            ConstVariables.oRecordset.DoQuery(sql);

                                            string xml = ConstVariables.oRecordset.GetAsXML();

                                            XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                                            XDocument xDoc = XDocument.Parse(ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData));

                                            var rows = (from t in xDoc.Descendants(ns + "Row")
                                                        select new
                                                        {
                                                            UrunKodu = (from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_UrunKodu" select new XElement(k.Element(ns + "Value"))).First().Value,
                                                            BirinciIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_BirinciIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                            IkinciIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_IkinciIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                            UcuncuIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_UcuncuIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                            DorduncuIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_DorduncuIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                            BesinciIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_BesinciIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                            ToplamIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_ToplamIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                        }).ToList();

                                            if (exist)
                                            {
                                                bool satirVar = false;
                                                for (int i = 0; i <= oDocuments.Lines.Count - 1; i++)
                                                {
                                                    oDocuments.Lines.SetCurrentLine(i);
                                                    var first = rows.Where(x => x.UrunKodu == oDocuments.Lines.ItemCode).FirstOrDefault();

                                                    if (first != null)
                                                    {
                                                        satirVar = true;
                                                        oDocuments.Lines.UserFields.Fields.Item("U_Iskonto1").Value = first.BirinciIskonto.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                                        oDocuments.Lines.UserFields.Fields.Item("U_Iskonto2").Value = first.IkinciIskonto.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                                        oDocuments.Lines.UserFields.Fields.Item("U_Iskonto3").Value = first.UcuncuIskonto.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                                        oDocuments.Lines.UserFields.Fields.Item("U_Iskonto4").Value = first.DorduncuIskonto.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                                        oDocuments.Lines.UserFields.Fields.Item("U_Iskonto5").Value = first.BesinciIskonto.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                                        oDocuments.Lines.UserFields.Fields.Item("DiscPrcnt").Value = first.ToplamIskonto.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                                    }
                                                }

                                                if (satirVar)
                                                {
                                                    var retval = oDocuments.Update();

                                                    if (retval == 0)
                                                    {
                                                        oDocuments.GetByKey(Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey()));
                                                        //bool mailgonderildi = mailGonder(oDocuments);

                                                        if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_OK_MODE)
                                                        {
                                                            Handler.SAPApplication.ActivateMenuItem("1304");
                                                        }
                                                        else if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                                                        {
                                                            try
                                                            {
                                                                frmSatinalmaSiparisi.Close();
                                                            }
                                                            catch (Exception)
                                                            {
                                                            }
                                                            var DocEntry = ConstVariables.oCompanyObject.GetNewObjectKey();
                                                            Handler.SAPApplication.OpenForm((BoFormObjectEnum)22, "", DocEntry);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        Handler.SAPApplication.MessageBox("Indirimler uygulanırken hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription());
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Handler.SAPApplication.MessageBox("İndirim güncelleme işleminde hata oluştu." + ex.Message);
                                        }
                                        finally
                                        {
                                            eklemeGuncelleme = false;
                                            eklemeGuncellemeBelgeNo = "";
                                        }
                                    }
                                    else if (oDocuments.UserFields.Fields.Item("U_IndirimTipi").Value.ToString() == "2")
                                    {
                                        try
                                        {
                                            string sql = "Select T2.\"ItemCode\",T2.\"LineNum\",T2.\"U_Iskonto1\",T2.\"U_Iskonto2\",T2.\"U_Iskonto3\",T2.\"U_Iskonto4\",T2.\"U_Iskonto5\" from POR1 as T2 where T2.\"DocEntry\" = '" + eklemeGuncellemeBelgeNo + "'";

                                            ConstVariables.oRecordset.DoQuery(sql);

                                            string xml = ConstVariables.oRecordset.GetAsXML();

                                            XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                                            XDocument xDoc = XDocument.Parse(ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData));

                                            var rows = (from t in xDoc.Descendants(ns + "Row")
                                                        select new
                                                        {
                                                            UrunKodu = (from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "ItemCode" select new XElement(k.Element(ns + "Value"))).First().Value,
                                                            Iskonto1 = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_Iskonto1" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                            Iskonto2 = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_Iskonto2" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                            Iskonto3 = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_Iskonto3" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                            Iskonto4 = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_Iskonto4" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                            Iskonto5 = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_Iskonto5" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                            SatirNumarasi = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "LineNum" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                        }).ToList();

                                            bool satirVarmi = false;
                                            double _toplamIsk = 0;
                                            foreach (var item in rows.Where(x => x.UrunKodu != ""))
                                            {
                                                _toplamIsk = 100 - (100 * (1 - (item.Iskonto1 / 100)) * (1 - (item.Iskonto2 / 100)) * (1 - (item.Iskonto3 / 100)) * (1 - (item.Iskonto4 / 100)) * (1 - (item.Iskonto5 / 100)));

                                                oDocuments.Lines.SetCurrentLine(Convert.ToInt32(item.SatirNumarasi));

                                                oDocuments.Lines.DiscountPercent = _toplamIsk;

                                                satirVarmi = true;
                                                //((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(item.SatirNumarasi).Specific).Value = _toplamIsk.ToString();

                                                //Progress += 1;
                                                //oProgressBar.Value = Progress;
                                            }

                                            if (satirVarmi)
                                            {
                                                int retval = oDocuments.Update();

                                                if (retval == 0)
                                                {
                                                    oDocuments.GetByKey(Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey()));
                                                    bool mailgonderildi = mailGonder(oDocuments);
                                                    if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_OK_MODE)
                                                    {
                                                        Handler.SAPApplication.ActivateMenuItem("1304");
                                                    }
                                                    else if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                                                    {
                                                        try
                                                        {
                                                            frmSatinalmaSiparisi.Close();
                                                        }
                                                        catch (Exception)
                                                        {
                                                        }
                                                        var DocEntry = ConstVariables.oCompanyObject.GetNewObjectKey();
                                                        Handler.SAPApplication.OpenForm((BoFormObjectEnum)22, "", DocEntry);
                                                    }
                                                }
                                                else
                                                {
                                                    Handler.SAPApplication.MessageBox("Indirimler uygulanırken hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription());
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Handler.SAPApplication.MessageBox("İndirim güncelleme işleminde hata oluştu." + ex.Message);
                                        }
                                        finally
                                        {
                                            eklemeGuncelleme = false;
                                            eklemeGuncellemeBelgeNo = "";
                                        }
                                    }
                                    else if (oDocuments.UserFields.Fields.Item("U_IndirimTipi").Value.ToString() == "3")
                                    {
                                        try
                                        {
                                            bool satirVar = false;
                                            for (int i = 0; i <= oDocuments.Lines.Count - 1; i++)
                                            {
                                                satirVar = true;
                                                oDocuments.Lines.SetCurrentLine(i);

                                                oDocuments.Lines.DiscountPercent = 0;
                                                oDocuments.Lines.UserFields.Fields.Item("U_Iskonto1").Value = 0;
                                                oDocuments.Lines.UserFields.Fields.Item("U_Iskonto2").Value = 0;
                                                oDocuments.Lines.UserFields.Fields.Item("U_Iskonto3").Value = 0;
                                                oDocuments.Lines.UserFields.Fields.Item("U_Iskonto4").Value = 0;
                                                oDocuments.Lines.UserFields.Fields.Item("U_Iskonto5").Value = 0;
                                            }

                                            if (satirVar)
                                            {
                                                int retval = oDocuments.Update();

                                                if (retval == 0)
                                                {
                                                    if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_OK_MODE)
                                                    {
                                                        Handler.SAPApplication.ActivateMenuItem("1304");
                                                    }
                                                    else if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                                                    {
                                                        try
                                                        {
                                                            frmSatinalmaSiparisi.Close();
                                                        }
                                                        catch (Exception)
                                                        {
                                                        }
                                                        var DocEntry = ConstVariables.oCompanyObject.GetNewObjectKey();
                                                        Handler.SAPApplication.OpenForm((BoFormObjectEnum)22, "", DocEntry);
                                                    }
                                                }
                                                else
                                                {
                                                    Handler.SAPApplication.MessageBox("Indirimler uygulanırken hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription());
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Handler.SAPApplication.MessageBox("İndirim güncelleme işleminde hata oluştu." + ex.Message);
                                        }
                                        finally
                                        {
                                            eklemeGuncelleme = false;
                                            eklemeGuncellemeBelgeNo = "";
                                        }
                                    }
                                }
                            }
                        }
                        //if (pVal.ItemUID == "1" && eklemeGuncelleme && !pVal.BeforeAction)
                        //{
                        //    SAPbobsCOM.Documents oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

                        //    bool exist = oDocuments.GetByKey(Convert.ToInt32(SASDocEntry));

                        //    if (exist)
                        //    {
                        //        string updatesql = "";
                        //        bool satirVar = false;
                        //        for (int i = 0; i <= oDocuments.Lines.Count - 1; i++)
                        //        {
                        //            satirVar = true;
                        //            oDocuments.Lines.SetCurrentLine(i);

                        //            if (oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value.ToString() != "")
                        //            {
                        //                updatesql = "UPDATE POR1 SET \"BaseEntry\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value + "' where \"DocEntry\" = '" + SASDocEntry + "' and \"LineNum\" = '" + i + "' ";

                        //                ConstVariables.oRecordset.DoQuery(updatesql);
                        //            }

                        //            if (oDocuments.Lines.UserFields.Fields.Item("U_SSLineNum").Value.ToString() != "")
                        //            {
                        //                updatesql = "UPDATE POR1 SET \"BaseLine\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSLineNum").Value + "' where \"DocEntry\" = '" + SASDocEntry + "' and \"LineNum\" = '" + i + "' ";

                        //                ConstVariables.oRecordset.DoQuery(updatesql);
                        //            }

                        //            if (oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value.ToString() != "")
                        //            {
                        //                updatesql = "UPDATE POR1 SET \"BaseType\" = '" + oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value + "' where \"DocEntry\" = '" + SASDocEntry + "' and \"LineNum\" = '" + i + "' ";

                        //                ConstVariables.oRecordset.DoQuery(updatesql);

                        //            }

                        //            //oDocuments.Lines.BaseEntry = Convert.ToInt32(oDocuments.Lines.UserFields.Fields.Item("U_SSDocEntry").Value);
                        //            //oDocuments.Lines.BaseLine = Convert.ToInt32(oDocuments.Lines.UserFields.Fields.Item("U_SSLineNum").Value);
                        //            //oDocuments.Lines.BaseType = Convert.ToInt32(oDocuments.Lines.UserFields.Fields.Item("U_SSDocType").Value);
                        //        }

                        //        if (satirVar)
                        //        {
                        //            string body = string.Empty;
                        //            using (StreamReader reader = new StreamReader(System.Windows.Forms.Application.StartupPath + "\\PDF\\mail.html"))
                        //            {
                        //                body = reader.ReadToEnd();
                        //            }

                        //            string kimden = "support@aifteam.com";
                        //            string sifre = "Ncckr1989.";
                        //            SmtpClient SMTP = new SmtpClient();
                        //            MailMessage Mail = new MailMessage();

                        //            Mail.From = new MailAddress(kimden);
                        //            Mail.Subject = "Onay isteği";
                        //            Mail.IsBodyHtml = true;

                        //            CreateAttachment(SASDocEntry.ToString());
                        //            Attachment attachment = new Attachment(System.Windows.Forms.Application.StartupPath + "\\PDF\\" + SASDocEntry + ".pdf");
                        //            Mail.Attachments.Add(attachment);

                        //            string anayazi = "1 adet " + oDocuments.DocTotal.ToString("#,#0.00") + " tutarında yeni satınalma siparişi yanıtınızı beklemektedir.";

                        //            string hostname = "http://172.55.10.8:3100" + "//Home//Index?docEntry=" + SASDocEntry.ToString();
                        //            //ConstVariables.oRecordset.Fields.Item("U_HostAdresi").Value.ToString() + "\\Home\\Index?guid=" + guid;
                        //            body = body.Replace("{yazi}", anayazi);
                        //            body = body.Replace("{onaylink}", hostname + "&onayTipi=1");
                        //            body = body.Replace("{redlink}", hostname + "&onayTipi=2");
                        //            body = body.Replace("{Yil}", DateTime.Now.Year.ToString());

                        //            Mail.To.Add("murat.yenisen@aifteam.com");

                        //            //Mail.CC.Add("kadir.degirmenci@aifteam.com");
                        //            //Mail.CC.Add("fatih.gulen@aifteam.com");

                        //            SMTP.Host = "smtp.outlook.com";
                        //            SMTP.DeliveryMethod = SmtpDeliveryMethod.Network;
                        //            SMTP.EnableSsl = true;
                        //            SMTP.UseDefaultCredentials = false;
                        //            SMTP.Credentials = new System.Net.NetworkCredential(kimden, sifre);
                        //            SMTP.Port = 587;

                        //            Mail.Body = body;

                        //            SMTP.Send(Mail);

                        //            //var retval = oDocuments.Update();

                        //            //if (retval == 0)
                        //            //{
                        //            //}
                        //            //else
                        //            //{
                        //            //    var tst = ConstVariables.oCompanyObject.GetLastErrorDescription();
                        //            //    Handler.SAPApplication.MessageBox("Hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription());
                        //            //}

                        //            eklemeGuncelleme = false;
                        //        }
                        //    }
                        //}
                    }
                    break;

                case BoEventTypes.et_KEY_DOWN:
                    break;

                case BoEventTypes.et_GOT_FOCUS:
                    break;

                case BoEventTypes.et_LOST_FOCUS:
                    if (Program.mKod == "010OTATURVT")
                    {
                        if (pVal.ItemUID == "38" && (pVal.ColUID == "U_NavlunBirim" || pVal.ColUID == "11") && !pVal.BeforeAction)
                        {
                            try
                            {
                                // RDR2.Linetotal = RDR1.U_NavlunBirim * RDR1.Quantity
                                // POR1.Linetotal = POR1.U_NavlunBirim * POR1.Quantity
                                double indirim = 0;

                                double navlunbirim = parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_NavlunBirim").Cells.Item(pVal.Row).Specific).Value);
                                double miktar = parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific).Value);

                                if (navlunbirim.ToString() != "" && miktar.ToString() != "")
                                {
                                    if (navlunbirim == 0)
                                    {
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("112").Cells.Item(pVal.Row).Specific).Value = "";
                                    }
                                    else
                                    {
                                        indirim = navlunbirim * miktar;

                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("112").Cells.Item(pVal.Row).Specific).Value = indirim.ToString();
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                            }
                        }
                    }
                    break;

                case BoEventTypes.et_COMBO_SELECT:
                    if (Program.mKod == "010OTATURVT")
                    {
                        if (pVal.ItemUID == "Item_3" && !pVal.BeforeAction)
                        {
                            string comboVal = ((SAPbouiCOM.ComboBox)frmSatinalmaSiparisi.Items.Item("Item_5").Specific).Value.Trim();

                            if (oncekiIndirimComboDegeri == comboVal)
                            {
                                return false;
                            }

                            if (yenidencomboseciliyor)
                            {
                                yenidencomboseciliyor = false;
                                return false;
                            }

                            string xmlSiparis = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                            var rowsSiparis = (from x in XDocument.Parse(xmlSiparis).Descendants("Row")
                                               select new
                                               {
                                                   SAPIndirim = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "15" select new XElement(y.Element("Value"))).First().Value),
                                               }).ToList();

                            var ocomboIndirimTipi = ((SAPbouiCOM.ComboBox)frmSatinalmaSiparisi.Items.Item("Item_5").Specific).Value.Trim();
                            if (ocomboIndirimTipi == "3")
                            {
                                if (rowsSiparis.Sum(x => Convert.ToDouble(x.SAPIndirim)) > 0)
                                {
                                    var retval = Handler.SAPApplication.MessageBox("Iskonto uygulanmasın seçtiğiniz için var olan indirimler sıfırlanacaktır, devam etmek istiyor musunuz?", 1, "Evet", "Hayır");

                                    if (retval == 2)
                                    {
                                        yenidencomboseciliyor = true;
                                        ((SAPbouiCOM.ComboBox)frmSatinalmaSiparisi.Items.Item("Item_5").Specific).Select(oncekiIndirimComboDegeri, BoSearchKey.psk_ByValue);
                                        BubbleEvent = false;
                                        return false;
                                    }
                                }
                            }
                            else if (ocomboIndirimTipi == "2")
                            {
                                var retval = Handler.SAPApplication.MessageBox("Indirimler değişmiş ise ekrandaki veriler ile yeniden hesaplanacaktır, devam etmek istiyor musunuz?", 1, "Evet", "Hayır");

                                if (retval == 2)
                                {
                                    yenidencomboseciliyor = true;
                                    ((SAPbouiCOM.ComboBox)frmSatinalmaSiparisi.Items.Item("Item_5").Specific).Select(oncekiIndirimComboDegeri, BoSearchKey.psk_ByValue);
                                    BubbleEvent = false;
                                    return false;
                                }
                            }
                            else if (ocomboIndirimTipi == "1")
                            {
                                var retval = Handler.SAPApplication.MessageBox("Indirimler değişmiş ise tablodaki veriler ile yeniden hesaplanacaktır, devam etmek istiyor musunuz?", 1, "Evet", "Hayır");

                                if (retval == 2)
                                {
                                    yenidencomboseciliyor = true;
                                    ((SAPbouiCOM.ComboBox)frmSatinalmaSiparisi.Items.Item("Item_5").Specific).Select(oncekiIndirimComboDegeri, BoSearchKey.psk_ByValue);
                                    BubbleEvent = false;
                                    return false;
                                }
                            }
                        }
                        else if (pVal.BeforeAction && pVal.ItemUID == "Item_5")
                        {
                            oncekiIndirimComboDegeri = ((SAPbouiCOM.ComboBox)frmSatinalmaSiparisi.Items.Item("Item_5").Specific).Value.Trim();
                        }
                    }
                    break;

                case BoEventTypes.et_CLICK:
                    if (Program.mKod == "010OTATURVT")
                    {
                        if (pVal.ItemUID == "Item_0" && !pVal.BeforeAction)
                        {
                            var durum = ((SAPbouiCOM.ComboBox)frmSatinalmaSiparisi.Items.Item("81").Specific).Value.Trim();

                            if (durum != "6")
                            {
                                if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    Handler.SAPApplication.MessageBox("Belge ekleme modunda olduğundan indirimler getirilemez.");
                                    return false;
                                }
                            }
                            var cml = frmSatinalmaSiparisi.GetAsXML();
                            if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_UPDATE_MODE)
                            {
                                frmSatinalmaSiparisi.Items.Item("1").Click();
                            }

                            bool basarili = false;
                            var coluid = oMatrix.GetCellFocus();

                            frmSatinalmaSiparisi.Freeze(true);

                            string sql = "Select T2.\"U_UrunKodu\",T2.\"U_BirinciIsk\",T2.\"U_IkinciIsk\",T2.\"U_UcuncuIsk\",T2.\"U_DorduncuIsk\",T2.\"U_BesinciIsk\",T2.\"U_ToplamIsk\" from \"@AIF_SATINALMAISK\" as T0 INNER JOIN \"@AIF_SATINALMAISK1\" as T2 ON T0.\"DocEntry\" = T2.\"DocEntry\" where '" + oEdtDocDate.Value + "' >= T0.\"U_BaslangicTarihi\" and '" + oEdtDocDate.Value + "' <= T0.\"U_BitisTarihi\" and T0.\"U_Durum\" = '1' and T0.\"U_SaticiKodu\" = '" + oEdtCardCode.Value + "' ";

                            SAPbouiCOM.ProgressBar oProgressBar = null;

                            try
                            {
                                ConstVariables.oRecordset.DoQuery(sql);

                                string xml = ConstVariables.oRecordset.GetAsXML();

                                XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                                XDocument xDoc = XDocument.Parse(ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData));

                                var rows = (from t in xDoc.Descendants(ns + "Row")
                                            select new
                                            {
                                                UrunKodu = (from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_UrunKodu" select new XElement(k.Element(ns + "Value"))).First().Value,
                                                BirinciIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_BirinciIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                IkinciIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_IkinciIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                UcuncuIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_UcuncuIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                DorduncuIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_DorduncuIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                BesinciIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_BesinciIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                ToplamIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_ToplamIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                            }).ToList();

                                SAPbobsCOM.Documents oDocuments = null;

                                if (durum == "6")
                                {
                                    oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                    oDocuments.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders;
                                }
                                else
                                {
                                    oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
                                }

                                var docentry = frmSatinalmaSiparisi.DataSources.DBDataSources.Item("OPOR").GetValue("DocEntry", 0);

                                if (docentry != "")
                                {
                                    var exist = oDocuments.GetByKey(Convert.ToInt32(docentry));

                                    if (exist)
                                    {
                                        bool satirVar = false;
                                        for (int i = 0; i <= oDocuments.Lines.Count - 1; i++)
                                        {
                                            oDocuments.Lines.SetCurrentLine(i);
                                            var first = rows.Where(x => x.UrunKodu == oDocuments.Lines.ItemCode).FirstOrDefault();

                                            if (first != null)
                                            {
                                                satirVar = true;
                                                oDocuments.Lines.UserFields.Fields.Item("U_Iskonto1").Value = first.BirinciIskonto.ToString();
                                                oDocuments.Lines.UserFields.Fields.Item("U_Iskonto2").Value = first.BirinciIskonto.ToString();
                                                oDocuments.Lines.UserFields.Fields.Item("U_Iskonto3").Value = first.UcuncuIskonto.ToString();
                                                oDocuments.Lines.UserFields.Fields.Item("U_Iskonto4").Value = first.DorduncuIskonto.ToString();
                                                oDocuments.Lines.UserFields.Fields.Item("U_Iskonto5").Value = first.BesinciIskonto.ToString();
                                                oDocuments.Lines.UserFields.Fields.Item("DiscPrcnt").Value = first.ToplamIskonto.ToString();
                                            }
                                        }

                                        if (satirVar)
                                        {
                                            var retval = oDocuments.Update();

                                            if (retval == 0)
                                            {
                                                //Handler.SAPApplication.ActivateMenuItem("1304");
                                                Handler.SAPApplication.StatusBar.SetText("İndirim getirme işlemi tamamlandı.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                                if (durum == "6")
                                                {
                                                    frmSatinalmaSiparisi.Items.Item("2").Click();
                                                    var draftDocEntry = ConstVariables.oCompanyObject.GetNewObjectKey();
                                                    Handler.SAPApplication.OpenForm((BoFormObjectEnum)112, "", draftDocEntry);
                                                }
                                                else
                                                {
                                                    Handler.SAPApplication.ActivateMenuItem("1304");
                                                }
                                            }
                                            else
                                            {
                                                Handler.SAPApplication.StatusBar.SetText("İndirim getirme işleminde hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                            }
                                        }
                                    }
                                    //string xmlSiparis = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                    //var rowsSiparis = (from x in XDocument.Parse(xmlSiparis).Descendants("Row")
                                    //                   select new
                                    //                   {
                                    //                       ItemCode = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "1" select new XElement(y.Element("Value"))).First().Value,
                                    //                       SatirNumarasi = x.ElementsBeforeSelf().Count() + 1
                                    //                   }).ToList();
                                    ////    //Hakan
                                    //oProgressBar = Handler.SAPApplication.StatusBar.CreateProgressBar("İndirimler Getiriliyor...", oMatrix.RowCount, true);
                                    //int Progress = 0;
                                    //oProgressBar.Text = "İndirimler Getiriliyor...";
                                    //foreach (var item in rowsSiparis)
                                    //{
                                    //    var first = rows.Where(x => x.UrunKodu == item.ItemCode).FirstOrDefault();
                                    //    if (first != null)
                                    //    {
                                    //        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Iskonto1").Cells.Item(item.SatirNumarasi).Specific).Value = first.BirinciIskonto.ToString();
                                    //        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Iskonto2").Cells.Item(item.SatirNumarasi).Specific).Value = first.IkinciIskonto.ToString();
                                    //        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Iskonto3").Cells.Item(item.SatirNumarasi).Specific).Value = first.UcuncuIskonto.ToString();
                                    //        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Iskonto4").Cells.Item(item.SatirNumarasi).Specific).Value = first.DorduncuIskonto.ToString();
                                    //        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Iskonto5").Cells.Item(item.SatirNumarasi).Specific).Value = first.BesinciIskonto.ToString();
                                    //        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(item.SatirNumarasi).Specific).Value = first.ToplamIskonto.ToString();
                                    //    }

                                    //    Progress += 1;
                                    //    oProgressBar.Value = Progress;
                                    //}

                                    //if (coluid != null)
                                    //{
                                    //    oMatrix.SetCellFocus(coluid.rowIndex, coluid.ColumnIndex);
                                    //}
                                    //else
                                    //{
                                    //    oMatrix.SetCellFocus(oMatrix.RowCount - 1, 1);
                                    //}

                                    //basarili = true;
                                }
                            }
                            catch (Exception ex)
                            {
                                Handler.SAPApplication.StatusBar.SetText("İndirim getirme işleminde hata oluştu.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            }
                            finally
                            {
                                if (oProgressBar != null)
                                {
                                    oProgressBar.Stop();
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar);
                                    GC.Collect();
                                }
                                frmSatinalmaSiparisi.Freeze(false);
                            }

                            //if (basarili)
                            //{
                            //    Handler.SAPApplication.StatusBar.SetText("İndirim getirme işlemi tamamlandı.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            //}
                        }
                        else if (!pVal.BeforeAction && pVal.ItemUID == "Item_1")
                        {
                            var durum = ((SAPbouiCOM.ComboBox)frmSatinalmaSiparisi.Items.Item("81").Specific).Value.Trim();

                            if (durum != "6")
                            {
                                if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    Handler.SAPApplication.MessageBox("Belge ekleme modunda olduğundan indirimler getirilemez.");
                                    return false;
                                }
                            }
                            var cml = frmSatinalmaSiparisi.GetAsXML();
                            if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_UPDATE_MODE)
                            {
                                frmSatinalmaSiparisi.Items.Item("1").Click();
                            }

                            SAPbobsCOM.Documents oDocuments = null;
                            string docentry = "";
                            bool satirVarmi = false;
                            if (durum == "6")
                            {
                                oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                oDocuments.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders;
                                docentry = frmSatinalmaSiparisi.DataSources.DBDataSources.Item("OPOR").GetValue("DocEntry", 0);
                                Handler.SAPApplication.ActivateMenuItem("5907");
                            }
                            else
                            {
                                oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
                                docentry = frmSatinalmaSiparisi.DataSources.DBDataSources.Item("OPOR").GetValue("DocEntry", 0);
                            }

                            double birinciiskonto = 0;
                            double ikinciiskonto = 0;
                            double ucuncuiskonto = 0;
                            double dorduncuiskonto = 0;
                            double besinciiskonto = 0;
                            double toplamiskonto = 0;
                            var exist = oDocuments.GetByKey(Convert.ToInt32(docentry));
                            if (exist)
                            {
                                for (int i = 0; i <= oDocuments.Lines.Count - 1; i++)
                                {
                                    oDocuments.Lines.SetCurrentLine(i);
                                    birinciiskonto = parseNumber.parservalues<double>(oDocuments.Lines.UserFields.Fields.Item("U_Iskonto1").Value.ToString());
                                    ikinciiskonto = parseNumber.parservalues<double>(oDocuments.Lines.UserFields.Fields.Item("U_Iskonto2").Value.ToString());
                                    ucuncuiskonto = parseNumber.parservalues<double>(oDocuments.Lines.UserFields.Fields.Item("U_Iskonto3").Value.ToString());
                                    dorduncuiskonto = parseNumber.parservalues<double>(oDocuments.Lines.UserFields.Fields.Item("U_Iskonto4").Value.ToString());
                                    besinciiskonto = parseNumber.parservalues<double>(oDocuments.Lines.UserFields.Fields.Item("U_Iskonto5").Value.ToString());

                                    toplamiskonto = 100 - (100 * (1 - (birinciiskonto / 100)) * (1 - (ikinciiskonto / 100)) * (1 - (ucuncuiskonto / 100)) * (1 - (dorduncuiskonto / 100)) * (1 - (besinciiskonto / 100)));

                                    oDocuments.Lines.DiscountPercent = toplamiskonto;

                                    satirVarmi = true;
                                }

                                if (satirVarmi)
                                {
                                    var retval = oDocuments.Update();

                                    if (retval == 0)
                                    {
                                        if (durum == "6")
                                        {
                                            frmSatinalmaSiparisi.Items.Item("2").Click();
                                            var draftDocEntry = ConstVariables.oCompanyObject.GetNewObjectKey();
                                            Handler.SAPApplication.OpenForm((BoFormObjectEnum)112, "", draftDocEntry);
                                        }
                                        else
                                        {
                                            Handler.SAPApplication.ActivateMenuItem("1304");
                                        }

                                        Handler.SAPApplication.StatusBar.SetText("İndirim getirme işlemi tamamlandı.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                    }
                                }
                            }
                            //SAPbouiCOM.ProgressBar oProgressBar = null;
                            //var coluid = oMatrix.GetCellFocus();
                            //bool basarili = false;
                            //frmSatinalmaSiparisi.Freeze(true);
                            //try
                            //{
                            //    string xmlSiparis = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                            //    var rowsSiparis = (from x in XDocument.Parse(xmlSiparis).Descendants("Row")
                            //                       select new
                            //                       {
                            //                           ItemCode = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "1" select new XElement(y.Element("Value"))).First().Value,
                            //                           Iskonto1 = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "U_Iskonto1" select new XElement(y.Element("Value"))).First().Value),
                            //                           Iskonto2 = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "U_Iskonto2" select new XElement(y.Element("Value"))).First().Value),
                            //                           Iskonto3 = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "U_Iskonto3" select new XElement(y.Element("Value"))).First().Value),
                            //                           Iskonto4 = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "U_Iskonto4" select new XElement(y.Element("Value"))).First().Value),
                            //                           Iskonto5 = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "U_Iskonto5" select new XElement(y.Element("Value"))).First().Value),
                            //                           //SatirNumarasi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "110" select new XElement(y.Element("Value"))).First().Value,
                            //                           SatirNumarasi = x.ElementsBeforeSelf().Count() + 1,
                            //                       }).ToList();

                            //    oProgressBar = Handler.SAPApplication.StatusBar.CreateProgressBar("Iskontolar uygulanıyor...", oMatrix.RowCount, true);
                            //    int Progress = 0;
                            //    oProgressBar.Text = "Iskontolar uygulanıyor...";

                            //    //SAPbobsCOM.Documents oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

                            //    //var docentry = frmSatisSiparisi.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);

                            //    //var exist = oDocuments.GetByKey(Convert.ToInt32(docentry));

                            //    //if (exist)
                            //    //{
                            //    bool satirVarmi = false;
                            //    double _toplamIsk = 0;
                            //    foreach (var item in rowsSiparis.Where(x => x.ItemCode != ""))
                            //    {
                            //        _toplamIsk = 100 - (100 * (1 - (item.Iskonto1 / 100)) * (1 - (item.Iskonto2 / 100)) * (1 - (item.Iskonto3 / 100)) * (1 - (item.Iskonto4 / 100)) * (1 - (item.Iskonto5 / 100)));

                            //        //oDocuments.Lines.SetCurrentLine(Convert.ToInt32(item.SatirNumarasi));

                            //        //oDocuments.Lines.DiscountPercent = _toplamIsk;

                            //        //satirVarmi = true;
                            //        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(item.SatirNumarasi).Specific).Value = _toplamIsk.ToString();

                            //        Progress += 1;
                            //        oProgressBar.Value = Progress;
                            //    }

                            //    //if (satirVarmi)
                            //    //{
                            //    //    int retval = oDocuments.Update();

                            //    //    if (retval == 0)
                            //    //    {
                            //    //        Handler.SAPApplication.ActivateMenuItem("1304");
                            //    //        Handler.SAPApplication.StatusBar.SetText("Iskontolar uygulandı.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                            //    //    }
                            //    //    else
                            //    //    {
                            //    //        Handler.SAPApplication.MessageBox("Iskontolar uygulanırken hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription());
                            //    //    }
                            //    //}

                            //    //}
                            //    if (coluid != null)
                            //    {
                            //        oMatrix.SetCellFocus(coluid.rowIndex, coluid.ColumnIndex);
                            //    }
                            //    else
                            //    {
                            //        oMatrix.SetCellFocus(oMatrix.RowCount - 1, 1);
                            //    }

                            //    basarili = true;

                            //}
                            //catch (Exception)
                            //{
                            //}
                            //finally
                            //{
                            //    frmSatinalmaSiparisi.Freeze(false);
                            //    if (oProgressBar != null)
                            //    {
                            //        oProgressBar.Stop();
                            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar);
                            //        GC.Collect();
                            //    }
                            //}

                            //if (basarili)
                            //{
                            //    Handler.SAPApplication.StatusBar.SetText("Iskontolar uygulandı.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            //}
                        }
                        else if (!pVal.BeforeAction && pVal.ItemUID == "Item_4")
                        {
                            try
                            {
                                var docentry = frmSatinalmaSiparisi.DataSources.DBDataSources.Item("OPOR").GetValue("DocEntry", 0);
                                CreateAttachment(docentry.ToString());

                                try
                                {
                                    string yol = System.Windows.Forms.Application.StartupPath + "\\PDF\\" + docentry + ".pdf";
                                    Process.Start(yol);
                                }
                                catch (Exception)
                                {
                                }
                            }
                            catch (Exception)
                            {
                            }
                        }
                        else if (!pVal.BeforeAction && pVal.ItemUID == "Item_5" && (frmSatinalmaSiparisi.Mode != BoFormMode.fm_ADD_MODE && frmSatinalmaSiparisi.Mode != BoFormMode.fm_FIND_MODE))
                        {
                            string DocEntry = ((SAPbouiCOM.EditText)frmSatinalmaSiparisi.Items.Item("8").Specific).Value.ToString();
                            SAPbobsCOM.Documents oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                            oDocuments.GetByKey(Convert.ToInt32(DocEntry));

                            bool mailgonderildi = mailGonder(oDocuments);
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
                        frmSatinalmaSiparisi = Handler.SAPApplication.Forms.Item(pVal.FormUID);
                        //frmInit = false;
                        var xml = frmSatinalmaSiparisi.GetAsXML();
                        formuid = pVal.FormUID;
                        AIFConn.Sys142.LoadForms();
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

        private static CultureInfo _parsCult;

        public static T parservalues<T>(string val) where T : struct
        {
            object returnvals = 0;
            try
            {
                if (string.IsNullOrEmpty(val)) return (T)Convert.ChangeType(returnvals, typeof(T));

                val = Regex.Replace(val, @"[^0-9.,]+", "").Trim();

                CultureInfo parser = parsCult;
                if (new Regex("^(?!0|\\.00)[0-9]+(,\\d{3})*([.][0-9]{0,9})$").IsMatch(val))
                {
                    parser = System.Globalization.CultureInfo.InvariantCulture;
                }

                if (typeof(T) == typeof(decimal))
                {
                    returnvals = decimal.Parse(val, parser);
                }
                else if (typeof(T) == typeof(double))
                {
                    returnvals = double.Parse(val, parser);
                }
                else if (typeof(T) == typeof(float))
                {
                    returnvals = float.Parse(val, parser);
                }
                else if (typeof(T) == typeof(int))
                {
                    returnvals = (int)decimal.Parse(val, parser);
                }
            }
            catch (Exception ex)
            {
            }

            return (T)returnvals;
        }

        private static CultureInfo parsCult
        {
            get
            {
                if (_parsCult == null)
                {
                    CultureInfo ci = CultureInfo.InvariantCulture;
                    _parsCult = (CultureInfo)ci.Clone();

                    _parsCult.NumberFormat.CurrencyDecimalSeparator = ",";
                    _parsCult.NumberFormat.NumberDecimalSeparator = ",";
                    _parsCult.NumberFormat.PercentDecimalSeparator = ",";
                    _parsCult.NumberFormat.CurrencyGroupSeparator = ".";
                    _parsCult.NumberFormat.NumberGroupSeparator = ".";
                    _parsCult.NumberFormat.PercentGroupSeparator = ".";
                }

                return _parsCult;
            }
        }

        public void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.MenuUID == "5907" && !pVal.BeforeAction)
            {
                if (Program.mKod == "010OTATURVT")
                {
                    #region indirim güncelleme
                    var docentry = ConstVariables.oCompanyObject.GetNewObjectKey();

                    SAPbobsCOM.Documents oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                    oDocuments.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders;

                    var exist = oDocuments.GetByKey(Convert.ToInt32(docentry));

                    if (exist)
                    {
                        if (oDocuments.UserFields.Fields.Item("U_IndirimTipi").Value.ToString() == "1")
                        {
                            string sql = "Select T2.\"U_UrunKodu\",T2.\"U_BirinciIsk\",T2.\"U_IkinciIsk\",T2.\"U_UcuncuIsk\",T2.\"U_DorduncuIsk\",T2.\"U_BesinciIsk\",T2.\"U_ToplamIsk\" from \"@AIF_SATINALMAISK\" as T0 INNER JOIN \"@AIF_SATINALMAISK1\" as T2 ON T0.\"DocEntry\" = T2.\"DocEntry\" where '" + oDocuments.DocDate.ToString("yyyyMMdd") + "' >= T0.\"U_BaslangicTarihi\" and '" + oDocuments.DocDate.ToString("yyyyMMdd") + "' <= T0.\"U_BitisTarihi\" and T0.\"U_Durum\" = '1' and T0.\"U_SaticiKodu\" = '" + oDocuments.CardCode + "' ";

                            try
                            {
                                ConstVariables.oRecordset.DoQuery(sql);

                                string xml = ConstVariables.oRecordset.GetAsXML();

                                XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                                XDocument xDoc = XDocument.Parse(ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData));

                                var rows = (from t in xDoc.Descendants(ns + "Row")
                                            select new
                                            {
                                                UrunKodu = (from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_UrunKodu" select new XElement(k.Element(ns + "Value"))).First().Value,
                                                BirinciIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_BirinciIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                IkinciIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_IkinciIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                UcuncuIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_UcuncuIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                DorduncuIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_DorduncuIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                BesinciIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_BesinciIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                ToplamIskonto = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_ToplamIsk" select new XElement(k.Element(ns + "Value"))).First().Value),
                                            }).ToList();

                                if (exist)
                                {
                                    bool satirVar = false;
                                    for (int i = 0; i <= oDocuments.Lines.Count - 1; i++)
                                    {
                                        oDocuments.Lines.SetCurrentLine(i);
                                        var first = rows.Where(x => x.UrunKodu == oDocuments.Lines.ItemCode).FirstOrDefault();

                                        if (first != null)
                                        {
                                            satirVar = true;
                                            oDocuments.Lines.UserFields.Fields.Item("U_Iskonto1").Value = first.BirinciIskonto.ToString();
                                            oDocuments.Lines.UserFields.Fields.Item("U_Iskonto2").Value = first.IkinciIskonto.ToString();
                                            oDocuments.Lines.UserFields.Fields.Item("U_Iskonto3").Value = first.UcuncuIskonto.ToString();
                                            oDocuments.Lines.UserFields.Fields.Item("U_Iskonto4").Value = first.DorduncuIskonto.ToString();
                                            oDocuments.Lines.UserFields.Fields.Item("U_Iskonto5").Value = first.BesinciIskonto.ToString();
                                            oDocuments.Lines.UserFields.Fields.Item("DiscPrcnt").Value = first.ToplamIskonto.ToString();
                                        }
                                    }

                                    if (satirVar)
                                    {
                                        var retval = oDocuments.Update();

                                        if (retval == 0)
                                        {
                                            oDocuments.GetByKey(Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey()));
                                            //bool mailgonderildi = mailGonder(oDocuments);

                                            if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_OK_MODE)
                                            {
                                                Handler.SAPApplication.ActivateMenuItem("1304");
                                            }
                                            else if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                                            {
                                                try
                                                {
                                                    frmSatinalmaSiparisi.Close();
                                                }
                                                catch (Exception)
                                                {
                                                }
                                                var DocEntry = ConstVariables.oCompanyObject.GetNewObjectKey();
                                                Handler.SAPApplication.OpenForm((BoFormObjectEnum)112, "", DocEntry);
                                            }
                                        }
                                        else
                                        {
                                            Handler.SAPApplication.MessageBox("Indirimler uygulanırken hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription());
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Handler.SAPApplication.MessageBox("İndirim güncelleme işleminde hata oluştu." + ex.Message);
                            }
                            finally
                            {
                            }
                        }
                        else if (oDocuments.UserFields.Fields.Item("U_IndirimTipi").Value.ToString() == "2")
                        {
                            try
                            {
                                string sql = "Select T2.\"ItemCode\",T2.\"LineNum\",T2.\"U_Iskonto1\",T2.\"U_Iskonto2\",T2.\"U_Iskonto3\",T2.\"U_Iskonto4\",T2.\"U_Iskonto5\" from DRF1 as T2 where T2.\"DocEntry\" = '" + eklemeGuncellemeBelgeNo + "'";

                                ConstVariables.oRecordset.DoQuery(sql);

                                string xml = ConstVariables.oRecordset.GetAsXML();

                                XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                                XDocument xDoc = XDocument.Parse(ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData));

                                var rows = (from t in xDoc.Descendants(ns + "Row")
                                            select new
                                            {
                                                UrunKodu = (from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "ItemCode" select new XElement(k.Element(ns + "Value"))).First().Value,
                                                Iskonto1 = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_Iskonto1" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                Iskonto2 = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_Iskonto2" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                Iskonto3 = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_Iskonto3" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                Iskonto4 = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_Iskonto4" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                Iskonto5 = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_Iskonto5" select new XElement(k.Element(ns + "Value"))).First().Value),
                                                SatirNumarasi = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "LineNum" select new XElement(k.Element(ns + "Value"))).First().Value),
                                            }).ToList();

                                bool satirVarmi = false;
                                double _toplamIsk = 0;
                                foreach (var item in rows.Where(x => x.UrunKodu != ""))
                                {
                                    _toplamIsk = 100 - (100 * (1 - (item.Iskonto1 / 100)) * (1 - (item.Iskonto2 / 100)) * (1 - (item.Iskonto3 / 100)) * (1 - (item.Iskonto4 / 100)) * (1 - (item.Iskonto5 / 100)));

                                    oDocuments.Lines.SetCurrentLine(Convert.ToInt32(item.SatirNumarasi));

                                    oDocuments.Lines.DiscountPercent = _toplamIsk;

                                    satirVarmi = true;
                                    //((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(item.SatirNumarasi).Specific).Value = _toplamIsk.ToString();

                                    //Progress += 1;
                                    //oProgressBar.Value = Progress;
                                }

                                if (satirVarmi)
                                {
                                    int retval = oDocuments.Update();

                                    if (retval == 0)
                                    {
                                        //oDocuments.GetByKey(Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey()));
                                        //bool mailgonderildi = mailGonder(oDocuments);
                                        if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_OK_MODE)
                                        {
                                            Handler.SAPApplication.ActivateMenuItem("1304");
                                        }
                                        else if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                                        {
                                            try
                                            {
                                                frmSatinalmaSiparisi.Close();
                                            }
                                            catch (Exception)
                                            {
                                            }
                                            var DocEntry = ConstVariables.oCompanyObject.GetNewObjectKey();
                                            Handler.SAPApplication.OpenForm((BoFormObjectEnum)112, "", DocEntry);
                                        }
                                    }
                                    else
                                    {
                                        Handler.SAPApplication.MessageBox("Indirimler uygulanırken hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription());
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Handler.SAPApplication.MessageBox("İndirim güncelleme işleminde hata oluştu." + ex.Message);
                            }
                            finally
                            {
                                eklemeGuncelleme = false;
                                eklemeGuncellemeBelgeNo = "";
                            }
                        }
                        else if (oDocuments.UserFields.Fields.Item("U_IndirimTipi").Value.ToString() == "3")
                        {
                            try
                            {
                                bool satirVar = false;
                                for (int i = 0; i <= oDocuments.Lines.Count - 1; i++)
                                {
                                    satirVar = true;
                                    oDocuments.Lines.SetCurrentLine(i);

                                    oDocuments.Lines.DiscountPercent = 0;
                                    oDocuments.Lines.UserFields.Fields.Item("U_Iskonto1").Value = 0;
                                    oDocuments.Lines.UserFields.Fields.Item("U_Iskonto2").Value = 0;
                                    oDocuments.Lines.UserFields.Fields.Item("U_Iskonto3").Value = 0;
                                    oDocuments.Lines.UserFields.Fields.Item("U_Iskonto4").Value = 0;
                                    oDocuments.Lines.UserFields.Fields.Item("U_Iskonto5").Value = 0;
                                }

                                if (satirVar)
                                {
                                    int retval = oDocuments.Update();

                                    if (retval == 0)
                                    {
                                        if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_OK_MODE)
                                        {
                                            Handler.SAPApplication.ActivateMenuItem("1304");
                                        }
                                        else if (frmSatinalmaSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                                        {
                                            try
                                            {
                                                frmSatinalmaSiparisi.Close();
                                            }
                                            catch (Exception)
                                            {
                                            }
                                            var DocEntry = ConstVariables.oCompanyObject.GetNewObjectKey();
                                            Handler.SAPApplication.OpenForm((BoFormObjectEnum)112, "", DocEntry);
                                        }
                                    }
                                    else
                                    {
                                        Handler.SAPApplication.MessageBox("Indirimler uygulanırken hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription());
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Handler.SAPApplication.MessageBox("İndirim güncelleme işleminde hata oluştu." + ex.Message);
                            }
                            finally
                            {
                            }
                        }
                    }
                    #endregion
                }
            }
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }
    }
}