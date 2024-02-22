using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using AIF.UVT.SAPB1.HelperClass;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class SatisSiparisi
    {
        [ItemAtt(AIFConn.SatisSiparisi_FormUID)]
        public SAPbouiCOM.Form frmSatisSiparisi;

        static string formuid = "";
        SAPbouiCOM.EditText oEdtCardCode = null;
        SAPbouiCOM.EditText oEdtDocDate = null;
        SAPbouiCOM.Matrix oMatrix = null;
        public void LoadForms()
        {
            if (Program.mKod == "010OTATURVT")
            {
                Functions.CreateUserOrSystemFormComponent<SatisSiparisi>(AIFConn.Sys139, true, formuid);

                System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
                System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("AIF.UVT.SAPB1.FormsView.SatisSiparisi.xml");

                System.IO.StreamReader streamreader = new System.IO.StreamReader(stream, true);
                xmldoc.LoadXml(string.Format(streamreader.ReadToEnd(), formuid));
                Handler.SAPApplication.LoadBatchActions(xmldoc.InnerXml);

                streamreader.Close();

                var cml = frmSatisSiparisi.GetAsXML();
                InitForms();
            }
        }
        public void InitForms()
        {
            try
            {
                frmSatisSiparisi.Freeze(true);
                //frmSatisSiparisi.Items.Item("Item_0").Top = frmSatisSiparisi.Items.Item("2").Top;
                //frmSatisSiparisi.Items.Item("Item_0").Left = frmSatisSiparisi.Items.Item("2").Left + 90;
                //frmSatisSiparisi.Items.Item("Item_0").Height = frmSatisSiparisi.Items.Item("2").Height;

                //frmSatisSiparisi.Items.Item("Item_1").Top = frmSatisSiparisi.Items.Item("Item_0").Top;
                //frmSatisSiparisi.Items.Item("Item_1").Left = frmSatisSiparisi.Items.Item("Item_0").Left + 110;
                //frmSatisSiparisi.Items.Item("Item_1").Height = frmSatisSiparisi.Items.Item("Item_0").Height;

                oEdtCardCode = (SAPbouiCOM.EditText)frmSatisSiparisi.Items.Item("4").Specific;
                oEdtDocDate = (SAPbouiCOM.EditText)frmSatisSiparisi.Items.Item("10").Specific;
                oMatrix = (SAPbouiCOM.Matrix)frmSatisSiparisi.Items.Item("38").Specific;

                //btnSatinalmaSipOlustur.Item.Top = btnIptal.Item.Top;
                //btnSatinalmaSipOlustur.Item.LinkTo = "2"; 


                frmSatisSiparisi.Items.Item("Item_2").Top = frmSatisSiparisi.Items.Item("2").Top;
                frmSatisSiparisi.Items.Item("Item_2").Left = frmSatisSiparisi.Items.Item("2").Left + 85;
                frmSatisSiparisi.Items.Item("Item_2").Width = 150;
                frmSatisSiparisi.Items.Item("Item_2").LinkTo = "2";


                frmSatisSiparisi.Items.Item("Item_3").Top = frmSatisSiparisi.Items.Item("2").Top - 20;
                frmSatisSiparisi.Items.Item("Item_3").Left = frmSatisSiparisi.Items.Item("16").Left;
                frmSatisSiparisi.Items.Item("Item_4").Top = frmSatisSiparisi.Items.Item("Item_3").Top;
            }
            catch (Exception ex)
            {
            }
            finally
            {
                frmSatisSiparisi.Freeze(false);
            }
        }
        bool yenidencomboseciliyor = false;
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
                            eklemeGuncelleme = true;
                            eklemeGuncellemeBelgeNo = frmSatisSiparisi.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).ToString();
                        }
                    }
                    break;
                case BoEventTypes.et_FORM_DATA_UPDATE:
                    if (Program.mKod == "010OTATURVT")
                    {
                        if (!BusinessObjectInfo.BeforeAction)
                        {
                            eklemeGuncelleme = true;
                            eklemeGuncellemeBelgeNo = frmSatisSiparisi.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0).ToString();
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

        public static string docEntrySatisSiparisi = "";
        public static string cardCodeSatisSiparisi = "";
        string val = "";
        bool eklemeGuncelleme = false;
        string eklemeGuncellemeBelgeNo = "";
        string oncekiIndirimComboDegeri = "";
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
                                SAPbobsCOM.Documents oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

                                var exist = oDocuments.GetByKey(Convert.ToInt32(eklemeGuncellemeBelgeNo));

                                if (exist)
                                {
                                    if (oDocuments.UserFields.Fields.Item("U_IndirimTipi").Value.ToString() == "1")
                                    {
                                        string sql = "Select T2.\"U_UrunKodu\",T2.\"U_BirinciIsk\",T2.\"U_IkinciIsk\",T2.\"U_UcuncuIsk\",T2.\"U_DorduncuIsk\",T2.\"U_BesinciIsk\",T2.\"U_ToplamIsk\" from \"@AIF_INDIRIMLER\" as T0 INNER JOIN \"@AIF_INDIRIMLER2\" as T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" INNER JOIN \"@AIF_INDIRIMLER1\" as T2 ON T0.\"DocEntry\" = T2.\"DocEntry\" where '" + oDocuments.DocDate.ToString("yyyyMMdd") + "' >= T0.\"U_Baslangic\" and '" + oDocuments.DocDate.ToString("yyyyMMdd") + "' <= T0.\"U_Bitis\" and T0.\"U_Durum\" = '1' and T1.\"U_MusteriKodu\" = '" + oDocuments.CardCode + "' ";


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
                                                double iskontoToplami = 0;
                                                bool satirVar = false;
                                                for (int i = 0; i <= oDocuments.Lines.Count - 1; i++)
                                                {
                                                    oDocuments.Lines.SetCurrentLine(i);
                                                    var first = rows.Where(x => x.UrunKodu == oDocuments.Lines.ItemCode).FirstOrDefault();

                                                    if (first != null)
                                                    {
                                                        satirVar = true;
                                                        oDocuments.Lines.UserFields.Fields.Item("U_Iskonto1").Value = first.BirinciIskonto.ToString().Replace(",", ".");
                                                        oDocuments.Lines.UserFields.Fields.Item("U_Iskonto2").Value = first.IkinciIskonto.ToString().Replace(",", ".");
                                                        oDocuments.Lines.UserFields.Fields.Item("U_Iskonto3").Value = first.UcuncuIskonto.ToString().Replace(",", ".");
                                                        oDocuments.Lines.UserFields.Fields.Item("U_Iskonto4").Value = first.DorduncuIskonto.ToString().Replace(",", ".");
                                                        oDocuments.Lines.UserFields.Fields.Item("U_Iskonto5").Value = first.BesinciIskonto.ToString(System.Globalization.CultureInfo.InvariantCulture);

                                                        iskontoToplami = parseNumber.parservalues<double>((100 - (100 * (1 - (parseNumber.parservalues<double>(first.BirinciIskonto.ToString()) / 100)) * (1 - (parseNumber.parservalues<double>(first.IkinciIskonto.ToString()) / 100)) * (1 - (parseNumber.parservalues<double>(first.UcuncuIskonto.ToString()) / 100)) * (1 - (parseNumber.parservalues<double>(first.DorduncuIskonto.ToString()) / 100)) * (1 - (parseNumber.parservalues<double>(first.BesinciIskonto.ToString()) / 100)))).ToString());


                                                        oDocuments.Lines.UserFields.Fields.Item("DiscPrcnt").Value = iskontoToplami.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                                    }
                                                }

                                                if (satirVar)
                                                {
                                                    var retval = oDocuments.Update();


                                                    if (retval == 0)
                                                    {
                                                        if (frmSatisSiparisi.Mode == BoFormMode.fm_OK_MODE)
                                                        {
                                                            Handler.SAPApplication.ActivateMenuItem("1304");
                                                        }
                                                        else if (frmSatisSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                                                        {
                                                            try
                                                            {
                                                                frmSatisSiparisi.Close();
                                                            }
                                                            catch (Exception)
                                                            {
                                                            }
                                                            var DocEntry = ConstVariables.oCompanyObject.GetNewObjectKey();
                                                            Handler.SAPApplication.OpenForm((BoFormObjectEnum)17, "", DocEntry);
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
                                            string sql = "Select T2.\"ItemCode\",T2.\"LineNum\",T2.\"U_Iskonto1\",T2.\"U_Iskonto2\",T2.\"U_Iskonto3\",T2.\"U_Iskonto4\",T2.\"U_Iskonto5\" from RDR1 as T2 where T2.\"DocEntry\" = '" + eklemeGuncellemeBelgeNo + "'";


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
                                                            SatirNumarasi = t.ElementsBeforeSelf().Count()//parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "LineNum" select new XElement(k.Element(ns + "Value"))).First().Value),
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
                                                    if (frmSatisSiparisi.Mode == BoFormMode.fm_OK_MODE)
                                                    {
                                                        Handler.SAPApplication.ActivateMenuItem("1304");
                                                    }
                                                    else if (frmSatisSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                                                    {
                                                        try
                                                        {
                                                            frmSatisSiparisi.Close();
                                                        }
                                                        catch (Exception)
                                                        {
                                                        }
                                                        var DocEntry = ConstVariables.oCompanyObject.GetNewObjectKey();
                                                        Handler.SAPApplication.OpenForm((BoFormObjectEnum)17, "", DocEntry);
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
                                                    if (frmSatisSiparisi.Mode == BoFormMode.fm_OK_MODE)
                                                    {
                                                        Handler.SAPApplication.ActivateMenuItem("1304");
                                                    }
                                                    else if (frmSatisSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                                                    {
                                                        try
                                                        {
                                                            frmSatisSiparisi.Close();
                                                        }
                                                        catch (Exception)
                                                        {
                                                        }
                                                        var DocEntry = ConstVariables.oCompanyObject.GetNewObjectKey();
                                                        Handler.SAPApplication.OpenForm((BoFormObjectEnum)17, "", DocEntry);
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

                            //try
                            //{


                            //    string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                            //    var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                            //                select new
                            //                {
                            //                    UrunKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "1" select new XElement(y.Element("Value"))).First().Value,
                            //                    NavlunBirim = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "U_NavlunBirim" select new XElement(y.Element("Value"))).First().Value,
                            //                    Miktar = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "11" select new XElement(y.Element("Value"))).First().Value,
                            //                }).ToList();




                            //    double indirim = 0;
                            //    foreach (var item in rows.Where(x => x.UrunKodu != ""))
                            //    {
                            //        indirim = Convert.ToDouble(item.NavlunBirim) * Convert.ToDouble(item.Miktar);

                            //        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("112").Cells.Item(pVal.Row).Specific).Value = indirim.ToString();

                            //        //oDocuments.Lines.LineTotal = indirim;
                            //    }
                            //}
                            //catch (Exception ex)
                            //{
                            //} 

                        }
                    }
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    if (Program.mKod == "010OTATURVT")
                    {
                        if (pVal.ItemUID == "Item_3" && !pVal.BeforeAction)
                        {
                            string comboVal = ((SAPbouiCOM.ComboBox)frmSatisSiparisi.Items.Item("Item_3").Specific).Value.Trim();

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

                            var ocomboIndirimTipi = ((SAPbouiCOM.ComboBox)frmSatisSiparisi.Items.Item("Item_3").Specific).Value.Trim();
                            if (ocomboIndirimTipi == "3")
                            {
                                if (rowsSiparis.Sum(x => Convert.ToDouble(x.SAPIndirim)) > 0)
                                {
                                    var retval = Handler.SAPApplication.MessageBox("Iskonto uygulanmasın seçtiğiniz için var olan indirimler sıfırlanacaktır, devam etmek istiyor musunuz?", 1, "Evet", "Hayır");

                                    if (retval == 2)
                                    {
                                        yenidencomboseciliyor = true;
                                        ((SAPbouiCOM.ComboBox)frmSatisSiparisi.Items.Item("Item_3").Specific).Select(oncekiIndirimComboDegeri, BoSearchKey.psk_ByValue);
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
                                    ((SAPbouiCOM.ComboBox)frmSatisSiparisi.Items.Item("Item_3").Specific).Select(oncekiIndirimComboDegeri, BoSearchKey.psk_ByValue);
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
                                    ((SAPbouiCOM.ComboBox)frmSatisSiparisi.Items.Item("Item_3").Specific).Select(oncekiIndirimComboDegeri, BoSearchKey.psk_ByValue);
                                    BubbleEvent = false;
                                    return false;
                                }
                            }
                        }
                        else if (pVal.BeforeAction && pVal.ItemUID == "Item_3")
                        {
                            oncekiIndirimComboDegeri = ((SAPbouiCOM.ComboBox)frmSatisSiparisi.Items.Item("Item_3").Specific).Value.Trim();
                        }
                    }
                    break;
                case BoEventTypes.et_CLICK:
                    if (Program.mKod == "010OTATURVT")
                    {
                        if (pVal.ItemUID == "Item_0" && !pVal.BeforeAction)
                        {
                            //var cml = frmSatisSiparisi.GetAsXML();
                            if (frmSatisSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                Handler.SAPApplication.MessageBox("Belge ekleme modunda olduğundan indirimler getirilemez.");
                                return false;
                            }
                            if (frmSatisSiparisi.Mode == BoFormMode.fm_UPDATE_MODE)
                            {
                                frmSatisSiparisi.Items.Item("1").Click();
                            }

                            //bool basarili = false;
                            //var coluid = oMatrix.GetCellFocus();

                            //frmSatisSiparisi.Freeze(true);

                            string sql = "Select T2.\"U_UrunKodu\",T2.\"U_BirinciIsk\",T2.\"U_IkinciIsk\",T2.\"U_UcuncuIsk\",T2.\"U_DorduncuIsk\",T2.\"U_BesinciIsk\",T2.\"U_ToplamIsk\" from \"@AIF_INDIRIMLER\" as T0 INNER JOIN \"@AIF_INDIRIMLER2\" as T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" INNER JOIN \"@AIF_INDIRIMLER1\" as T2 ON T0.\"DocEntry\" = T2.\"DocEntry\" where '" + oEdtDocDate.Value + "' >= T0.\"U_Baslangic\" and '" + oEdtDocDate.Value + "' <= T0.\"U_Bitis\" and T0.\"U_Durum\" = '1' and T1.\"U_MusteriKodu\" = '" + oEdtCardCode.Value + "' ";

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

                                SAPbobsCOM.Documents oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

                                var docentry = frmSatisSiparisi.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);

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
                                            Handler.SAPApplication.ActivateMenuItem("1304");
                                            Handler.SAPApplication.StatusBar.SetText("İndirim getirme işlemi tamamlandı.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                                        }
                                        else
                                        {
                                            Handler.SAPApplication.StatusBar.SetText("İndirim getirme işleminde hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        }
                                    }
                                }

                            }
                            catch (Exception ex)
                            {
                                Handler.SAPApplication.StatusBar.SetText("İndirim getirme işleminde hata oluştu.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            }
                            finally
                            {
                                //oProgressBar.Stop();
                                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar);
                                //GC.Collect();
                                //frmSatisSiparisi.Freeze(false);
                            }

                            #region Gerekli mi?
                            //if (basarili)
                            //{
                            //    Handler.SAPApplication.StatusBar.SetText("İndirim getirme işlemi tamamlandı.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            //}
                            //try
                            //{

                            //    try
                            //    {


                            //        frmSatisSiparisi.Freeze(true);
                            //        //string templateCode = edtTemplateCode.Value;

                            //        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            //        oRS.DoQuery("Select * from OSPP where \"CardCode\" = '" + oEdtCardCode.Value + "'");

                            //        string xml = oRS.GetAsXML();

                            //        XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                            //        XDocument xDoc = XDocument.Parse(oRS.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData));

                            //        var rows = (from t in xDoc.Descendants(ns + "Row")
                            //                    select new
                            //                    {
                            //                        ItemCode = (from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "ItemCode" select new XElement(k.Element(ns + "Value"))).First().Value,
                            //                        Discount = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "Discount" select new XElement(k.Element(ns + "Value"))).First().Value),
                            //                        TemplateCode = (from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_TemplateCode" select new XElement(k.Element(ns + "Value"))).First().Value,

                            //                    }).ToList();

                            //        string itemCode = "";
                            //        double price = 0;
                            //        AddDiscountDetails s = new AddDiscountDetails();
                            //        for (int i = 1; i <= oMatrix.RowCount - 1; i++)
                            //        {
                            //            itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value;
                            //            if (((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_AIF_DiscountNo").Cells.Item(i).Specific).Value != "")
                            //            {
                            //                continue;
                            //            }

                            //            if (rows.Where(x => x.ItemCode == itemCode).Count() == 0)
                            //            {
                            //                continue;
                            //            }

                            //            price = parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("14").Cells.Item(i).Specific).Value);
                            //            var val = rows.Where(x => x.ItemCode == itemCode).Select(y => y.Discount).FirstOrDefault();
                            //            var templatecode = rows.Where(x => x.ItemCode == itemCode).Select(y => y.TemplateCode).FirstOrDefault();
                            //            var res = s.addDiscountDetails(Convert.ToInt32(templatecode), price);

                            //            if (res == -1)
                            //            {
                            //                Handler.SAPApplication.StatusBar.SetText("İndirimler getirilirken hata oluştu.");
                            //                return false;
                            //            }
                            //            else
                            //            {
                            //                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_AIF_DiscountNo").Cells.Item(i).Specific).Value = res.ToString();
                            //                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(i).Specific).Value = rows.Where(x => x.ItemCode == itemCode).Select(y => y.Discount).FirstOrDefault().ToString();
                            //            }
                            //            Progress += 1;
                            //            oProgressBar.Value = Progress;
                            //        }
                            //        oMatrix.Columns.Item("1").Cells.Item(1).Click();
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //    }
                            //    finally
                            //    {
                            //        oProgressBar.Stop();
                            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar);
                            //        GC.Collect();
                            //        frmSatisSiparisi.Freeze(false);
                            //    }
                            //}
                            //catch (Exception)
                            //{
                            //} 
                            #endregion
                        }
                        else if (!pVal.BeforeAction && pVal.ItemUID == "Item_1")
                        {
                            if (frmSatisSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                Handler.SAPApplication.MessageBox("Belge ekleme modunda olduğundan iskonto uygulanamaz.");
                                return false;
                            }

                            if (frmSatisSiparisi.Mode == BoFormMode.fm_UPDATE_MODE)
                            {
                                frmSatisSiparisi.Items.Item("1").Click();
                            }
                            SAPbouiCOM.ProgressBar oProgressBar = null;
                            var coluid = oMatrix.GetCellFocus();
                            bool basarili = false;
                            frmSatisSiparisi.Freeze(true);
                            try
                            {
                                string xmlSiparis = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                var rowsSiparis = (from x in XDocument.Parse(xmlSiparis).Descendants("Row")
                                                   select new
                                                   {
                                                       ItemCode = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "1" select new XElement(y.Element("Value"))).First().Value,
                                                       Iskonto1 = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "U_Iskonto1" select new XElement(y.Element("Value"))).First().Value),
                                                       Iskonto2 = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "U_Iskonto2" select new XElement(y.Element("Value"))).First().Value),
                                                       Iskonto3 = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "U_Iskonto3" select new XElement(y.Element("Value"))).First().Value),
                                                       Iskonto4 = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "U_Iskonto4" select new XElement(y.Element("Value"))).First().Value),
                                                       Iskonto5 = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "U_Iskonto5" select new XElement(y.Element("Value"))).First().Value),
                                                       SatirNumarasi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "110" select new XElement(y.Element("Value"))).First().Value,
                                                   }).ToList();

                                //oProgressBar = Handler.SAPApplication.StatusBar.CreateProgressBar("Iskontolar uygulanıyor...", oMatrix.RowCount, true);
                                //int Progress = 0;
                                //oProgressBar.Text = "Iskontolar uygulanıyor...";

                                SAPbobsCOM.Documents oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

                                var docentry = frmSatisSiparisi.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);

                                var exist = oDocuments.GetByKey(Convert.ToInt32(docentry));


                                if (exist)
                                {
                                    bool satirVarmi = false;
                                    double _toplamIsk = 0;
                                    foreach (var item in rowsSiparis.Where(x => x.ItemCode != ""))
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
                                            Handler.SAPApplication.ActivateMenuItem("1304");
                                            Handler.SAPApplication.StatusBar.SetText("Iskontolar uygulandı.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                                        }
                                        else
                                        {
                                            Handler.SAPApplication.MessageBox("Iskontolar uygulanırken hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription());
                                        }
                                    }

                                }
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
                            catch (Exception)
                            {
                            }
                            finally
                            {
                                frmSatisSiparisi.Freeze(false);
                                //oProgressBar.Stop();
                                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar);
                                //GC.Collect();
                                //frmSatisSiparisi.Freeze(false);
                            }

                            //if (basarili)
                            //{
                            //    Handler.SAPApplication.StatusBar.SetText("Iskontolar uygulandı.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            //}
                        }
                        else if (pVal.ItemUID == "Item_2" && !pVal.BeforeAction)
                        {
                            try
                            {
                                if (frmSatisSiparisi.Mode == BoFormMode.fm_ADD_MODE)
                                {
                                    Handler.SAPApplication.MessageBox("Belge ekleme modunda olduğundan indirimler getirilemez.");
                                    return false;
                                }
                                if (frmSatisSiparisi.Mode == BoFormMode.fm_UPDATE_MODE)
                                {
                                    frmSatisSiparisi.Items.Item("1").Click();
                                }
                                string xmlSiparis = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                var rowsSiparis = (from x in XDocument.Parse(xmlSiparis).Descendants("Row")
                                                   select new
                                                   {
                                                       ItemCode = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "1" select new XElement(y.Element("Value"))).First().Value,
                                                       Miktar = parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "11" select new XElement(y.Element("Value"))).First().Value),
                                                       SatirNumarasi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "110" select new XElement(y.Element("Value"))).First().Value,
                                                   }).ToList();


                                string cardCode = "";
                                foreach (var item in rowsSiparis.Where(x => x.ItemCode != ""))
                                {
                                    ConstVariables.oRecordset.DoQuery("Select \"CardCode\" from \"OITM\" where \"ItemCode\" = '" + item.ItemCode + "'");

                                    if (ConstVariables.oRecordset.Fields.Item(0).Value.ToString() != "")
                                    {
                                        cardCode = ConstVariables.oRecordset.Fields.Item(0).Value.ToString();

                                        if (cardCode != "")
                                        {
                                            cardCodeSatisSiparisi = cardCode;
                                            break;
                                        }
                                    }
                                }

                                if (cardCode != "")
                                {

                                    string sql = "Select T2.\"U_UrunKodu\",T2.\"U_BirinciIsk\",T2.\"U_IkinciIsk\",T2.\"U_UcuncuIsk\",T2.\"U_DorduncuIsk\",T2.\"U_BesinciIsk\",T2.\"U_ToplamIsk\" from \"@AIF_SATINALMAISK\" as T0 INNER JOIN \"@AIF_SATINALMAISK1\" as T2 ON T0.\"DocEntry\" = T2.\"DocEntry\" where '" + DateTime.Now.ToString("yyyy-MM-dd") + "' >= T0.\"U_BaslangicTarihi\" and '" + DateTime.Now.ToString("yyyy-MM-dd") + "' <= T0.\"U_BitisTarihi\" and T0.\"U_Durum\" = '1' and T0.\"U_SaticiKodu\" = '" + cardCode + "' ";

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


                                    SAPbobsCOM.Documents oDraft = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                                    oDraft.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders;

                                    oDraft.CardCode = cardCode;

                                    var docentry = frmSatisSiparisi.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);
                                    oDraft.UserFields.Fields.Item("U_SSDocEntry").Value = docentry;
                                    oDraft.UserFields.Fields.Item("U_IndirimTipi").Value = "1";
                                    foreach (var item in rowsSiparis.Where(x => x.ItemCode != ""))
                                    {

                                        oDraft.Lines.ItemCode = item.ItemCode;
                                        oDraft.Lines.Quantity = item.Miktar;

                                        oDraft.Lines.UserFields.Fields.Item("U_Iskonto1").Value = rows.Where(x => x.UrunKodu == item.ItemCode).Count() > 0 ? parservalues<double>(rows.Where(x => x.UrunKodu == item.ItemCode).Select(y => y.BirinciIskonto).FirstOrDefault().ToString()) : 0;
                                        oDraft.Lines.UserFields.Fields.Item("U_Iskonto2").Value = rows.Where(x => x.UrunKodu == item.ItemCode).Count() > 0 ? parservalues<double>(rows.Where(x => x.UrunKodu == item.ItemCode).Select(y => y.IkinciIskonto).FirstOrDefault().ToString()) : 0;
                                        oDraft.Lines.UserFields.Fields.Item("U_Iskonto3").Value = rows.Where(x => x.UrunKodu == item.ItemCode).Count() > 0 ? parservalues<double>(rows.Where(x => x.UrunKodu == item.ItemCode).Select(y => y.UcuncuIskonto).FirstOrDefault().ToString()) : 0;
                                        oDraft.Lines.UserFields.Fields.Item("U_Iskonto4").Value = rows.Where(x => x.UrunKodu == item.ItemCode).Count() > 0 ? parservalues<double>(rows.Where(x => x.UrunKodu == item.ItemCode).Select(y => y.DorduncuIskonto).FirstOrDefault().ToString()) : 0;
                                        oDraft.Lines.UserFields.Fields.Item("U_Iskonto5").Value = rows.Where(x => x.UrunKodu == item.ItemCode).Count() > 0 ? parservalues<double>(rows.Where(x => x.UrunKodu == item.ItemCode).Select(y => y.BesinciIskonto).FirstOrDefault().ToString()) : 0;
                                        oDraft.Lines.DiscountPercent = rows.Where(x => x.UrunKodu == item.ItemCode).Count() > 0 ? parservalues<double>(rows.Where(x => x.UrunKodu == item.ItemCode).Select(y => y.ToplamIskonto).FirstOrDefault().ToString()) : 0; ;

                                        oDraft.Lines.UserFields.Fields.Item("U_SSDocEntry").Value = docentry;
                                        oDraft.Lines.UserFields.Fields.Item("U_SSDocType").Value = "17";
                                        oDraft.Lines.UserFields.Fields.Item("U_SSLineNum").Value = item.SatirNumarasi;

                                        oDraft.Lines.Add();
                                    }

                                    var ret = oDraft.Add();

                                    if (ret == 0)
                                    {
                                        var draftDocEntry = ConstVariables.oCompanyObject.GetNewObjectKey();
                                        Handler.SAPApplication.OpenForm((BoFormObjectEnum)112, "", draftDocEntry);
                                    }
                                    else
                                    {
                                        var asdas = ConstVariables.oCompanyObject.GetLastErrorDescription();
                                        Handler.SAPApplication.MessageBox("Hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription());
                                    }
                                }
                                else
                                {
                                    Handler.SAPApplication.MessageBox("Eşleşen satıcı kodu bulunamadı.");
                                }
                            }
                            catch (Exception)
                            {
                            }

                            //string xmlSiparis = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                            //var rowsSiparis = (from x in XDocument.Parse(xmlSiparis).Descendants("Row")
                            //                   select new
                            //                   {
                            //                       ItemCode = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "1" select new XElement(y.Element("Value"))).First().Value,
                            //                   }).ToList();

                            //string cardCode = "";
                            //foreach (var item in rowsSiparis.Where(x => x.ItemCode != ""))
                            //{
                            //    ConstVariables.oRecordset.DoQuery("Select \"CardCode\" from \"OITM\" where \"ItemCode\" = '" + item.ItemCode + "'");

                            //    if (ConstVariables.oRecordset.Fields.Item(0).Value.ToString() != "")
                            //    {
                            //        cardCode = ConstVariables.oRecordset.Fields.Item(0).Value.ToString();

                            //        if (cardCode != "")
                            //        {
                            //            cardCodeSatisSiparisi = cardCode;
                            //            break;
                            //        }
                            //    }
                            //}

                            //if (cardCode != "")
                            //{
                            //    //AIFConn.Sys142.LoadForms();
                            //    docEntrySatisSiparisi = frmSatisSiparisi.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0);

                            //    Handler.SAPApplication.ActivateMenuItem("2305");

                            //}
                        }
                    }
                    break;
                case BoEventTypes.et_DOUBLE_CLICK:
                    if (Program.mKod == "010OTATURVT")
                    {
                        if (pVal.ColUID == "15" && !pVal.BeforeAction)
                        {
                            AIFConn.IndirimShr.LoadForms(frmSatisSiparisi, "38", pVal.Row);
                        }
                    }
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
                        frmSatisSiparisi = Handler.SAPApplication.Forms.Item(pVal.FormUID);
                        //frmInit = false; 
                        var xml = frmSatisSiparisi.GetAsXML();
                        formuid = pVal.FormUID;
                        AIFConn.Sys139.LoadForms();
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
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }
    }
}
