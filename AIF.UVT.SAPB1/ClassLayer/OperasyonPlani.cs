using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing;
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
    public class OperasyonPlani
    {
        [ItemAtt(AIFConn.OperasyonPlaniUID)]
        public SAPbouiCOM.Form frmOperasyonPlani;

        [ItemAtt("Item_18")]
        public SAPbouiCOM.Matrix oMatrix;
        [ItemAtt("Item_17")]
        public SAPbouiCOM.EditText edtTarih;

        List<Tuple<string, string>> UrunAgaciIstasyonu = new List<Tuple<string, string>>();
        public void LoadForms()
        {
            try
            {
                ConstVariables.oFnc.LoadSAPXML(AIFConn.OperasyonPlaniXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.OperasyonPlaniXML));
                Functions.CreateUserOrSystemFormComponent<OperasyonPlani>(AIFConn.OpPlani);

                frmOperasyonPlani.Freeze(true);
                InitForms();
                frmOperasyonPlani.Freeze(false);
            }
            catch (Exception)
            {
            }
        }
        public void InitForms()
        {
            oMatrix.AutoResizeColumns();
            UrunAgaciIstasyonu.Add(Tuple.Create("U_IST001", "AYRAN"));
            UrunAgaciIstasyonu.Add(Tuple.Create("U_IST002", "YOĞURT"));
            //UrunAgaciIstasyonu.Add(Tuple.Create("IST003", "TELEME"));
            UrunAgaciIstasyonu.Add(Tuple.Create("U_IST004", "KAŞAR PEYNİRİ"));
            UrunAgaciIstasyonu.Add(Tuple.Create("U_IST005", "TEREYAĞI"));


        }
        string header = @"<?xml version=""1.0"" encoding=""UTF-16"" ?><dbDataSources uid=""@AIF_PROD_PLAN1""><rows>{0}</rows></dbDataSources>";

        List<Datas> OriginalData = new List<Datas>();
        //string row = "<row><cells><cell><uid>U_Tur</uid><value>{0}</value></cell><cell><uid>U_Kod</uid><value>{1}</value></cell><cell><uid>U_Adi</uid><value>{2}</value></cell><cell><uid>U_Adet</uid><value>{3}</value></cell><cell><uid>U_KG</uid><value>{4}</value></cell><cell><uid>U_Grup</uid><value>{5}</value></cell></cells></row>";

        string row = "<row><cells><cell><uid>U_Grup</uid><value>{0}</value></cell><cell><uid>U_Tur</uid><value>{1}</value></cell><cell><uid>U_Kod</uid><value>{2}</value></cell><cell><uid>U_Adi</uid><value>{3}</value></cell><cell><uid>U_Seviye</uid><value>{4}</value></cell><cell><uid>U_UstSeviye</uid><value>{5}</value></cell><cell><uid>U_KokSeviye</uid><value>{6}</value></cell><cell><uid>U_Sira</uid><value>{7}</value></cell><cell><uid>U_Palet</uid><value>{8}</value></cell><cell><uid>U_Adet</uid><value>{9}</value></cell><cell><uid>U_KG</uid><value>{10}</value></cell></cells></row>";
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
                        OriginalData = new List<Datas>();
                        var xml = frmOperasyonPlani.DataSources.DBDataSources.Item(1).GetAsXML();

                        var rows = (from x in XDocument.Parse(xml).Descendants("row")
                                    select new Datas()
                                    {
                                        Grup = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).First().Value : "",
                                        Tur = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).First().Value : "",
                                        Kod = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).First().Value : "",
                                        Ad = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).First().Value : "",
                                        Seviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Seviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Seviye" select new XElement(y.Element("value"))).First().Value : "",
                                        UstSeviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_UstSeviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_UstSeviye" select new XElement(y.Element("value"))).First().Value : "",
                                        KokSeviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KokSeviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KokSeviye" select new XElement(y.Element("value"))).First().Value : "",
                                        Sira = x.ElementsBeforeSelf().Count() + 1
                                    }).ToList();

                        OriginalData.AddRange(rows);


                        var empty = rows.Where(x => x.Grup != "" && x.Tur != "").ToList();

                        empty.ForEach(x => x.Grup = "");


                        empty = rows.Where(x => x.Tur != "" && x.Kod != "").ToList();
                        empty.ForEach(x => x.Tur = "");
                        //foreach (var item in empty)
                        //{
                        //    var value = rows.Where(x => x.row == item.row - 1).Select(y => y.Grup).First();
                        //}

                        string data2 = string.Join("", rows.Select(s => string.Format(row, s.Grup, s.Tur, s.Kod, s.Ad, s.Seviye, s.UstSeviye, s.KokSeviye, s.Sira)));



                        frmOperasyonPlani.DataSources.DBDataSources.Item("@AIF_PROD_PLAN1").LoadFromXML(string.Format(header, data2));
                        oMatrix.AutoResizeColumns();

                        var xxx = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                        var vals = rows.Where(x => x.Grup != "").ToList();

                        int color = ColorTranslator.ToOle(System.Drawing.Color.Gray);
                        int forecolor = ColorTranslator.ToOle(System.Drawing.Color.White);

                        foreach (var item in vals)
                        {
                            oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                            oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), forecolor);
                        }

                        color = ColorTranslator.ToOle(System.Drawing.Color.DarkGray);

                        vals = rows.Where(x => x.Tur != "").ToList();

                        foreach (var item in vals)
                        {
                            oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                            oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), -1);
                        }

                        color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                        vals = rows.Where(x => x.Kod != "").ToList();

                        foreach (var item in vals)
                        {
                            oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                            oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), -1);
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


        private void renklendir()
        {

            var xml = frmOperasyonPlani.DataSources.DBDataSources.Item(1).GetAsXML();

            var rows2 = (from x in XDocument.Parse(xml).Descendants("row")
                         select new Datas()
                         {
                             Grup = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).First().Value : "",
                             Tur = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).First().Value : "",
                             Kod = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).First().Value : "",
                             Ad = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).First().Value : "",
                             Seviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Seviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Seviye" select new XElement(y.Element("value"))).First().Value : "",
                             UstSeviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_UstSeviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_UstSeviye" select new XElement(y.Element("value"))).First().Value : "",
                             KokSeviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KokSeviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KokSeviye" select new XElement(y.Element("value"))).First().Value : "",
                             Palet = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Palet" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Palet" select new XElement(y.Element("value"))).First().Value) : 0,
                             Adet = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adet" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adet" select new XElement(y.Element("value"))).First().Value) : 0,
                             KG = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KG" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KG" select new XElement(y.Element("value"))).First().Value) : 0,
                             Sira = x.ElementsBeforeSelf().Count() + 1
                         }).ToList();

            foreach (var item in rows2.Where(x => x.Adet > 0 || x.KG > 0))
            {
                string grup = rows2.Where(x => x.KokSeviye == item.KokSeviye).Select(y => y.Grup).FirstOrDefault();
                string val = UrunAgaciIstasyonu.Where(x => x.Item2 == grup).Select(y => y.Item1).FirstOrDefault();
                int basarili = ColorTranslator.ToOle(System.Drawing.Color.LimeGreen);
                int basarisiz = ColorTranslator.ToOle(System.Drawing.Color.OrangeRed);

                bool resp = false;

                resp = GunlukPersonelPlaniKontrolEt(val);
                if (item.Adet > 0)
                {

                    if (resp)
                    {
                        oMatrix.CommonSetting.SetCellBackColor(Convert.ToInt32(item.Sira), 6, basarili);
                    }
                    else
                    {
                        oMatrix.CommonSetting.SetCellBackColor(Convert.ToInt32(item.Sira), 6, basarisiz);
                    }
                }
                else if (item.KG > 0)
                {
                    if (resp)
                    {
                        oMatrix.CommonSetting.SetCellBackColor(Convert.ToInt32(item.Sira), 7, basarili);
                    }
                    else
                    {
                        oMatrix.CommonSetting.SetCellBackColor(Convert.ToInt32(item.Sira), 7, basarisiz);
                    }
                }
            }



        }


        private bool GunlukPersonelPlaniKontrolEt(string val)
        {
            ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


            DateTime dtTarih = new DateTime(Convert.ToInt32(edtTarih.Value.Substring(0, 4)), Convert.ToInt32(edtTarih.Value.Substring(4, 2)), Convert.ToInt32(edtTarih.Value.Substring(6, 2)));
            string gunfield = "U_Gun" + dtTarih.Day;

            string sql = "Select \"U_PersonelNo\" as \"Personel No\", \"U_PersonelAdi\" as \"Personel Adı\" from \"@AIF_GUNLUKPERSPLAN\" as T0 INNER JOIN \"@AIF_GUNLUKPERSPLAN1\" as T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where T0.\"U_Aylar\" = '" + dtTarih.Month.ToString().PadLeft(2, '0') + "' and T0.\"U_Yil\" = '" + dtTarih.Year.ToString() + "' and (T1.\"U_Bolum1\" = '" + val + "' or T1.\"U_Bolum2\" = '" + val + "' or T1.\"U_Bolum3\" = '" + val + "') and " + gunfield + " = 'X' ";

            //string sql = "Select T0.\"DocEntry\" from \"@AIF_GUNPERSPLAN\" as T0 INNER JOIN \"@AIF_GUNPERSPLAN1\" as T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where \"U_Tarih\" = '" + edtTarih.Value + "' and \"" + val + "\" = 'Y'";

            //string sql = "Select \"DocEntry\" from OWOR as T0 where \"DueDate\" = '" + edtTarih.Value + "' and \"U_Istasyon\" = '" + val + "'";
            ConstVariables.oRecordset.DoQuery(sql);

            if (ConstVariables.oRecordset.RecordCount > 0)
            {
                return true;
            }
            return false;
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
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    break;
                case BoEventTypes.et_CLICK:
                    if (pVal.ItemUID == "Item_23" && !pVal.BeforeAction)
                    {
                        Listele();
                    }
                    else if (pVal.ItemUID == "Item_24" && !pVal.BeforeAction)
                    {
                        /*
                        var xml = frmOperasyonPlani.DataSources.DBDataSources.Item(1).GetAsXML();

                        //var rows = (from x in XDocument.Parse(xml).Descendants("row")
                                    select new Datas2()
                                    {
                                        Grup = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).First().Value : "",
                                        Tur = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).First().Value : "",
                                        Kod = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).First().Value : "",
                                        Ad = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).First().Value : "",
                                        row = x.ElementsBeforeSelf().Count()
                                    }).ToList();

                        var empty = rows.Where(x => x.Tur == "" && x.Kod != "").ToList();

                        empty.ForEach(x => x.Tur = rows.Where(y => y.row == x.row - 1).Select(u => u.Tur).First());


                        empty = rows.Where(x => x.Grup == "").ToList();
                        empty.ForEach(x => x.Grup = rows.Where(y => y.row == x.row - 1).Select(u => u.Grup).First());
                        //foreach (var item in empty)
                        //{
                        //    var value = rows.Where(x => x.row == item.row - 1).Select(y => y.Grup).First();
                        //}

                        string data2 = string.Join("", rows.Select(s => string.Format(row, s.Grup, s.Tur, s.Kod, s.Ad)));

                        frmOperasyonPlani.DataSources.DBDataSources.Item("@AIF_PROD_PLAN1").LoadFromXML(string.Format(header, data2));
                        */


                    }
                    else if (pVal.ColUID == "#" && !pVal.BeforeAction)
                    {
                        try
                        {
                            frmOperasyonPlani.Freeze(true);
                            string kokSeviye = "";
                            string ustSeviye = "";
                            string seviye = "";

                            kokSeviye = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific).Value;
                            ustSeviye = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_8").Cells.Item(pVal.Row).Specific).Value;
                            seviye = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_7").Cells.Item(pVal.Row).Specific).Value;


                            var xml = frmOperasyonPlani.DataSources.DBDataSources.Item(1).GetAsXML();

                            var rows2 = (from x in XDocument.Parse(xml).Descendants("row")
                                         select new Datas()
                                         {
                                             Grup = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).First().Value : "",
                                             Tur = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).First().Value : "",
                                             Kod = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).First().Value : "",
                                             Ad = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).First().Value : "",
                                             Seviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Seviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Seviye" select new XElement(y.Element("value"))).First().Value : "",
                                             UstSeviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_UstSeviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_UstSeviye" select new XElement(y.Element("value"))).First().Value : "",
                                             KokSeviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KokSeviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KokSeviye" select new XElement(y.Element("value"))).First().Value : "",
                                             Palet = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Palet" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Palet" select new XElement(y.Element("value"))).First().Value) : 0,
                                             Adet = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adet" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adet" select new XElement(y.Element("value"))).First().Value) : 0,
                                             KG = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KG" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KG" select new XElement(y.Element("value"))).First().Value) : 0,
                                             Sira = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Sira" select new XElement(y.Element("value"))).Any() == true ? Convert.ToInt32((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Sira" select new XElement(y.Element("value"))).First().Value) : 0,
                                         }).ToList();

                            string data2 = "";
                            data2 = string.Join("", rows2.Select(s => string.Format(row, s.Grup, s.Tur, s.Kod, s.Ad, s.Seviye, s.UstSeviye, s.KokSeviye, s.Sira, s.Palet, s.Adet, s.KG)));
                            if (ustSeviye == "")
                            {
                                var exist = rows2.Where(x => x.KokSeviye == kokSeviye && x.UstSeviye == ustSeviye).Count();
                                if (exist > 0)
                                {
                                    var existProduct = rows2.Where(x => x.KokSeviye == kokSeviye && x.Seviye == "3").Count();
                                    if (existProduct > 0)
                                    {
                                        var ddd = rows2.Where(x => x.KokSeviye == kokSeviye && x.Seviye == "3").ToList();

                                        var finalrows = rows2.Except(ddd).ToList();
                                        data2 = string.Join("", finalrows.Select(s => string.Format(row, s.Grup, s.Tur, s.Kod, s.Ad, s.Seviye, s.UstSeviye, s.KokSeviye, s.Sira, s.Palet, s.Adet, s.KG)));
                                    }
                                    else
                                    {
                                        var ddd = OriginalData.Where(x => x.KokSeviye == kokSeviye && x.Seviye == "3").ToList();

                                        rows2.AddRange(ddd);

                                        var finalrows = rows2.OrderBy(x => x.Sira).ToList();


                                        data2 = string.Join("", finalrows.Select(s => string.Format(row, s.Grup, s.Tur, s.Kod, s.Ad, s.Seviye, s.UstSeviye, s.KokSeviye, s.Sira, s.Palet, s.Adet, s.KG)));
                                    }
                                }


                            }
                            else
                            {
                                var exist = rows2.Where(x => x.KokSeviye == kokSeviye && x.UstSeviye == ustSeviye).Count();

                                if (exist > 0)
                                {
                                    var existProduct = rows2.Where(x => x.KokSeviye == kokSeviye && x.UstSeviye == ustSeviye && x.Seviye == "3").Count();

                                    if (existProduct > 0)
                                    {
                                        var ddd = rows2.Where(x => x.KokSeviye == kokSeviye && x.UstSeviye == ustSeviye && x.Seviye == "3").ToList();

                                        var finalrows = rows2.Except(ddd).ToList();
                                        data2 = string.Join("", finalrows.Select(s => string.Format(row, s.Grup, s.Tur, s.Kod, s.Ad, s.Seviye, s.UstSeviye, s.KokSeviye, s.Sira, s.Palet, s.Adet, s.KG)));
                                    }
                                    else
                                    {
                                        var ddd = OriginalData.Where(x => x.KokSeviye == kokSeviye && x.UstSeviye == ustSeviye && x.Seviye == "3").ToList();

                                        rows2.AddRange(ddd);

                                        var finalrows = rows2.OrderBy(x => x.Sira).ToList();


                                        data2 = string.Join("", finalrows.Select(s => string.Format(row, s.Grup, s.Tur, s.Kod, s.Ad, s.Seviye, s.UstSeviye, s.KokSeviye, s.Sira, s.Palet, s.Adet, s.KG)));
                                    }
                                }
                            }

                            frmOperasyonPlani.DataSources.DBDataSources.Item("@AIF_PROD_PLAN1").LoadFromXML(string.Format(header, data2));

                            xml = frmOperasyonPlani.DataSources.DBDataSources.Item(1).GetAsXML();
                            rows2 = (from x in XDocument.Parse(xml).Descendants("row")
                                     select new Datas()
                                     {
                                         Grup = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).First().Value : "",
                                         Tur = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).First().Value : "",
                                         Kod = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).First().Value : "",
                                         Ad = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).First().Value : "",
                                         Seviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Seviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Seviye" select new XElement(y.Element("value"))).First().Value : "",
                                         UstSeviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_UstSeviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_UstSeviye" select new XElement(y.Element("value"))).First().Value : "",
                                         KokSeviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KokSeviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KokSeviye" select new XElement(y.Element("value"))).First().Value : "",
                                         Palet = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Palet" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Palet" select new XElement(y.Element("value"))).First().Value) : 0,
                                         Adet = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adet" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adet" select new XElement(y.Element("value"))).First().Value) : 0,
                                         KG = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KG" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KG" select new XElement(y.Element("value"))).First().Value) : 0,
                                         Sira = x.ElementsBeforeSelf().Count() + 1
                                     }).ToList();

                            var vals = rows2.Where(x => x.Grup != "").ToList();

                            int color = ColorTranslator.ToOle(System.Drawing.Color.Gray);
                            int forecolor = ColorTranslator.ToOle(System.Drawing.Color.White);

                            foreach (var item in vals)
                            {
                                oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                                oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), forecolor);
                            }

                            color = ColorTranslator.ToOle(System.Drawing.Color.DarkGray);

                            vals = rows2.Where(x => x.Tur != "").ToList();

                            foreach (var item in vals)
                            {
                                oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                                oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), -1);
                            }

                            color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                            vals = rows2.Where(x => x.Kod != "").ToList();

                            foreach (var item in vals)
                            {
                                oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                                oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), -1);
                            }

                            renklendir();
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmOperasyonPlani.Freeze(false);
                        }

                    }
                    else if (pVal.ItemUID == "Item_1" && !pVal.BeforeAction)
                    {
                        try
                        {
                            frmOperasyonPlani.Freeze(true);

                            string data2 = string.Join("", OriginalData.Where(x => x.Seviye != "3").Select(s => string.Format(row, s.Grup, s.Tur, s.Kod, s.Ad, s.Seviye, s.UstSeviye, s.KokSeviye, s.Sira, s.Palet, s.Adet, s.KG)));

                            frmOperasyonPlani.DataSources.DBDataSources.Item("@AIF_PROD_PLAN1").LoadFromXML(string.Format(header, data2));


                            var xml = frmOperasyonPlani.DataSources.DBDataSources.Item(1).GetAsXML();

                            var rows2 = (from x in XDocument.Parse(xml).Descendants("row")
                                         select new Datas()
                                         {
                                             Grup = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).First().Value : "",
                                             Tur = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).First().Value : "",
                                             Kod = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).First().Value : "",
                                             Ad = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).First().Value : "",
                                             Seviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Seviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Seviye" select new XElement(y.Element("value"))).First().Value : "",
                                             UstSeviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_UstSeviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_UstSeviye" select new XElement(y.Element("value"))).First().Value : "",
                                             KokSeviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KokSeviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KokSeviye" select new XElement(y.Element("value"))).First().Value : "",
                                             Palet = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Palet" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Palet" select new XElement(y.Element("value"))).First().Value) : 0,
                                             Adet = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adet" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adet" select new XElement(y.Element("value"))).First().Value) : 0,
                                             KG = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KG" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KG" select new XElement(y.Element("value"))).First().Value) : 0,
                                             Sira = x.ElementsBeforeSelf().Count() + 1
                                         }).ToList();

                            var vals = rows2.Where(x => x.Grup != "").ToList();

                            int color = ColorTranslator.ToOle(System.Drawing.Color.Gray);
                            int forecolor = ColorTranslator.ToOle(System.Drawing.Color.White);
                            foreach (var item in vals)
                            {
                                oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                                oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), forecolor);

                            }

                            color = ColorTranslator.ToOle(System.Drawing.Color.DarkGray);

                            vals = rows2.Where(x => x.Tur != "").ToList();

                            foreach (var item in vals)
                            {
                                oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                                oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), -1);
                            }

                            color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                            vals = rows2.Where(x => x.Kod != "").ToList();

                            foreach (var item in vals)
                            {
                                oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                                oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), -1);
                            }
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmOperasyonPlani.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "Item_19" && !pVal.BeforeAction)
                    {
                        try
                        {

                            frmOperasyonPlani.Freeze(true);
                            string data2 = "";
                            data2 = string.Join("", OriginalData.Select(s => string.Format(row, s.Grup, s.Tur, s.Kod, s.Ad, s.Seviye, s.UstSeviye, s.KokSeviye, s.Sira, s.Palet, s.Adet, s.KG)));

                            frmOperasyonPlani.DataSources.DBDataSources.Item("@AIF_PROD_PLAN1").LoadFromXML(string.Format(header, data2));

                            var xml = frmOperasyonPlani.DataSources.DBDataSources.Item(1).GetAsXML();

                            var rows2 = (from x in XDocument.Parse(xml).Descendants("row")
                                         select new Datas()
                                         {
                                             Grup = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).First().Value : "",
                                             Tur = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).First().Value : "",
                                             Kod = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).First().Value : "",
                                             Ad = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).First().Value : "",
                                             Seviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Seviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Seviye" select new XElement(y.Element("value"))).First().Value : "",
                                             UstSeviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_UstSeviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_UstSeviye" select new XElement(y.Element("value"))).First().Value : "",
                                             KokSeviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KokSeviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KokSeviye" select new XElement(y.Element("value"))).First().Value : "",
                                             Palet = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Palet" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Palet" select new XElement(y.Element("value"))).First().Value) : 0,
                                             Adet = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adet" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adet" select new XElement(y.Element("value"))).First().Value) : 0,
                                             KG = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KG" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KG" select new XElement(y.Element("value"))).First().Value) : 0,
                                             Sira = x.ElementsBeforeSelf().Count() + 1
                                         }).ToList();

                            var vals = rows2.Where(x => x.Grup != "").ToList();

                            int color = ColorTranslator.ToOle(System.Drawing.Color.Gray);

                            int forecolor = ColorTranslator.ToOle(System.Drawing.Color.White);
                            foreach (var item in vals)
                            {
                                oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                                oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), forecolor);
                            }

                            color = ColorTranslator.ToOle(System.Drawing.Color.DarkGray);

                            vals = rows2.Where(x => x.Tur != "").ToList();

                            foreach (var item in vals)
                            {
                                oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                                oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), -1);
                            }

                            color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                            vals = rows2.Where(x => x.Kod != "").ToList();

                            foreach (var item in vals)
                            {
                                oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                                oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), -1);
                            }

                            renklendir();
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmOperasyonPlani.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "Item_20" && !pVal.BeforeAction)
                    {
                        GunlukPersonelPlanlama.operasyonplaniTarihi = edtTarih.Value;
                        AIFConn.GunPersPl.LoadForms();
                        GunlukPersonelPlanlama.operasyonplaniTarihi = "";
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
        public void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }
        public class Datas
        {
            public string KokSeviye { get; set; }
            public string UstSeviye { get; set; }
            public string Seviye { get; set; }
            public string Grup { get; set; }
            public string Tur { get; set; }
            public string Kod { get; set; }
            public string Ad { get; set; }
            public double Palet { get; set; }
            public double Adet { get; set; }
            public double KG { get; set; }
            public int Sira { get; set; }


        }
        public class Datas2
        {
            public string KokSeviye { get; set; }
            public string UstSeviye { get; set; }
            public string Seviye { get; set; }
            public string Grup { get; set; }
            public string Tur { get; set; }
            public string Kod { get; set; }
            public string Ad { get; set; }
            public int row { get; set; }
        }

        private void Visible()
        {
            string MatrixXml = oMatrix.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);

            //string xmlRow = @"<row><Visible>{0}</Visible></row>";

            //var rows = (from x in XDocument.Parse(MatrixXml).Descendants("Row")
            //            select new
            //            {
            //                CostName = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
            //                Cost = ParseNumber<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value),
            //            }).ToList();


            //string data = string.Join("", rows.Select(x => string.Format(xmlRow, x.CostName, x.Cost)));
            //oMatrix.
        }

        List<Datas> datas = new List<Datas>();
        private void Listele()
        {
            try
            {
                frmOperasyonPlani.Freeze(true);
                string date = edtTarih.Value;
                datas = new List<Datas>();
                OriginalData = new List<Datas>();

                SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRSSeviye2 = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRSOItem = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                int i = 0;
                oRS.DoQuery("Select * from \"@AIF_SEVIYE1\"");
                oRSSeviye2.DoQuery("Select * from \"@AIF_SEVIYE2\"");
                while (!oRS.EoF)
                {
                    i++;
                    datas.Add(new Datas { Grup = oRS.Fields.Item("U_Name").Value.ToString(), KokSeviye = oRS.Fields.Item("DocEntry").Value.ToString(), Sira = i });
                    while (!oRSSeviye2.EoF)
                    {
                        i++;
                        datas.Add(new Datas { Seviye = oRSSeviye2.Fields.Item("DocEntry").Value.ToString(), Grup = "", Tur = oRSSeviye2.Fields.Item("U_Name").Value.ToString(), UstSeviye = oRSSeviye2.Fields.Item("DocEntry").Value.ToString(), KokSeviye = oRS.Fields.Item("DocEntry").Value.ToString(), Sira = i });

                        oRSOItem.DoQuery("Select \"ItemCode\",\"ItemName\",case when (SELECT TOP 1 Uom FROM \"OWOR\" T1 WHERE T1.\"DueDate\" = '" + date + "' AND  T1.\"ItemCode\" = T0.\"ItemCode\") = 'Adet' then (SELECT SUM(T1.\"PlannedQty\") FROM \"OWOR\" T1 WHERE T1.\"DueDate\" = '" + date + "' AND  T1.\"ItemCode\" = T0.\"ItemCode\") else '0' end as \"Adet\",case when (SELECT TOP 1 Uom FROM \"OWOR\" T1 WHERE T1.\"DueDate\" = '" + date + "' AND  T1.\"ItemCode\" = T0.\"ItemCode\") = 'Litre' then (SELECT SUM(T1.\"PlannedQty\") FROM \"OWOR\" T1 WHERE T1.\"DueDate\" = '" + date + "' AND  T1.\"ItemCode\" = T0.\"ItemCode\") else '0' end as \"KG\" from \"OITM\" AS T0 where \"U_SEVIYE1\" = '" + oRS.Fields.Item("DocEntry").Value.ToString() + "' and \"U_SEVIYE2\" = '" + oRSSeviye2.Fields.Item("DocEntry").Value.ToString() + "' ");

                        while (!oRSOItem.EoF)
                        {
                            i++;
                            datas.Add(new Datas { Seviye = "3", Grup = "", Tur = "", Kod = oRSOItem.Fields.Item("ItemCode").Value.ToString(), Ad = oRSOItem.Fields.Item("ItemName").Value.ToString(), UstSeviye = oRSSeviye2.Fields.Item("DocEntry").Value.ToString(), KokSeviye = oRS.Fields.Item("DocEntry").Value.ToString(), Sira = i, Adet = parservalues<double>(oRSOItem.Fields.Item("Adet").Value.ToString()), KG = parservalues<double>(oRSOItem.Fields.Item("KG").Value.ToString()) });
                            oRSOItem.MoveNext();
                        }
                        //}
                        oRSSeviye2.MoveNext();
                    }

                    oRSSeviye2.MoveFirst();

                    oRS.MoveNext();
                }

                string data2 = string.Join("", datas.Select(s => string.Format(row, s.Grup, s.Tur, s.Kod, s.Ad, s.Seviye, s.UstSeviye, s.KokSeviye, s.Sira, s.Palet, s.Adet, s.KG)));

                frmOperasyonPlani.DataSources.DBDataSources.Item("@AIF_PROD_PLAN1").LoadFromXML(string.Format(header, data2));



                var xml = frmOperasyonPlani.DataSources.DBDataSources.Item(1).GetAsXML();

                var rows2 = (from x in XDocument.Parse(xml).Descendants("row")
                             select new Datas()
                             {
                                 Grup = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Grup" select new XElement(y.Element("value"))).First().Value : "",
                                 Tur = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Tur" select new XElement(y.Element("value"))).First().Value : "",
                                 Kod = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Kod" select new XElement(y.Element("value"))).First().Value : "",
                                 Ad = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adi" select new XElement(y.Element("value"))).First().Value : "",
                                 Seviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Seviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Seviye" select new XElement(y.Element("value"))).First().Value : "",
                                 UstSeviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_UstSeviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_UstSeviye" select new XElement(y.Element("value"))).First().Value : "",
                                 KokSeviye = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KokSeviye" select new XElement(y.Element("value"))).Any() == true ? (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KokSeviye" select new XElement(y.Element("value"))).First().Value : "",
                                 Palet = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Palet" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Palet" select new XElement(y.Element("value"))).First().Value) : 0,
                                 Adet = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adet" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_Adet" select new XElement(y.Element("value"))).First().Value) : 0,
                                 KG = (from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KG" select new XElement(y.Element("value"))).Any() == true ? parservalues<double>((from y in x.Elements("cells").Elements("cell") where y.Element("uid").Value == "U_KG" select new XElement(y.Element("value"))).First().Value) : 0,
                                 Sira = x.ElementsBeforeSelf().Count() + 1
                             }).ToList();

                OriginalData.AddRange(rows2);

                var empty = rows2.Where(x => x.Grup != "" && x.Tur != "").ToList();

                empty.ForEach(x => x.Grup = "");


                empty = rows2.Where(x => x.Tur != "" && x.Kod != "").ToList();
                empty.ForEach(x => x.Tur = "");

                var vals = rows2.Where(x => x.Grup != "").ToList();

                int color = ColorTranslator.ToOle(System.Drawing.Color.Gray);
                int forecolor = ColorTranslator.ToOle(System.Drawing.Color.White);

                foreach (var item in vals)
                {
                    oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                    oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), forecolor);
                }

                color = ColorTranslator.ToOle(System.Drawing.Color.DarkGray);

                vals = rows2.Where(x => x.Tur != "").ToList();

                foreach (var item in vals)
                {
                    oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                    oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), -1);
                }

                color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                vals = rows2.Where(x => x.Kod != "").ToList();

                foreach (var item in vals)
                {
                    oMatrix.CommonSetting.SetRowBackColor(Convert.ToInt32(item.Sira), color);
                    oMatrix.CommonSetting.SetRowFontColor(Convert.ToInt32(item.Sira), -1);
                }


                renklendir();


                oMatrix.AutoResizeColumns();
                return;
                oRS = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string qry = "Select Case When T0.\"QryGroup2\" = 'Y' then 'Mamül' when T0.\"QryGroup3\" = 'Y' then 'Yarımamül' else '' end as \"Tur\",T0.\"ItemCode\",T0.\"ItemName\",case when (SELECT TOP 1 Uom FROM \"OWOR\" T1 WHERE T1.\"DueDate\" = '" + date + "' AND  T1.\"ItemCode\" = T0.\"ItemCode\") = 'Adet' then (SELECT SUM(T1.\"PlannedQty\") FROM \"OWOR\" T1 WHERE T1.\"DueDate\" = '" + date + "' AND  T1.\"ItemCode\" = T0.\"ItemCode\") else '0' end as \"Adet\",case when (SELECT TOP 1 Uom FROM \"OWOR\" T1 WHERE T1.\"DueDate\" = '" + date + "' AND  T1.\"ItemCode\" = T0.\"ItemCode\") = 'Litre' then (SELECT SUM(T1.\"PlannedQty\") FROM \"OWOR\" T1 WHERE T1.\"DueDate\" = '" + date + "' AND  T1.\"ItemCode\" = T0.\"ItemCode\") else '0' end as \"Litre\",T0.\"SWW\" as \"Grup\" from OITM as T0 where (T0.\"QryGroup2\" = 'Y' or  T0.\"QryGroup3\" = 'Y') ";

                //string qry = "Select ItemCode,ItemName,Tur,Grup,SUM(\"Adet\") as \"Adet\",SUM(\"Litre\") as \"Litre\" from(Select Case When T0.\"QryGroup2\" = 'Y' then 'Mamül' when T0.\"QryGroup3\" = 'Y' then 'Yarımamül' else '' end as \"Tur\",T0.\"ItemCode\",T0.\"ItemName\",case when (SELECT TOP 1 Uom FROM \"OWOR\" T1 WHERE T1.\"DueDate\" = '" + date + "' AND  T1.\"ItemCode\" = T0.\"ItemCode\") = 'Adet' then (SELECT SUM(T1.\"PlannedQty\") FROM \"OWOR\" T1 WHERE T1.\"DueDate\" = '" + date + "' AND  T1.\"ItemCode\" = T0.\"ItemCode\") else '0' end as \"Adet\",case when (SELECT TOP 1 Uom FROM \"OWOR\" T1 WHERE T1.\"DueDate\" = '" + date + "' AND  T1.\"ItemCode\" = T0.\"ItemCode\") = 'Litre' then (SELECT SUM(T1.\"PlannedQty\") FROM \"OWOR\" T1 WHERE T1.\"DueDate\" = '" + date + "' AND  T1.\"ItemCode\" = T0.\"ItemCode\") else '0' end as \"Litre\",T0.\"SWW\" as \"Grup\" from OITM as T0 where (T0.\"QryGroup2\" = 'Y' or  T0.\"QryGroup3\" = 'Y')) as tbl GROUP BY tbl.\"ItemCode\",tbl.\"ItemName\",tbl.\"Grup\",tbl.\"Tur\" ";
                oRS.DoQuery(qry);


                string xmll = oRS.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                XDocument xDoc = XDocument.Parse(xmll);
                XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                var rows = (from t in xDoc.Descendants(ns + "Row")
                            select new
                            {
                                ItemCode = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "ItemCode" select new XElement(y.Element(ns + "Value"))).First().Value,
                                ItemName = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "ItemName" select new XElement(y.Element(ns + "Value"))).First().Value,
                                Tur = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "Tur" select new XElement(y.Element(ns + "Value"))).First().Value,
                                Adet = parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "Adet" select new XElement(y.Element(ns + "Value"))).First().Value),
                                Kilogram = parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "Litre" select new XElement(y.Element(ns + "Value"))).First().Value),
                                Grup = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "Grup" select new XElement(y.Element(ns + "Value"))).First().Value,
                            }).OrderBy(y => y.Tur).ToList();



                string data = string.Join("", rows.Select(s => string.Format(row, s.Tur, s.ItemCode, s.ItemName, s.Adet, s.Kilogram, s.Grup)));

                frmOperasyonPlani.DataSources.DBDataSources.Item("@AIF_PROD_PLAN1").LoadFromXML(string.Format(header, data));

                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
            }
            finally
            {
                frmOperasyonPlani.Freeze(false);
            }
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
    }
}
