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
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class GunlukPersonelPlanlama3
    {
        [ItemAtt(AIFConn.GunlukPersonelPlan3UID)]
        public SAPbouiCOM.Form frmGunlukPersonelPlanlama3;

        [ItemAtt("Item_2")]
        public SAPbouiCOM.Matrix oMatrix;

        [ItemAtt("Item_8")]
        public SAPbouiCOM.EditText oEditArama;

        [ItemAtt("Item_1")]
        public SAPbouiCOM.EditText oEditTarih;

        [ItemAtt("Item_6")]
        public SAPbouiCOM.EditText oEditDocEntry;

        [ItemAtt("1")]
        public SAPbouiCOM.Button btnAddOrUpdate;

        [ItemAtt("Item_10")]
        public SAPbouiCOM.Button btnClean;
        //[ItemAtt("1")]
        //public SAPbouiCOM.Button btnAddOrUpdate;

        private string header = @"<?xml version=""1.0"" encoding=""UTF-16"" ?><dbDataSources uid=""@AIF_GUNLUKPLAN1""><rows>{0}</rows></dbDataSources>";

        public static string operasyonplaniTarihi = "";

        private string row = "<row><cells><cell><uid>U_PersonelNo</uid><value>{0}</value></cell><cell><uid>U_PersonelAdi</uid><value>{1}</value></cell><cell><uid>U_Bolum1</uid><value>{2}</value></cell><cell><uid>U_Bolum2</uid><value>{3}</value></cell><cell><uid>U_Bolum3</uid><value>{4}</value></cell><cell><uid>U_Durum</uid><value>{5}</value></cell></cells></row>";

        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.GunlukPersonelPlan3XML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.GunlukPersonelPlan3XML));
            Functions.CreateUserOrSystemFormComponent<GunlukPersonelPlanlama3>(AIFConn.PrsPlan);

            InitForms();
        }
        public void InitForms()
        {
            try
            {
                frmGunlukPersonelPlanlama3.Freeze(true);
                frmGunlukPersonelPlanlama3.EnableMenu("1283", false);
                frmGunlukPersonelPlanlama3.EnableMenu("1284", false);
                frmGunlukPersonelPlanlama3.EnableMenu("1286", false);
                CalisanListele();
                oMatrix.AutoResizeColumns();

                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql = "Select \"U_BolumKodu\",\"U_BolumAdi\" from \"@AIF_BOLUMLER\"";

                ConstVariables.oRecordset.DoQuery(sql);
                SAPbouiCOM.Column oColumn;
                SAPbouiCOM.Column oColumn1;
                SAPbouiCOM.Column oColumn2;

                oColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item("Col_1");
                oColumn1 = (SAPbouiCOM.Column)oMatrix.Columns.Item("Col_2");
                oColumn2 = (SAPbouiCOM.Column)oMatrix.Columns.Item("Col_3");
                while (!ConstVariables.oRecordset.EoF)
                {
                    oColumn.ValidValues.Add(ConstVariables.oRecordset.Fields.Item(0).Value.ToString(), ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                    oColumn1.ValidValues.Add(ConstVariables.oRecordset.Fields.Item(0).Value.ToString(), ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                    oColumn2.ValidValues.Add(ConstVariables.oRecordset.Fields.Item(0).Value.ToString(), ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                    ConstVariables.oRecordset.MoveNext();
                }


                frmGunlukPersonelPlanlama3.EnableMenu("5890", true);
            }
            catch (Exception ex)
            {
            }
            finally
            {
                frmGunlukPersonelPlanlama3.Freeze(false);
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
                    if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    {
                        try
                        {
                            string sql = "";
                            string sql2 = "";
                            foreach (var item in gunlukPersonelPlanlama2s)
                            {
                                DateTime tarih = new DateTime(Convert.ToInt32(item.Tarih.ToString().Substring(0, 4)), Convert.ToInt32(item.Tarih.ToString().Substring(4, 2)), Convert.ToInt32(item.Tarih.ToString().Substring(6, 2)));
                                sql = "Select \"DocEntry\" from \"@AIF_GUNLUKPERSPLAN\" where \"U_Aylar\"='" + tarih.Month.ToString().PadLeft(2, '0') + "' and \"U_Yil\" = '" + tarih.Year + "'";

                                ConstVariables.oRecordset.DoQuery(sql);

                                if (ConstVariables.oRecordset.RecordCount > 0)
                                {
                                    var field = "U_Gun" + tarih.Day;
                                    sql2 = "UPDATE \"@AIF_GUNLUKPERSPLAN1\" SET " + field + " = '" + item.Durum + "' where \"U_PersonelNo\" = '" + item.KullaniciId + "' and \"DocEntry\" = '" + ConstVariables.oRecordset.Fields.Item(0).Value + "'";
                                    ConstVariables.oRecordset.DoQuery(sql2);
                                }
                            }

                            if (oMatrix.RowCount == 0)
                            {
                                CalisanListele();
                            }
                            else
                            {
                                renklendir();
                            }

                            gunlukPersonelPlanlama2s = new List<GunlukPersonelPlanlama2>();
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
                    if (pVal.ItemUID == "Item_1" && !pVal.BeforeAction)
                    {
                        try
                        {
                            if (oEditTarih == null)
                            {
                                return false;
                            }
                            ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
                            ConstVariables.oRecordset.DoQuery("Select * from \"@AIF_GUNLUKPLAN\" where CONVERT(varchar, \"U_Tarih\", 112)  = '" + oEditTarih.Value + "' ");


                            if (ConstVariables.oRecordset.RecordCount > 0)
                            {
                                frmGunlukPersonelPlanlama3.Mode = BoFormMode.fm_FIND_MODE;
                                oEditDocEntry.Value = ConstVariables.oRecordset.Fields.Item("DocEntry").Value.ToString();
                                btnAddOrUpdate.Item.Click();
                                oEditArama.Item.Click();
                                oEditTarih.Item.Enabled = false;

                                if (operasyonplaniTarihi == "")
                                {
                                    ConstVariables.oRecordset.DoQuery("Select case when ISNULL(\"middleName\",'') = '' then  (\"firstName\" + ' ' + \"lastName\") else (\"firstName\" + ' ' + \"middleName\" + ' ' + \"lastName\") end as \"AdSoyad\",\"empID\"  from \"OHEM\" where \"Active\" = 'Y' and \"U_UretimCalisani\" = 'Evet'");

                                    string xmll = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                                    XDocument xDoc = XDocument.Parse(xmll);
                                    XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                                    var rows = (from tx in xDoc.Descendants(ns + "Row")
                                                select new
                                                {
                                                    AdSoyad = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "AdSoyad" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                    Id = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "empID" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                }).ToList();




                                    string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                    var rowsMatrix = (from x in XDocument.Parse(xml).Descendants("Row")
                                                      select new
                                                      {
                                                          AdSoyad = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "#" select new XElement(y.Element("Value"))).First().Value,
                                                          Id = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                                      }).ToList();

                                    var FinalXml = rows.Except(rowsMatrix).ToList();

                                    //rowsMatrix.AddRange(FinalXml);


                                    foreach (var item in FinalXml)
                                    {
                                        frmGunlukPersonelPlanlama3.DataSources.DBDataSources.Item(1).Clear();
                                        oMatrix.AddRow();
                                        //oMatrix.ClearRowData(oMatrix.RowCount);
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("#").Cells.Item(oMatrix.RowCount).Specific).Value = item.AdSoyad;
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific).Value = item.Id;
                                    }

                                    //if (FinalXml.Count > 0)
                                    //{
                                    //}

                                    oMatrix.AutoResizeColumns();


                                    #region Yeni kullanıcı eklenirse buradan gelir.
                                    if (FinalXml.Count > 0)
                                    {
                                        DateTime tarih = new DateTime(Convert.ToInt32(this.oEditTarih.Value.Substring(0, 4)), Convert.ToInt32(this.oEditTarih.Value.Substring(4, 2)), Convert.ToInt32(this.oEditTarih.Value.Substring(6, 2)));

                                        string field = "U_Gun" + tarih.Day;
                                        string sql = "Select " + field + " as \"Gun\",\"U_PersonelNo\" from \"@AIF_GUNLUKPERSPLAN\" as T0 INNER JOIN \"@AIF_GUNLUKPERSPLAN1\" as T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where  \"U_Aylar\" = '" + tarih.Month.ToString().PadLeft(2, '0') + "' and \"U_Yil\" = '" + tarih.Year + "' ";
                                        ConstVariables.oRecordset.DoQuery(sql);

                                        xmll = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                                        xDoc = XDocument.Parse(xmll);
                                        ns = "http://www.sap.com/SBO/SDK/DI";
                                        var rowsGunlukPersonelPlanlama2 = (from tx in xDoc.Descendants(ns + "Row")
                                                                           select new
                                                                           {
                                                                               Gun = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "Gun" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                                               Id = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_PersonelNo" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                                           }).ToList();


                                        if (rowsGunlukPersonelPlanlama2.Count > 0)
                                        {

                                            var rowsFinal = (from x in XDocument.Parse(xml).Descendants("Row")
                                                             select new MatrixDetay()
                                                             {
                                                                 AdSoyad = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "#" select new XElement(y.Element("Value"))).First().Value,
                                                                 Id = Convert.ToInt32((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value),
                                                                 Bolum1 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                                                                 Bolum2 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value,
                                                                 Bolum3 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value,
                                                                 Bolum1Adi = "",
                                                                 Bolum2Adi = "",
                                                                 Bolum3Adi = "",
                                                                 Durum = rowsGunlukPersonelPlanlama2.Where(z => z.Id == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.Gun).Count() > 0 ? rowsGunlukPersonelPlanlama2.Where(z => z.Id == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.Gun).FirstOrDefault() : (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,

                                                             }).ToList();

                                            string data = "";

                                            data = string.Join("", rowsFinal.Select(s => string.Format(row, s.Id, s.AdSoyad, s.Bolum1, s.Bolum2, s.Bolum3, s.Durum)));

                                            string s1 = string.Format(header, data);
                                            frmGunlukPersonelPlanlama3.DataSources.DBDataSources.Item("@AIF_GUNLUKPLAN1").LoadFromXML(string.Format(header, data));

                                            oMatrix.AutoResizeColumns();

                                        }
                                    }
                                    #endregion


                                    renklendir();

                                    MatrisXMLiniListeyeYukle();

                                }
                            }
                            else
                            {
                                string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                DateTime tarih = new DateTime(Convert.ToInt32(this.oEditTarih.Value.Substring(0, 4)), Convert.ToInt32(this.oEditTarih.Value.Substring(4, 2)), Convert.ToInt32(this.oEditTarih.Value.Substring(6, 2)));

                                string field = "U_Gun" + tarih.Day;
                                string sql = "Select " + field + " as \"Gun\",\"U_PersonelNo\",\"U_Bolum1\",\"U_Bolum2\",\"U_Bolum3\" from \"@AIF_GUNLUKPERSPLAN\" as T0 INNER JOIN \"@AIF_GUNLUKPERSPLAN1\" as T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where  \"U_Aylar\" = '" + tarih.Month.ToString().PadLeft(2, '0') + "' and \"U_Yil\" = '" + tarih.Year + "' ";
                                ConstVariables.oRecordset.DoQuery(sql);

                                string xmll = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                                XDocument xDoc = XDocument.Parse(xmll);
                                XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                                var rowsGunlukPersonelPlanlama2 = (from tx in xDoc.Descendants(ns + "Row")
                                                                   select new
                                                                   {
                                                                       Gun = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "Gun" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                                       Id = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_PersonelNo" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                                       Bolum1 = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Bolum1" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                                       Bolum2 = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Bolum2" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                                       Bolum3 = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Bolum3" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                                   }).ToList();


                                if (rowsGunlukPersonelPlanlama2.Count > 0)
                                {

                                    var rowsFinal = (from x in XDocument.Parse(xml).Descendants("Row")
                                                     select new MatrixDetay()
                                                     {
                                                         AdSoyad = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "#" select new XElement(y.Element("Value"))).First().Value,
                                                         Id = Convert.ToInt32((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value),
                                                         Bolum1 = rowsGunlukPersonelPlanlama2.Where(z => z.Id == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.Bolum1).Count() > 0 ? rowsGunlukPersonelPlanlama2.Where(z => z.Id == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.Bolum1).FirstOrDefault() : "",
                                                         Bolum2 = rowsGunlukPersonelPlanlama2.Where(z => z.Id == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.Bolum2).Count() > 0 ? rowsGunlukPersonelPlanlama2.Where(z => z.Id == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.Bolum2).FirstOrDefault() : "",
                                                         Bolum3 = rowsGunlukPersonelPlanlama2.Where(z => z.Id == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.Bolum3).Count() > 0 ? rowsGunlukPersonelPlanlama2.Where(z => z.Id == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.Bolum3).FirstOrDefault() : "",
                                                         Bolum1Adi = "",
                                                         Bolum2Adi = "",
                                                         Bolum3Adi = "",
                                                         Durum = rowsGunlukPersonelPlanlama2.Where(z => z.Id == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.Gun).Count() > 0 ? rowsGunlukPersonelPlanlama2.Where(z => z.Id == (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value).Select(y => y.Gun).FirstOrDefault() : (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,

                                                     }).ToList();

                                    string data = "";

                                    data = string.Join("", rowsFinal.Select(s => string.Format(row, s.Id, s.AdSoyad, s.Bolum1, s.Bolum2, s.Bolum3, s.Durum)));

                                    string s1 = string.Format(header, data);
                                    frmGunlukPersonelPlanlama3.DataSources.DBDataSources.Item("@AIF_GUNLUKPLAN1").LoadFromXML(string.Format(header, data));

                                    oMatrix.AutoResizeColumns();

                                }
                                MatrisXMLiniListeyeYukle();
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    if (!pVal.BeforeAction && (pVal.ColUID == "Col_1" || pVal.ColUID == "Col_2" || pVal.ColUID == "Col_3") || pVal.ColUID == "Col_4")
                    {
                        try
                        {
                            var id = Convert.ToInt32(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value);
                            string value = "";
                            value = StringReplace(((SAPbouiCOM.ComboBox)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).Value.ToString());
                            string calismaDurumu = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific).Value.Trim();

                            gunlukPersonelPlanlama2s.RemoveAll(x => x.KullaniciId == id);

                            gunlukPersonelPlanlama2s.Add(new GunlukPersonelPlanlama2 { KullaniciId = id, Durum = calismaDurumu, Tarih = this.oEditTarih.Value });


                            #region Güncellemeler
                            if (pVal.ColUID == "Col_1")
                            {
                                MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.Bolum1 = value);
                                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
                                string sql = "";
                                sql = "Select \"U_BolumKodu\",\"U_BolumAdi\" from \"@AIF_BOLUMLER\" where  \"U_BolumKodu\"='" + value + "'";

                                ConstVariables.oRecordset.DoQuery(sql);

                                MatrixOrjinalData.Where(x => x.Bolum1 == value).ToList().ForEach(y => y.Bolum1Adi = ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                            }
                            else if (pVal.ColUID == "Col_2")
                            {
                                MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.Bolum2 = value);

                                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
                                string sql = "";
                                sql = "Select \"U_BolumKodu\",\"U_BolumAdi\" from \"@AIF_BOLUMLER\" where  \"U_BolumKodu\"='" + value + "'";

                                ConstVariables.oRecordset.DoQuery(sql);
                                MatrixOrjinalData.Where(x => x.Bolum2 == value).ToList().ForEach(y => y.Bolum2Adi = ConstVariables.oRecordset.Fields.Item(1).Value.ToString());

                            }
                            else if (pVal.ColUID == "Col_3")
                            {
                                MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.Bolum3 = value);

                                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
                                string sql = "";
                                sql = "Select \"U_BolumKodu\",\"U_BolumAdi\" from \"@AIF_BOLUMLER\" where  \"U_BolumKodu\"='" + value + "'";

                                ConstVariables.oRecordset.DoQuery(sql);
                                MatrixOrjinalData.Where(x => x.Bolum3 == value).ToList().ForEach(y => y.Bolum3Adi = ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                            }
                            else if (pVal.ColUID == "Col_4")
                                MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.Durum = value);



                            #endregion
                        }
                        catch (Exception)
                        {
                        }
                    }
                    break;
                case BoEventTypes.et_CLICK:
                    if (pVal.ItemUID == "Item_9" && !pVal.BeforeAction)
                    {
                        try
                        {
                            var value = StringReplace(oEditArama.Value.ToUpper());

                            frmGunlukPersonelPlanlama3.Freeze(true);

                            List<MatrixDetay> rows = new List<MatrixDetay>();
                            bool check = value.IsNumeric();

                            if (check)
                            {
                                rows = MatrixOrjinalData.Where(x => StringReplace(x.AdSoyad.ToUpper()).Contains(value) || x.Id == (Convert.ToInt32(value)) || StringReplace(x.Bolum1Adi.ToUpper()).Contains(value) || StringReplace(x.Bolum2Adi.ToUpper()).Contains(value) || StringReplace(x.Bolum3Adi.ToUpper()).Contains(value)).ToList();
                            }
                            else
                            {
                                rows = MatrixOrjinalData.Where(x => StringReplace(x.AdSoyad.ToUpper()).Contains(value) || StringReplace(x.Bolum1Adi.ToUpper()).Contains(value) || StringReplace(x.Bolum2Adi.ToUpper()).Contains(value) || StringReplace(x.Bolum3Adi.ToUpper()).Contains(value)).ToList();
                            }



                            string data = string.Join("", rows.Select(s => string.Format(row, s.Id, s.AdSoyad, s.Bolum1, s.Bolum2, s.Bolum3, s.Durum)));

                            string s1 = string.Format(header, data);
                            frmGunlukPersonelPlanlama3.DataSources.DBDataSources.Item("@AIF_GUNLUKPLAN1").LoadFromXML(string.Format(header, data));

                            oMatrix.AutoResizeColumns();
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmGunlukPersonelPlanlama3.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "Item_10" && !pVal.BeforeAction)
                    {
                        try
                        {
                            oEditArama.Value = "";
                            var value = StringReplace(oEditArama.Value.ToUpper());

                            frmGunlukPersonelPlanlama3.Freeze(true);

                            var rows = MatrixOrjinalData.ToList();

                            string data = string.Join("", rows.Select(s => string.Format(row, s.Id, s.AdSoyad, s.Bolum1, s.Bolum2, s.Bolum3, s.Durum)));

                            string s1 = string.Format(header, data);
                            frmGunlukPersonelPlanlama3.DataSources.DBDataSources.Item("@AIF_GUNLUKPLAN1").LoadFromXML(string.Format(header, data));

                            oMatrix.AutoResizeColumns();
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmGunlukPersonelPlanlama3.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "Item_11" && !pVal.BeforeAction)
                    {
                        ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        ConstVariables.oRecordset.DoQuery("Select case when ISNULL(\"middleName\",'') = '' then  (\"firstName\" + ' ' + \"lastName\") else (\"firstName\" + ' ' + \"middleName\" + ' ' + \"lastName\") end as \"AdSoyad\",\"empID\"  from \"OHEM\" where \"Active\" = 'Y' and \"U_UretimCalisani\" = 'Evet'");

                        string xmll = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                        XDocument xDoc = XDocument.Parse(xmll);
                        XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                        var OHEMList = (from t in xDoc.Descendants(ns + "Row")
                                        select new
                                        {
                                            AdSoyad = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "AdSoyad" select new XElement(y.Element(ns + "Value"))).First().Value,
                                            Id = Convert.ToInt32((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "empID" select new XElement(y.Element(ns + "Value"))).First().Value),
                                        }).ToList();



                        string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                        var rowsMatrix = (from x in XDocument.Parse(xml).Descendants("Row")
                                          select new
                                          {
                                              AdSoyad = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "#" select new XElement(y.Element("Value"))).First().Value,
                                              Id = Convert.ToInt32((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value),
                                          }).ToList();

                        var finalRows = rowsMatrix.Except(OHEMList).ToList();
                        var rows = MatrixOrjinalData.OrderBy(y => y.Id).ToList();

                        foreach (var item in finalRows)
                        {
                            rows.RemoveAll(x => x.Id == item.Id);
                            MatrixOrjinalData.RemoveAll(x => x.Id == item.Id);
                        }


                        var ohemdevar = OHEMList.Except(rowsMatrix).ToList();

                        foreach (var item in ohemdevar)
                        {
                            rows.Add(new MatrixDetay { Id = item.Id, AdSoyad = item.AdSoyad });
                        }


                        rows = rows.OrderBy(y => y.Id).ToList();

                        string data = string.Join("", rows.Select(s => string.Format(row, s.Id, s.AdSoyad, s.Bolum1, s.Bolum2, s.Bolum3, s.Durum)));

                        string s1 = string.Format(header, data);
                        frmGunlukPersonelPlanlama3.DataSources.DBDataSources.Item("@AIF_GUNLUKPERSPLAN1").LoadFromXML(string.Format(header, data));

                        oMatrix.AutoResizeColumns();



                        //renklendir();

                        //Handler.SAPApplication.StatusBar.SetText("Listeleme tamamlandı.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);


                        if (frmGunlukPersonelPlanlama3.Mode != BoFormMode.fm_ADD_MODE)
                        {
                            frmGunlukPersonelPlanlama3.Mode = BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                    else if (pVal.ItemUID == "1" && pVal.BeforeAction)
                    {
                        var tarih = oEditTarih.Value;

                        if (tarih == "")
                        {
                            Handler.SAPApplication.MessageBox("Tarih alanı boş geçilemez.");
                        }

                        oEditArama.Value = "";
                        btnClean.Item.Click();
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
            if (pVal.MenuUID == "1282" && !pVal.BeforeAction)
            {
                oEditTarih.Item.Enabled = true;
                oEditArama.Value = "";
                oEditTarih.Item.Click();
                CalisanListele();
            }
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }

        private void CalisanListele()
        {
            try
            {
                frmGunlukPersonelPlanlama3.Freeze(true);
                if (operasyonplaniTarihi == "")
                {
                    ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    ConstVariables.oRecordset.DoQuery("Select case when ISNULL(\"middleName\",'') = '' then  (\"firstName\" + ' ' + \"lastName\") else (\"firstName\" + ' ' + \"middleName\" + ' ' + \"lastName\") end as \"AdSoyad\",\"empID\"  from \"OHEM\" where \"Active\" = 'Y' and \"U_UretimCalisani\" = 'Evet'");

                    string xmll = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                    XDocument xDoc = XDocument.Parse(xmll);
                    XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                    var rows = (from t in xDoc.Descendants(ns + "Row")
                                select new
                                {
                                    AdSoyad = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "AdSoyad" select new XElement(y.Element(ns + "Value"))).First().Value,
                                    Id = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "empID" select new XElement(y.Element(ns + "Value"))).First().Value,
                                }).ToList();

                    string data = string.Join("", rows.Select(s => string.Format(row, s.Id, s.AdSoyad, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")));

                    string s1 = string.Format(header, data);
                    frmGunlukPersonelPlanlama3.DataSources.DBDataSources.Item("@AIF_GUNLUKPLAN1").LoadFromXML(string.Format(header, data));

                    oMatrix.AutoResizeColumns();

                    MatrisXMLiniListeyeYukle();
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                frmGunlukPersonelPlanlama3.Freeze(false);
            }
        }

        public class MatrixDetay
        {
            public string AdSoyad { get; set; }

            public int Id { get; set; }

            public string Bolum1 { get; set; }

            public string Bolum2 { get; set; }

            public string Bolum3 { get; set; }

            public string Bolum1Adi { get; set; }

            public string Bolum2Adi { get; set; }

            public string Bolum3Adi { get; set; }

            public string Durum { get; set; }
        }

        public string StringReplace(string text)
        {
            text = text.Replace("İ", "I");
            text = text.Replace("ı", "I");
            text = text.Replace("Ğ", "G");
            text = text.Replace("ğ", "G");
            text = text.Replace("Ö", "O");
            text = text.Replace("ö", "O");
            text = text.Replace("Ü", "U");
            text = text.Replace("ü", "U");
            text = text.Replace("Ş", "S");
            text = text.Replace("ş", "S");
            text = text.Replace("Ç", "C");
            text = text.Replace("ç", "C");
            text = text.Replace(" ", "_");
            return text;
        }

        List<MatrixDetay> MatrixOrjinalData = new List<MatrixDetay>();


        private void MatrisXMLiniListeyeYukle()
        {

            var xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

            MatrixOrjinalData = (from x in XDocument.Parse(xml).Descendants("Row")
                                 select new MatrixDetay()
                                 {
                                     AdSoyad = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "#" select new XElement(y.Element("Value"))).First().Value,
                                     Id = Convert.ToInt32((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value),
                                     Bolum1 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                                     Bolum2 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value,
                                     Bolum3 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value,
                                     Bolum1Adi = "",
                                     Bolum2Adi = "",
                                     Bolum3Adi = "",
                                     Durum = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,

                                 }).ToList();

            ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
            string sql = "";
            var Bolum1Distinct = MatrixOrjinalData.Select(x => x.Bolum1).Distinct().ToList();
            var Bolum2Distinct = MatrixOrjinalData.Select(x => x.Bolum2).Distinct().ToList();
            var Bolum3Distinct = MatrixOrjinalData.Select(x => x.Bolum3).Distinct().ToList();

            foreach (var item in Bolum1Distinct)
            {
                sql = "Select \"U_BolumKodu\",\"U_BolumAdi\" from \"@AIF_BOLUMLER\" where  \"U_BolumKodu\"='" + item + "'";

                ConstVariables.oRecordset.DoQuery(sql);

                MatrixOrjinalData.Where(x => x.Bolum1 == item).ToList().ForEach(y => y.Bolum1Adi = StringReplace(ConstVariables.oRecordset.Fields.Item(1).Value.ToString()));
            }

            foreach (var item in Bolum2Distinct)
            {
                sql = "Select \"U_BolumKodu\",\"U_BolumAdi\" from \"@AIF_BOLUMLER\" where  \"U_BolumKodu\"='" + item + "'";

                ConstVariables.oRecordset.DoQuery(sql);

                MatrixOrjinalData.Where(x => x.Bolum2 == item).ToList().ForEach(y => y.Bolum2Adi = StringReplace(ConstVariables.oRecordset.Fields.Item(1).Value.ToString()));
            }

            foreach (var item in Bolum3Distinct)
            {
                sql = "Select \"U_BolumKodu\",\"U_BolumAdi\" from \"@AIF_BOLUMLER\" where  \"U_BolumKodu\"='" + item + "'";

                ConstVariables.oRecordset.DoQuery(sql);

                MatrixOrjinalData.Where(x => x.Bolum3 == item).ToList().ForEach(y => y.Bolum3Adi = StringReplace(ConstVariables.oRecordset.Fields.Item(1).Value.ToString()));
            }


        }


        private void renklendir(bool addbtn = false)
        {
            try
            {
                frmGunlukPersonelPlanlama3.Freeze(true);

                //for (int i = 1; i <= oMatrix.RowCount; i++)
                //{
                //    oMatrix.CommonSetting.SetRowBackColor(i, -1);
                //}
                //if (!addbtn)
                //{
                string xmlColor = oMatrix.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);
                XDocument xDoc = XDocument.Parse(xmlColor);
                var table = (from x in xDoc.Elements("Matrix").Elements("Rows").Elements("Row")
                             select new
                             {
                                 Durum = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,
                                 INDEX = x.ElementsBeforeSelf().Count() + 1,

                             }).ToList();




                Color val1 = Color.Green;
                Color val2 = Color.Red;
                Color val3 = Color.Blue;
                Color val4 = Color.LightGoldenrodYellow;
                Color val5 = Color.LightGoldenrodYellow;
                Color val6 = Color.Silver;
                int Yesil = ColorTranslator.ToOle(val1);
                int Kirmizi = ColorTranslator.ToOle(val2);
                int Mavi = ColorTranslator.ToOle(val3);
                int Sari = ColorTranslator.ToOle(val4);
                int Turuncu = ColorTranslator.ToOle(val5);
                int Gri = ColorTranslator.ToOle(val6);

                #region Çalışanlar Yeşile Çevir.
                var customrenk = table.Where(x => x.Durum == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetRowBackColor(item.INDEX, Yesil);
                }


                #endregion

                #region Hafta Tatili Kırmızı Çevir.
                customrenk = table.Where(x => x.Durum == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetRowBackColor(item.INDEX, Kirmizi);
                }

                #endregion

                #region Raporluları Maviye Çevir.
                customrenk = table.Where(x => x.Durum == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetRowBackColor(item.INDEX, Mavi);
                }


                #endregion

                #region Ücretsiz İzinlileri Sarıya Çevir.
                customrenk = table.Where(x => x.Durum == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetRowBackColor(item.INDEX, Sari);
                }

                #endregion

                #region Resmi Tatilleri SarıTuruncuya Çevir.
                customrenk = table.Where(x => x.Durum == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetRowBackColor(item.INDEX, Turuncu);
                }

                #endregion

                #region Yıllık İzinleri Griye Çevir.
                customrenk = table.Where(x => x.Durum == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetRowBackColor(item.INDEX, Gri);
                }

                #endregion

                xmlColor = oMatrix.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);
            }
            catch (Exception)
            {
            }
            finally
            {
                frmGunlukPersonelPlanlama3.Freeze(false);
            }
        }

        List<GunlukPersonelPlanlama2> gunlukPersonelPlanlama2s = new List<GunlukPersonelPlanlama2>();
        public class GunlukPersonelPlanlama2
        {
            public int KullaniciId { get; set; }

            public string Tarih { get; set; }

            public string Durum { get; set; }
        }
    }
}
