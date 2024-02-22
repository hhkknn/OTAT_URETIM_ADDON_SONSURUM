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
    public class GunlukPersonelPlanlama2
    {
        [ItemAtt(AIFConn.GunlukPersonelPlan2UID)]
        public SAPbouiCOM.Form frmGunlukPersonelPlanlama2;

        [ItemAtt("Item_2")]
        public SAPbouiCOM.Matrix oMatrix;

        [ItemAtt("1")]
        public SAPbouiCOM.Button btnAddOrUpdate;

        [ItemAtt("Item_9")]
        public SAPbouiCOM.Button btnUpdatePersoneller;

        [ItemAtt("Item_6")]
        public SAPbouiCOM.EditText oEditArama;

        //[ItemAtt("Item_1")]
        //public SAPbouiCOM.EditText oEditTariheKopyala;

        [ItemAtt("Item_4")]
        public SAPbouiCOM.EditText oEditDocEntry;


        [ItemAtt("Item_1")]
        public SAPbouiCOM.ComboBox oComboAy;

        [ItemAtt("Item_8")]
        public SAPbouiCOM.ComboBox oComboYil;

        [ItemAtt("Item_11")]
        public SAPbouiCOM.Button btnTemizle;

        private string header = @"<?xml version=""1.0"" encoding=""UTF-16"" ?><dbDataSources uid=""@AIF_GUNLUKPERSPLAN1""><rows>{0}</rows></dbDataSources>";

        public static string operasyonplaniTarihi = "";

        private string row = "<row><cells><cell><uid>U_PersonelNo</uid><value>{0}</value></cell><cell><uid>U_PersonelAdi</uid><value>{1}</value></cell><cell><uid>U_Bolum1</uid><value>{2}</value></cell><cell><uid>U_Bolum2</uid><value>{3}</value></cell><cell><uid>U_Bolum3</uid><value>{4}</value></cell><cell><uid>U_Gun1</uid><value>{5}</value></cell><cell><uid>U_Gun2</uid><value>{6}</value></cell><cell><uid>U_Gun3</uid><value>{7}</value></cell><cell><uid>U_Gun4</uid><value>{8}</value></cell><cell><uid>U_Gun5</uid><value>{9}</value></cell><cell><uid>U_Gun6</uid><value>{10}</value></cell><cell><uid>U_Gun7</uid><value>{11}</value></cell><cell><uid>U_Gun8</uid><value>{12}</value></cell><cell><uid>U_Gun9</uid><value>{13}</value></cell><cell><uid>U_Gun10</uid><value>{14}</value></cell><cell><uid>U_Gun11</uid><value>{15}</value></cell><cell><uid>U_Gun12</uid><value>{16}</value></cell><cell><uid>U_Gun13</uid><value>{17}</value></cell><cell><uid>U_Gun14</uid><value>{18}</value></cell><cell><uid>U_Gun15</uid><value>{19}</value></cell>><cell><uid>U_Gun16</uid><value>{20}</value></cell>><cell><uid>U_Gun17</uid><value>{21}</value></cell>><cell><uid>U_Gun18</uid><value>{22}</value></cell>><cell><uid>U_Gun19</uid><value>{23}</value></cell>><cell><uid>U_Gun20</uid><value>{24}</value></cell>><cell><uid>U_Gun21</uid><value>{25}</value></cell>><cell><uid>U_Gun22</uid><value>{26}</value></cell>><cell><uid>U_Gun23</uid><value>{27}</value></cell>><cell><uid>U_Gun24</uid><value>{28}</value></cell>><cell><uid>U_Gun25</uid><value>{29}</value></cell>><cell><uid>U_Gun26</uid><value>{30}</value></cell>><cell><uid>U_Gun27</uid><value>{31}</value></cell>><cell><uid>U_Gun28</uid><value>{32}</value></cell>><cell><uid>U_Gun29</uid><value>{33}</value></cell>><cell><uid>U_Gun30</uid><value>{34}</value></cell>><cell><uid>U_Gun31</uid><value>{35}</value></cell></cells></row>";

        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.GunlukPersonelPlan2XML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.GunlukPersonelPlan2XML));
            Functions.CreateUserOrSystemFormComponent<GunlukPersonelPlanlama2>(AIFConn.PersPlanG);
            InitForms();
        }

        public void InitForms()
        {
            try
            {
                frmGunlukPersonelPlanlama2.Freeze(true);
                frmGunlukPersonelPlanlama2.EnableMenu("1283", false);
                frmGunlukPersonelPlanlama2.EnableMenu("1284", false);
                frmGunlukPersonelPlanlama2.EnableMenu("1286", false);
                oMatrix.AutoResizeColumns();
                CalisanListele();
                oComboYil.Select(DateTime.Now.Year.ToString());
                //oEditTariheKopyala.Item.AffectsFormMode = false;

                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
                string sql = "Select \"U_BolumKodu\",\"U_BolumAdi\" from \"@AIF_BOLUMLER\"";

                ConstVariables.oRecordset.DoQuery(sql);
                SAPbouiCOM.Column oColumn;
                SAPbouiCOM.Column oColumn1;
                SAPbouiCOM.Column oColumn2;

                oColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item("Col_2");
                oColumn1 = (SAPbouiCOM.Column)oMatrix.Columns.Item("Col_3");
                oColumn2 = (SAPbouiCOM.Column)oMatrix.Columns.Item("Col_4");
                while (!ConstVariables.oRecordset.EoF)
                {
                    oColumn.ValidValues.Add(ConstVariables.oRecordset.Fields.Item(0).Value.ToString(), ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                    oColumn1.ValidValues.Add(ConstVariables.oRecordset.Fields.Item(0).Value.ToString(), ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                    oColumn2.ValidValues.Add(ConstVariables.oRecordset.Fields.Item(0).Value.ToString(), ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                    ConstVariables.oRecordset.MoveNext();
                }


                frmGunlukPersonelPlanlama2.EnableMenu("5890", true);
            }
            catch (Exception ex)
            {
            }
            finally
            {
                frmGunlukPersonelPlanlama2.Freeze(false);
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

            public string BirinciGun { get; set; }

            public string IkinciGun { get; set; }

            public string UcuncuGun { get; set; }

            public string DorduncuGun { get; set; }

            public string BesinciGun { get; set; }

            public string AltinciGun { get; set; }

            public string YedinciGun { get; set; }

            public string SekizinciGun { get; set; }

            public string DokuzuncuGun { get; set; }

            public string OnuncuGun { get; set; }

            public string OnbirinciGun { get; set; }

            public string OnikinciGun { get; set; }

            public string OnUcuncuGun { get; set; }

            public string OnDorduncuGun { get; set; }

            public string OnBesinciGun { get; set; }

            public string OnAltinciGun { get; set; }

            public string OnYedinciGun { get; set; }

            public string OnSekizinciGun { get; set; }

            public string OnDokuzuncuGun { get; set; }

            public string YirminciGun { get; set; }

            public string YirmiBirinciGun { get; set; }

            public string YirmiIkinciGun { get; set; }

            public string YirmiUcuncuGun { get; set; }

            public string YirmiDorduncuGun { get; set; }

            public string YirmiBesinciGun { get; set; }

            public string YirmiAltinciGun { get; set; }

            public string YirmiYedinciGun { get; set; }

            public string YirmiSekizinciGun { get; set; }

            public string YirmiDokuzuncuGun { get; set; }

            public string OtuzuncuGun { get; set; }

            public string OtuzBirinciGun { get; set; }
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
                                     Bolum1 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value,
                                     Bolum2 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value,
                                     Bolum3 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,
                                     Bolum1Adi = "",
                                     Bolum2Adi = "",
                                     Bolum3Adi = "",
                                     BirinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                                     IkinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value,
                                     UcuncuGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value,
                                     DorduncuGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value,
                                     BesinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_9" select new XElement(y.Element("Value"))).First().Value,
                                     AltinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_10" select new XElement(y.Element("Value"))).First().Value,
                                     YedinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_11" select new XElement(y.Element("Value"))).First().Value,
                                     SekizinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_12" select new XElement(y.Element("Value"))).First().Value,
                                     DokuzuncuGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_13" select new XElement(y.Element("Value"))).First().Value,
                                     OnuncuGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_14" select new XElement(y.Element("Value"))).First().Value,
                                     OnbirinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_15" select new XElement(y.Element("Value"))).First().Value,
                                     OnikinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_16" select new XElement(y.Element("Value"))).First().Value,
                                     OnUcuncuGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_17" select new XElement(y.Element("Value"))).First().Value,
                                     OnDorduncuGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_18" select new XElement(y.Element("Value"))).First().Value,
                                     OnBesinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_19" select new XElement(y.Element("Value"))).First().Value,
                                     OnAltinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_20" select new XElement(y.Element("Value"))).First().Value,
                                     OnYedinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_21" select new XElement(y.Element("Value"))).First().Value,
                                     OnSekizinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_22" select new XElement(y.Element("Value"))).First().Value,
                                     OnDokuzuncuGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_23" select new XElement(y.Element("Value"))).First().Value,
                                     YirminciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_24" select new XElement(y.Element("Value"))).First().Value,
                                     YirmiBirinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_25" select new XElement(y.Element("Value"))).First().Value,
                                     YirmiIkinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_26" select new XElement(y.Element("Value"))).First().Value,
                                     YirmiUcuncuGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_27" select new XElement(y.Element("Value"))).First().Value,
                                     YirmiDorduncuGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_28" select new XElement(y.Element("Value"))).First().Value,
                                     YirmiBesinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_29" select new XElement(y.Element("Value"))).First().Value,
                                     YirmiAltinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_30" select new XElement(y.Element("Value"))).First().Value,
                                     YirmiYedinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_31" select new XElement(y.Element("Value"))).First().Value,
                                     YirmiSekizinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_32" select new XElement(y.Element("Value"))).First().Value,
                                     YirmiDokuzuncuGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_33" select new XElement(y.Element("Value"))).First().Value,
                                     OtuzuncuGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_34" select new XElement(y.Element("Value"))).First().Value,
                                     OtuzBirinciGun = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_35" select new XElement(y.Element("Value"))).First().Value,

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

        private void CalisanListele()
        {
            try
            {
                frmGunlukPersonelPlanlama2.Freeze(true);
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
                    frmGunlukPersonelPlanlama2.DataSources.DBDataSources.Item("@AIF_GUNLUKPERSPLAN1").LoadFromXML(string.Format(header, data));

                    oMatrix.AutoResizeColumns();

                    MatrisXMLiniListeyeYukle();
                }
                //else
                //{
                //    ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //    ConstVariables.oRecordset.DoQuery("Select * from \"@AIF_GUNPERSPLAN\" where \"U_Tarih\"= '" + operasyonplaniTarihi + "'");
                //    string docEntry = ConstVariables.oRecordset.Fields.Item("DocEntry").Value.ToString();

                //    if (docEntry != "0")
                //    {
                //        frmGunlukPersonelPlanlama2.Mode = BoFormMode.fm_FIND_MODE;
                //        oEditDocEntry.Value = docEntry;
                //        btnAddOrUpdate.Item.Click();
                //    }
                //}
            }
            catch (Exception)
            {
            }
            finally
            {
                frmGunlukPersonelPlanlama2.Freeze(false);
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
                        //renklendir();
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
            //string val = tarih;
            //ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //ConstVariables.oRecordset.DoQuery("Select * from \"@AIF_GUNPERSPLAN\" where \"U_Tarih\"= '" + val + "'");

            //if (ConstVariables.oRecordset.RecordCount > 0 && !question)
            //{
            //    question = true;
            //    int retval = Handler.SAPApplication.MessageBox("İlgili tarih için daha önce giriş yapılmıştır.", 1, "İlgili Kaydı Getir  ", "İlgili Kaydı Sil", "İptal");

            //    if (retval == 1)
            //    {
            //        string docEntry = ConstVariables.oRecordset.Fields.Item("DocEntry").Value.ToString();
            //        frmGunlukPersonelPlanlama.Mode = BoFormMode.fm_FIND_MODE;
            //        oEditDocEntry.Value = docEntry;
            //        btnAddOrUpdate.Item.Click();
            //    }
            //    else if (retval == 2)
            //    {
            //        string docEntry = ConstVariables.oRecordset.Fields.Item("DocEntry").Value.ToString();
            //        SAPbobsCOM.GeneralService oGeneralService = null;

            //        SAPbobsCOM.GeneralData oGeneralData = null;

            //        SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            //        SAPbobsCOM.CompanyService oCompService = null;

            //        oCompService = ConstVariables.oCompanyObject.GetCompanyService();

            //        oGeneralService = oCompService.GetGeneralService("AIF_GUNPERSPLAN");

            //        oGeneralParams = ((SAPbobsCOM.GeneralDataParams)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)));

            //        oGeneralParams.SetProperty("DocEntry", docEntry);

            //        oGeneralData = oGeneralService.GetByParams(oGeneralParams);

            //        oGeneralService.Delete(oGeneralParams);
            //    }
            //}
            //else
            //    question = false;
        }

        private bool question = false;

        private void TarihleriOlustur(string yil, string ay)
        {
            try
            {
                frmGunlukPersonelPlanlama2.Freeze(true);
                DateTime dt = new DateTime(Convert.ToInt32(yil), Convert.ToInt32(ay), 01);

                int colNum = 5;
                oMatrix.Columns.Item("Col_33").Width = 50;
                oMatrix.Columns.Item("Col_34").Width = 50;
                oMatrix.Columns.Item("Col_35").Width = 50;
                for (int i = 1; i <= 31; i++)
                {
                    if (Convert.ToInt32(dt.Month.ToString().PadLeft(2, '0')) > Convert.ToInt32(ay))
                    {
                        oMatrix.Columns.Item("Col_" + colNum).Width = 0;
                        colNum++;
                        continue;
                    }
                    oMatrix.Columns.Item("Col_" + colNum).TitleObject.Caption = dt.Day.ToString().PadLeft(2, '0') + "." + dt.Month.ToString().PadLeft(2, '0') + "." + dt.Year.ToString();
                    colNum++;
                    dt = dt.AddDays(1);
                }

            }
            catch (Exception)
            {
            }
            finally
            {
                frmGunlukPersonelPlanlama2.Freeze(false);
            }
        }

        private void kololariSifirla()
        {
            try
            {
                frmGunlukPersonelPlanlama2.Freeze(true);
                int colNum = 5;
                oMatrix.Columns.Item("Col_33").Width = 50;
                oMatrix.Columns.Item("Col_34").Width = 50;
                oMatrix.Columns.Item("Col_35").Width = 50;
                for (int i = 1; i <= 31; i++)
                {
                    oMatrix.Columns.Item("Col_" + colNum).TitleObject.Caption = "Col" + i;
                    colNum++;
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                frmGunlukPersonelPlanlama2.Freeze(false);
            }
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
                            foreach (var item in gunlukPersonelPlanlama3s)
                            {
                                sql = "Select \"DocEntry\" from \"@AIF_GUNLUKPLAN\" where \"U_Tarih\"='" + item.Tarih + "'";

                                ConstVariables.oRecordset.DoQuery(sql);

                                if (ConstVariables.oRecordset.RecordCount > 0)
                                {
                                    sql2 = "UPDATE \"@AIF_GUNLUKPLAN1\" SET \"U_Durum\" = '" + item.Durum + "' where \"U_PersonelNo\" = '" + item.KullaniciId + "' and \"DocEntry\" = '" + ConstVariables.oRecordset.Fields.Item(0).Value + "'";
                                    ConstVariables.oRecordset.DoQuery(sql2);
                                }
                            }

                            if (oMatrix.RowCount == 0)
                            {
                                kololariSifirla();
                                CalisanListele();
                                renklendir(true);
                            }
                            else
                            {
                                renklendir();
                            }

                            gunlukPersonelPlanlama3s = new List<GunlukPersonelPlanlama3>();
                        }
                        catch (Exception)
                        {
                        }
                    }
                    break;

                case BoEventTypes.et_KEY_DOWN:
                    if (!pVal.BeforeAction && pVal.ItemUID == "Item_6")
                    {

                    }
                    break;

                case BoEventTypes.et_GOT_FOCUS:
                    break;

                case BoEventTypes.et_LOST_FOCUS:
                    if (pVal.ItemUID == "Item_8" && !pVal.BeforeAction)
                    {
                        try
                        {
                            //if (frmGunlukPersonelPlanlama != null)
                            //{
                            //    if (frmGunlukPersonelPlanlama.Mode != BoFormMode.fm_FIND_MODE && frmGunlukPersonelPlanlama.Mode == BoFormMode.fm_ADD_MODE)
                            //    {
                            //        TarihKontrol(oEditTarih.Value.ToString());
                            //    }
                            //}
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_1" && !pVal.BeforeAction)
                    {
                        try
                        {
                            //if (frmGunlukPersonelPlanlama.Mode != BoFormMode.fm_FIND_MODE)
                            //{
                            //    TarihKontrol(oEditTariheKopyala.Value.ToString());
                            //}
                        }
                        catch (Exception)
                        {
                        }
                    }
                    break;

                case BoEventTypes.et_COMBO_SELECT:
                    if (!pVal.BeforeAction && pVal.ItemUID == "Item_1")
                    {
                        DateTime baslangic = DateTime.Now;
                        ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
                        ConstVariables.oRecordset.DoQuery("Select * from \"@AIF_GUNLUKPERSPLAN\" where \"U_Aylar\" = '" + oComboAy.Value.ToString().Trim() + "' and \"U_Yil\" = '" + oComboYil.Value.ToString().Trim() + "' ");


                        if (ConstVariables.oRecordset.RecordCount > 0)
                        {
                            frmGunlukPersonelPlanlama2.Mode = BoFormMode.fm_FIND_MODE;
                            oEditDocEntry.Value = ConstVariables.oRecordset.Fields.Item("DocEntry").Value.ToString();
                            btnAddOrUpdate.Item.Click();

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
                                    frmGunlukPersonelPlanlama2.DataSources.DBDataSources.Item(1).Clear();
                                    oMatrix.AddRow();
                                    //oMatrix.ClearRowData(oMatrix.RowCount);
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("#").Cells.Item(oMatrix.RowCount).Specific).Value = item.AdSoyad;
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(oMatrix.RowCount).Specific).Value = item.AdSoyad;
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific).Value = item.Id;
                                }

                                if (FinalXml.Count > 0)
                                {
                                    renklendir();
                                }

                                oMatrix.AutoResizeColumns();

                                MatrisXMLiniListeyeYukle();

                            }

                        }
                        else
                        {
                            string yil = oComboYil.Value.Trim();
                            string ay = oComboAy.Value.ToString();
                            frmGunlukPersonelPlanlama2.Mode = BoFormMode.fm_ADD_MODE;
                            CalisanListele();
                            oComboYil.Select(yil.ToString());
                            frmGunlukPersonelPlanlama2.DataSources.DBDataSources.Item(0).SetValue("U_Aylar", 0, ay);
                            renklendir();

                            MatrisXMLiniListeyeYukle();

                        }

                        TarihleriOlustur(oComboYil.Value.ToString(), oComboAy.Value.ToString());



                        oMatrix.AutoResizeColumns();



                        Handler.SAPApplication.StatusBar.SetText("Listeleme tamamlandı.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

                        //DateTime bitis = DateTime.Now;

                        //TimeSpan t = bitis - baslangic;

                        //Handler.SAPApplication.MessageBox("Dakika:" + t.TotalMinutes + "Saniye:" + t.TotalSeconds);

                    }
                    else if (!pVal.BeforeAction && (pVal.ColUID == "Col_2" || pVal.ColUID == "Col_3" || pVal.ColUID == "Col_4") || pVal.ColUID == "Col_5" || pVal.ColUID == "Col_6" || pVal.ColUID == "Col_7" || pVal.ColUID == "Col_8" || pVal.ColUID == "Col_9" || pVal.ColUID == "Col_10" || pVal.ColUID == "Col_11" || pVal.ColUID == "Col_12" || pVal.ColUID == "Col_13" || pVal.ColUID == "Col_14" || pVal.ColUID == "Col_15" || pVal.ColUID == "Col_16" || pVal.ColUID == "Col_17" || pVal.ColUID == "Col_18" || pVal.ColUID == "Col_19" || pVal.ColUID == "Col_20" || pVal.ColUID == "Col_21" || pVal.ColUID == "Col_22" || pVal.ColUID == "Col_23" || pVal.ColUID == "Col_24" || pVal.ColUID == "Col_25" || pVal.ColUID == "Col_26" || pVal.ColUID == "Col_27" || pVal.ColUID == "Col_28" || pVal.ColUID == "Col_29" || pVal.ColUID == "Col_30" || pVal.ColUID == "Col_31" || pVal.ColUID == "Col_32" || pVal.ColUID == "Col_33" || pVal.ColUID == "Col_34" || pVal.ColUID == "Col_35")
                    {
                        var id = Convert.ToInt32(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value);
                        string value = "";
                        value = StringReplace(((SAPbouiCOM.ComboBox)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).Value.ToString());

                        if (pVal.ColUID != "Col_2" && pVal.ColUID != "Col_3" && pVal.ColUID != "Col_4")
                        {
                            var date = oMatrix.Columns.Item(pVal.ColUID).TitleObject.Caption;

                            var datefinal = date.ToString().Substring(6, 4) + date.Substring(3, 2) + date.Substring(0, 2);


                            gunlukPersonelPlanlama3s.RemoveAll(x => x.KullaniciId == id);

                            gunlukPersonelPlanlama3s.Add(new GunlukPersonelPlanlama3 { KullaniciId = id, Durum = value, Tarih = datefinal });
                        }

                        #region Güncellemeler
                        if (pVal.ColUID == "Col_2")
                        {
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.Bolum1 = value);
                            ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
                            string sql = "";
                            sql = "Select \"U_BolumKodu\",\"U_BolumAdi\" from \"@AIF_BOLUMLER\" where  \"U_BolumKodu\"='" + value + "'";

                            ConstVariables.oRecordset.DoQuery(sql);

                            MatrixOrjinalData.Where(x => x.Bolum1 == value).ToList().ForEach(y => y.Bolum1Adi = ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                        }
                        else if (pVal.ColUID == "Col_3")
                        {
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.Bolum2 = value);

                            ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
                            string sql = "";
                            sql = "Select \"U_BolumKodu\",\"U_BolumAdi\" from \"@AIF_BOLUMLER\" where  \"U_BolumKodu\"='" + value + "'";

                            ConstVariables.oRecordset.DoQuery(sql);
                            MatrixOrjinalData.Where(x => x.Bolum2 == value).ToList().ForEach(y => y.Bolum2Adi = ConstVariables.oRecordset.Fields.Item(1).Value.ToString());

                        }
                        else if (pVal.ColUID == "Col_4")
                        {
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.Bolum3 = value);

                            ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
                            string sql = "";
                            sql = "Select \"U_BolumKodu\",\"U_BolumAdi\" from \"@AIF_BOLUMLER\" where  \"U_BolumKodu\"='" + value + "'";

                            ConstVariables.oRecordset.DoQuery(sql);
                            MatrixOrjinalData.Where(x => x.Bolum3 == value).ToList().ForEach(y => y.Bolum3Adi = ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                        }
                        else if (pVal.ColUID == "Col_5")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.BirinciGun = value);
                        else if (pVal.ColUID == "Col_6")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.IkinciGun = value);
                        else if (pVal.ColUID == "Col_7")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.UcuncuGun = value);
                        else if (pVal.ColUID == "Col_8")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.DorduncuGun = value);
                        else if (pVal.ColUID == "Col_9")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.BesinciGun = value);
                        else if (pVal.ColUID == "Col_10")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.AltinciGun = value);
                        else if (pVal.ColUID == "Col_11")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.YedinciGun = value);
                        else if (pVal.ColUID == "Col_12")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.SekizinciGun = value);
                        else if (pVal.ColUID == "Col_13")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.DokuzuncuGun = value);
                        else if (pVal.ColUID == "Col_14")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.OnuncuGun = value);
                        else if (pVal.ColUID == "Col_15")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.OnbirinciGun = value);
                        else if (pVal.ColUID == "Col_16")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.OnikinciGun = value);
                        else if (pVal.ColUID == "Col_17")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.OnUcuncuGun = value);
                        else if (pVal.ColUID == "Col_18")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.OnDorduncuGun = value);
                        else if (pVal.ColUID == "Col_19")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.OnBesinciGun = value);
                        else if (pVal.ColUID == "Col_20")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.OnAltinciGun = value);
                        else if (pVal.ColUID == "Col_21")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.OnYedinciGun = value);
                        else if (pVal.ColUID == "Col_22")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.OnSekizinciGun = value);
                        else if (pVal.ColUID == "Col_23")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.OnDokuzuncuGun = value);
                        else if (pVal.ColUID == "Col_24")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.YirminciGun = value);
                        else if (pVal.ColUID == "Col_25")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.YirmiBirinciGun = value);
                        else if (pVal.ColUID == "Col_26")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.YirmiIkinciGun = value);
                        else if (pVal.ColUID == "Col_27")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.YirmiUcuncuGun = value);
                        else if (pVal.ColUID == "Col_28")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.YirmiDorduncuGun = value);
                        else if (pVal.ColUID == "Col_29")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.YirmiBesinciGun = value);
                        else if (pVal.ColUID == "Col_30")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.YirmiAltinciGun = value);
                        else if (pVal.ColUID == "Col_31")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.YirmiYedinciGun = value);
                        else if (pVal.ColUID == "Col_32")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.YirmiSekizinciGun = value);
                        else if (pVal.ColUID == "Col_33")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.YirmiDokuzuncuGun = value);
                        else if (pVal.ColUID == "Col_34")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.OtuzuncuGun = value);
                        else if (pVal.ColUID == "Col_35")
                            MatrixOrjinalData.Where(x => x.Id == id).ToList().ForEach(y => y.OtuzBirinciGun = value);

                        #endregion
                    }
                    break;

                case BoEventTypes.et_CLICK:
                    if (pVal.ItemUID == "Item_0" && !pVal.BeforeAction)
                    {
                        try
                        {
                            //frmGunlukPersonelPlanlama.Freeze(true);
                            //string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                            //var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                            //            select new
                            //            {
                            //                AdSoyad = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "#" select new XElement(y.Element("Value"))).First().Value,
                            //                Id = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_18" select new XElement(y.Element("Value"))).First().Value,
                            //            }).ToList();

                            //ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            //ConstVariables.oRecordset.DoQuery("Select (\"firstName\" + ' ' + \"lastName\") as \"AdSoyad\",\"empID\"  from \"OHEM\" where \"Active\" = 'Y'");

                            //string xmll = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                            //XDocument xDoc = XDocument.Parse(xmll);
                            //XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                            //var rowsOhem = (from t in xDoc.Descendants(ns + "Row")
                            //                select new
                            //                {
                            //                    AdSoyad = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "AdSoyad" select new XElement(y.Element(ns + "Value"))).First().Value,
                            //                    Id = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "empID" select new XElement(y.Element(ns + "Value"))).First().Value,
                            //                }).ToList();

                            //var DeleteList = rows.Except(rowsOhem).ToList();
                            //var addList = rowsOhem.Except(rows).ToList();

                            //var rows2 = (from x in XDocument.Parse(xml).Descendants("Row")
                            //             select new
                            //             {
                            //                 PersonelAdi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "#" select new XElement(y.Element("Value"))).First().Value,
                            //                 PersonelNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_18" select new XElement(y.Element("Value"))).First().Value,
                            //                 Ayran = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                            //                 Yogurt = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                            //                 Teleme = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value,
                            //                 Kasar = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value,
                            //                 Kaymak = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value,
                            //                 Lor = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                            //                 Tereyag = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value,
                            //                 Dondurma = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value,
                            //                 Kolileme = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value,
                            //                 SutAlim = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_9" select new XElement(y.Element("Value"))).First().Value,
                            //                 BulasikHane = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_10" select new XElement(y.Element("Value"))).First().Value,
                            //                 SevkiyatDepo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_11" select new XElement(y.Element("Value"))).First().Value,
                            //                 AmbalajDepo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_12" select new XElement(y.Element("Value"))).First().Value,
                            //                 Tarihleme = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_13" select new XElement(y.Element("Value"))).First().Value,
                            //                 GenelTemizlik = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_14" select new XElement(y.Element("Value"))).First().Value,
                            //                 Mutfak = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_15" select new XElement(y.Element("Value"))).First().Value,
                            //                 KazanDairesi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_16" select new XElement(y.Element("Value"))).First().Value,
                            //                 BakimOnarim = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_17" select new XElement(y.Element("Value"))).First().Value,
                            //             }).ToList();

                            //foreach (var item in DeleteList)
                            //{
                            //    rows2.RemoveAll(x => x.PersonelNo == item.Id);
                            //}

                            //foreach (var item in addList)
                            //{
                            //    rows2.Insert(rows2.Count, new { PersonelAdi = item.AdSoyad, PersonelNo = item.Id, Ayran = "N", Yogurt = "N", Teleme = "N", Kasar = "N", Kaymak = "N", Lor = "N", Tereyag = "N", Dondurma = "N", Kolileme = "N", SutAlim = "N", BulasikHane = "N", SevkiyatDepo = "N", AmbalajDepo = "N", Tarihleme = "N", GenelTemizlik = "N", Mutfak = "N", KazanDairesi = "N", BakimOnarim = "N" });
                            //}

                            //frmGunlukPersonelPlanlama.Mode = BoFormMode.fm_ADD_MODE;
                            //oEditTarih.Value = oEditTariheKopyala.Value;

                            //string data = string.Join("", rows2.Select(s => string.Format(row, s.PersonelAdi, s.PersonelNo, s.Ayran, s.Yogurt, s.Teleme, s.Kasar, s.Kaymak, s.Lor, s.Tereyag, s.Dondurma, s.Kolileme, s.SutAlim, s.BulasikHane, s.SevkiyatDepo, s.AmbalajDepo, s.Tarihleme, s.GenelTemizlik, s.Mutfak, s.KazanDairesi, s.BakimOnarim)));

                            //frmGunlukPersonelPlanlama.DataSources.DBDataSources.Item("@AIF_GUNPERSPLAN1").LoadFromXML(string.Format(header, data));

                            //oMatrix.AutoResizeColumns();

                            //renklendir();

                            //Handler.SAPApplication.MessageBox("Kopyalama işlemi tamamlandı. Ekleme işlemi ile devam edebilirsiniz.");

                            //oEditTariheKopyala.Value = "";

                            //oMatrix.Clear();
                            //oMatrix.AddRow(160);
                        }
                        catch (Exception ex)
                        {
                        }
                        finally
                        {
                            //frmGunlukPersonelPlanlama.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    {
                        try
                        {
                            oEditArama.Value = "";
                            var value = StringReplace(oEditArama.Value.ToUpper());

                            frmGunlukPersonelPlanlama2.Freeze(true);

                            var rows = MatrixOrjinalData.ToList();

                            string data = string.Join("", rows.Select(s => string.Format(row, s.Id, s.AdSoyad, s.Bolum1, s.Bolum2, s.Bolum3, s.BirinciGun, s.IkinciGun, s.UcuncuGun, s.DorduncuGun, s.BesinciGun, s.AltinciGun, s.YedinciGun, s.SekizinciGun, s.DokuzuncuGun, s.OnuncuGun, s.OnbirinciGun, s.OnikinciGun, s.OnUcuncuGun, s.OnDorduncuGun, s.OnBesinciGun, s.OnAltinciGun, s.OnYedinciGun, s.OnSekizinciGun, s.OnDokuzuncuGun, s.YirminciGun, s.YirmiBirinciGun, s.YirmiIkinciGun, s.YirmiUcuncuGun, s.YirmiDorduncuGun, s.YirmiBesinciGun, s.YirmiAltinciGun, s.YirmiYedinciGun, s.YirmiSekizinciGun, s.YirmiDokuzuncuGun, s.OtuzuncuGun, s.OtuzBirinciGun)));

                            string s1 = string.Format(header, data);
                            frmGunlukPersonelPlanlama2.DataSources.DBDataSources.Item("@AIF_GUNLUKPERSPLAN1").LoadFromXML(string.Format(header, data));

                            oMatrix.AutoResizeColumns();



                            renklendir();

                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmGunlukPersonelPlanlama2.Freeze(false);
                        }
                    }
                    //btnTemizle.Item.Click();                        //btnTemizle.Item.Click(); 
                    else if (pVal.ItemUID == "Item_11" && !pVal.BeforeAction)
                    {
                        try
                        {
                            oEditArama.Value = "";
                            var value = StringReplace(oEditArama.Value.ToUpper());

                            frmGunlukPersonelPlanlama2.Freeze(true);

                            var rows = MatrixOrjinalData.ToList();

                            string data = string.Join("", rows.Select(s => string.Format(row, s.Id, s.AdSoyad, s.Bolum1, s.Bolum2, s.Bolum3, s.BirinciGun, s.IkinciGun, s.UcuncuGun, s.DorduncuGun, s.BesinciGun, s.AltinciGun, s.YedinciGun, s.SekizinciGun, s.DokuzuncuGun, s.OnuncuGun, s.OnbirinciGun, s.OnikinciGun, s.OnUcuncuGun, s.OnDorduncuGun, s.OnBesinciGun, s.OnAltinciGun, s.OnYedinciGun, s.OnSekizinciGun, s.OnDokuzuncuGun, s.YirminciGun, s.YirmiBirinciGun, s.YirmiIkinciGun, s.YirmiUcuncuGun, s.YirmiDorduncuGun, s.YirmiBesinciGun, s.YirmiAltinciGun, s.YirmiYedinciGun, s.YirmiSekizinciGun, s.YirmiDokuzuncuGun, s.OtuzuncuGun, s.OtuzBirinciGun)));

                            string s1 = string.Format(header, data);
                            frmGunlukPersonelPlanlama2.DataSources.DBDataSources.Item("@AIF_GUNLUKPERSPLAN1").LoadFromXML(string.Format(header, data));

                            oMatrix.AutoResizeColumns();



                            renklendir();



                            Handler.SAPApplication.StatusBar.SetText("Listeleme tamamlandı.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmGunlukPersonelPlanlama2.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "Item_10" && !pVal.BeforeAction)
                    {
                        try
                        {
                            var value = StringReplace(oEditArama.Value.ToUpper());

                            frmGunlukPersonelPlanlama2.Freeze(true);

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



                            string data = string.Join("", rows.Select(s => string.Format(row, s.Id, s.AdSoyad, s.Bolum1, s.Bolum2, s.Bolum3, s.BirinciGun, s.IkinciGun, s.UcuncuGun, s.DorduncuGun, s.BesinciGun, s.AltinciGun, s.YedinciGun, s.SekizinciGun, s.DokuzuncuGun, s.OnuncuGun, s.OnbirinciGun, s.OnikinciGun, s.OnUcuncuGun, s.OnDorduncuGun, s.OnBesinciGun, s.OnAltinciGun, s.OnYedinciGun, s.OnSekizinciGun, s.OnDokuzuncuGun, s.YirminciGun, s.YirmiBirinciGun, s.YirmiIkinciGun, s.YirmiUcuncuGun, s.YirmiDorduncuGun, s.YirmiBesinciGun, s.YirmiAltinciGun, s.YirmiYedinciGun, s.YirmiSekizinciGun, s.YirmiDokuzuncuGun, s.OtuzuncuGun, s.OtuzBirinciGun)));

                            string s1 = string.Format(header, data);
                            frmGunlukPersonelPlanlama2.DataSources.DBDataSources.Item("@AIF_GUNLUKPERSPLAN1").LoadFromXML(string.Format(header, data));

                            oMatrix.AutoResizeColumns();



                            renklendir();

                            Handler.SAPApplication.StatusBar.SetText("Filtreleme sonucu kayıtlar listelendi.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                            //var value = oEditArama.Value.ToUpper();
                            //frmGunlukPersonelPlanlama2.Freeze(true);
                            //string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                            //var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                            //            select new
                            //            {
                            //                AdSoyad = KarakterDegistir((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "#" select new XElement(y.Element("Value"))).First().Value.ToUpper()),
                            //                Id = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                            //                Satir = Convert.ToString(x.ElementsBeforeSelf().Count() + 1)
                            //            }).ToList();

                            //if (value != "")
                            //{
                            //    value = KarakterDegistir(value).ToUpper();


                            //    string sira = rows.Where(x => x.AdSoyad.Contains(value)).Select(y => y.Satir).FirstOrDefault();
                            //    if (sira != null)
                            //    {
                            //        oMatrix.SelectRow(Convert.ToInt32(sira), true, false);
                            //    }
                            //    else
                            //    {
                            //        try
                            //        {
                            //            sira = rows.Where(x => x.Id == value).Select(y => y.Satir).FirstOrDefault().ToString();

                            //            if (sira != null)
                            //            {
                            //                oMatrix.SelectRow(Convert.ToInt32(sira), true, false);
                            //            }

                            //        }
                            //        catch (Exception)
                            //        {
                            //            oMatrix.ClearSelections();
                            //        }
                            //    }
                            //}
                            //else
                            //{
                            //    oMatrix.ClearSelections();
                            //}

                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmGunlukPersonelPlanlama2.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "Item_9" && !pVal.BeforeAction)
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

                        string data = string.Join("", rows.Select(s => string.Format(row, s.Id, s.AdSoyad, s.Bolum1, s.Bolum2, s.Bolum3, s.BirinciGun, s.IkinciGun, s.UcuncuGun, s.DorduncuGun, s.BesinciGun, s.AltinciGun, s.YedinciGun, s.SekizinciGun, s.DokuzuncuGun, s.OnuncuGun, s.OnbirinciGun, s.OnikinciGun, s.OnUcuncuGun, s.OnDorduncuGun, s.OnBesinciGun, s.OnAltinciGun, s.OnYedinciGun, s.OnSekizinciGun, s.OnDokuzuncuGun, s.YirminciGun, s.YirmiBirinciGun, s.YirmiIkinciGun, s.YirmiUcuncuGun, s.YirmiDorduncuGun, s.YirmiBesinciGun, s.YirmiAltinciGun, s.YirmiYedinciGun, s.YirmiSekizinciGun, s.YirmiDokuzuncuGun, s.OtuzuncuGun, s.OtuzBirinciGun)));

                        string s1 = string.Format(header, data);
                        frmGunlukPersonelPlanlama2.DataSources.DBDataSources.Item("@AIF_GUNLUKPERSPLAN1").LoadFromXML(string.Format(header, data));

                        oMatrix.AutoResizeColumns();



                        renklendir();

                        Handler.SAPApplication.StatusBar.SetText("Listeleme tamamlandı.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);


                        if (frmGunlukPersonelPlanlama2.Mode != BoFormMode.fm_ADD_MODE)
                        {
                            frmGunlukPersonelPlanlama2.Mode = BoFormMode.fm_UPDATE_MODE;
                        }
                        //xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                        //var rowsSira = (from x in XDocument.Parse(xml).Descendants("Row")
                        //                select new
                        //                {
                        //                    AdSoyad = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "#" select new XElement(y.Element("Value"))).First().Value,
                        //                    Id = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                        //                    Sira = x.ElementsBeforeSelf().Count() + 1
                        //                }).ToList();


                        //foreach (var item in finalRows)
                        //{
                        //    var sira = rowsSira.Where(x => x.Id == item.Id).Count() == 0 ? -1 : rowsSira.Where(x => x.Id == item.Id).Select(y => y.Sira).FirstOrDefault();

                        //    if (sira != -1)
                        //    {
                        //        oMatrix.DeleteRow(sira);

                        //        if (frmGunlukPersonelPlanlama2.Mode != BoFormMode.fm_ADD_MODE)
                        //        {
                        //            frmGunlukPersonelPlanlama2.Mode = BoFormMode.fm_UPDATE_MODE;
                        //        }
                        //    }
                        //}



                    }
                    else if (pVal.ItemUID == "Item_0" && pVal.BeforeAction)
                    {
                        //if (frmGunlukPersonelPlanlama.Mode == BoFormMode.fm_ADD_MODE || frmGunlukPersonelPlanlama.Mode == BoFormMode.fm_UPDATE_MODE)
                        //{
                        //    Handler.SAPApplication.MessageBox("Ekleme veya güncelleme modunda iken kopyalama yapılamaz.");
                        //    BubbleEvent = false;
                        //    return false;
                        //}
                        //else
                        //    oEditDocEntry.Item.Click();
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

        private string KarakterDegistir(string text)
        {
            string val = text;

            val = val.Replace("i", "I");
            val = val.Replace("İ", "I");
            val = val.Replace("ç", "C");
            val = val.Replace("Ç", "C");
            val = val.Replace("ş", "S");
            val = val.Replace("Ş", "S");
            val = val.Replace("Ü", "U");
            val = val.Replace("ü", "U");
            val = val.Replace("Ö", "O");
            val = val.Replace("ö", "O");
            val = val.Replace("Ğ", "G");
            val = val.Replace("ğ", "G");

            return val;
        }
        private void renklendir(bool addbtn = false)
        {
            try
            {
                frmGunlukPersonelPlanlama2.Freeze(true);

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    oMatrix.CommonSetting.SetRowBackColor(i, -1);
                }
                //if (!addbtn)
                //{
                string xmlColor = oMatrix.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);
                XDocument xDoc = XDocument.Parse(xmlColor);
                var table = (from x in xDoc.Elements("Matrix").Elements("Rows").Elements("Row")
                             select new
                             {
                                 Gun1 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                                 Gun1Index = 6,
                                 Gun2 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value,
                                 Gun2Index = 7,
                                 Gun3 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value,
                                 Gun3Index = 8,
                                 Gun4 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value,
                                 Gun4Index = 9,
                                 Gun5 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_9" select new XElement(y.Element("Value"))).First().Value,
                                 Gun5Index = 10,
                                 Gun6 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_10" select new XElement(y.Element("Value"))).First().Value,
                                 Gun6Index = 11,
                                 Gun7 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_11" select new XElement(y.Element("Value"))).First().Value,
                                 Gun7Index = 12,
                                 Gun8 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_12" select new XElement(y.Element("Value"))).First().Value,
                                 Gun8Index = 13,
                                 Gun9 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_13" select new XElement(y.Element("Value"))).First().Value,
                                 Gun9Index = 14,
                                 Gun10 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_14" select new XElement(y.Element("Value"))).First().Value,
                                 Gun10Index = 15,
                                 Gun11 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_15" select new XElement(y.Element("Value"))).First().Value,
                                 Gun11Index = 16,
                                 Gun12 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_16" select new XElement(y.Element("Value"))).First().Value,
                                 Gun12Index = 17,
                                 Gun13 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_17" select new XElement(y.Element("Value"))).First().Value,
                                 Gun13Index = 18,
                                 Gun14 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_18" select new XElement(y.Element("Value"))).First().Value,
                                 Gun14Index = 19,
                                 Gun15 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_19" select new XElement(y.Element("Value"))).First().Value,
                                 Gun15Index = 20,
                                 Gun16 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_20" select new XElement(y.Element("Value"))).First().Value,
                                 Gun16Index = 21,
                                 Gun17 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_21" select new XElement(y.Element("Value"))).First().Value,
                                 Gun17Index = 22,
                                 Gun18 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_22" select new XElement(y.Element("Value"))).First().Value,
                                 Gun18Index = 23,
                                 Gun19 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_23" select new XElement(y.Element("Value"))).First().Value,
                                 Gun19Index = 24,
                                 Gun20 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_24" select new XElement(y.Element("Value"))).First().Value,
                                 Gun20Index = 25,
                                 Gun21 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_25" select new XElement(y.Element("Value"))).First().Value,
                                 Gun21Index = 26,
                                 Gun22 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_26" select new XElement(y.Element("Value"))).First().Value,
                                 Gun22Index = 27,
                                 Gun23 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_27" select new XElement(y.Element("Value"))).First().Value,
                                 Gun23Index = 28,
                                 Gun24 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_28" select new XElement(y.Element("Value"))).First().Value,
                                 Gun24Index = 29,
                                 Gun25 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_29" select new XElement(y.Element("Value"))).First().Value,
                                 Gun25Index = 30,
                                 Gun26 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_30" select new XElement(y.Element("Value"))).First().Value,
                                 Gun26Index = 31,
                                 Gun27 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_31" select new XElement(y.Element("Value"))).First().Value,
                                 Gun27Index = 32,
                                 Gun28 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_32" select new XElement(y.Element("Value"))).First().Value,
                                 Gun28Index = 33,
                                 Gun29 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_33" select new XElement(y.Element("Value"))).First().Value,
                                 Gun29Index = 34,
                                 Gun30 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_34" select new XElement(y.Element("Value"))).First().Value,
                                 Gun30Index = 35,
                                 Gun31 = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_35" select new XElement(y.Element("Value"))).First().Value,
                                 Gun31Index = 36,
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
                var customrenk = table.Where(x => x.Gun1 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun1Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun2 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun2Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun3 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun3Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun4 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun4Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun5 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun5Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun6 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun6Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun7 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun7Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun8 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun8Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun9 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun9Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun10 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun10Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun11 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun11Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun12 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun12Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun13 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun13Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun13 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun13Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun14 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun14Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun15 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun15Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun16 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun16Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun17 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun17Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun18 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun18Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun19 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun19Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun20 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun20Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun21 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun21Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun22 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun22Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun23 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun23Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun24 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun24Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun25 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun25Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun26 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun26Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun27 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun27Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun28 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun28Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun29 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun29Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun30 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun30Index, Yesil);
                }

                customrenk = table.Where(x => x.Gun31 == "X").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun31Index, Yesil);
                }

                #endregion

                #region Hafta Tatili Kırmızı Çevir.
                customrenk = table.Where(x => x.Gun1 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun1Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun2 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun2Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun3 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun3Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun4 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun4Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun5 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun5Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun6 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun6Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun7 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun7Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun8 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun8Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun9 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun9Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun10 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun10Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun11 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun11Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun12 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun12Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun13 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun13Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun13 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun13Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun14 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun14Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun15 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun15Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun16 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun16Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun17 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun17Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun18 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun18Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun19 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun19Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun20 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun20Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun21 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun21Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun22 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun22Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun23 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun23Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun24 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun24Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun25 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun25Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun26 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun26Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun27 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun27Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun28 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun28Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun29 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun29Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun30 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun30Index, Kirmizi);
                }

                customrenk = table.Where(x => x.Gun31 == "HT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun31Index, Kirmizi);
                }

                #endregion

                #region Raporluları Maviye Çevir.
                customrenk = table.Where(x => x.Gun1 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun1Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun2 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun2Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun3 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun3Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun4 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun4Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun5 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun5Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun6 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun6Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun7 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun7Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun8 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun8Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun9 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun9Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun10 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun10Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun11 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun11Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun12 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun12Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun13 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun13Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun13 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun13Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun14 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun14Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun15 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun15Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun16 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun16Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun17 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun17Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun18 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun18Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun19 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun19Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun20 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun20Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun21 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun21Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun22 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun22Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun23 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun23Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun24 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun24Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun25 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun25Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun26 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun26Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun27 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun27Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun28 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun28Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun29 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun29Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun30 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun30Index, Mavi);
                }

                customrenk = table.Where(x => x.Gun31 == "R").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun31Index, Mavi);
                }

                #endregion

                #region Ücretsiz İzinlileri Sarıya Çevir.
                customrenk = table.Where(x => x.Gun1 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun1Index, Sari);
                }

                customrenk = table.Where(x => x.Gun2 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun2Index, Sari);
                }

                customrenk = table.Where(x => x.Gun3 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun3Index, Sari);
                }

                customrenk = table.Where(x => x.Gun4 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun4Index, Sari);
                }

                customrenk = table.Where(x => x.Gun5 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun5Index, Sari);
                }

                customrenk = table.Where(x => x.Gun6 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun6Index, Sari);
                }

                customrenk = table.Where(x => x.Gun7 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun7Index, Sari);
                }

                customrenk = table.Where(x => x.Gun8 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun8Index, Sari);
                }

                customrenk = table.Where(x => x.Gun9 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun9Index, Sari);
                }

                customrenk = table.Where(x => x.Gun10 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun10Index, Sari);
                }

                customrenk = table.Where(x => x.Gun11 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun11Index, Sari);
                }

                customrenk = table.Where(x => x.Gun12 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun12Index, Sari);
                }

                customrenk = table.Where(x => x.Gun13 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun13Index, Sari);
                }

                customrenk = table.Where(x => x.Gun13 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun13Index, Sari);
                }

                customrenk = table.Where(x => x.Gun14 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun14Index, Sari);
                }

                customrenk = table.Where(x => x.Gun15 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun15Index, Sari);
                }

                customrenk = table.Where(x => x.Gun16 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun16Index, Sari);
                }

                customrenk = table.Where(x => x.Gun17 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun17Index, Sari);
                }

                customrenk = table.Where(x => x.Gun18 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun18Index, Sari);
                }

                customrenk = table.Where(x => x.Gun19 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun19Index, Sari);
                }

                customrenk = table.Where(x => x.Gun20 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun20Index, Sari);
                }

                customrenk = table.Where(x => x.Gun21 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun21Index, Sari);
                }

                customrenk = table.Where(x => x.Gun22 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun22Index, Sari);
                }

                customrenk = table.Where(x => x.Gun23 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun23Index, Sari);
                }

                customrenk = table.Where(x => x.Gun24 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun24Index, Sari);
                }

                customrenk = table.Where(x => x.Gun25 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun25Index, Sari);
                }

                customrenk = table.Where(x => x.Gun26 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun26Index, Sari);
                }

                customrenk = table.Where(x => x.Gun27 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun27Index, Sari);
                }

                customrenk = table.Where(x => x.Gun28 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun28Index, Sari);
                }

                customrenk = table.Where(x => x.Gun29 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun29Index, Sari);
                }

                customrenk = table.Where(x => x.Gun30 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun30Index, Sari);
                }

                customrenk = table.Where(x => x.Gun31 == "UC").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun31Index, Sari);
                }

                #endregion

                #region Resmi Tatilleri SarıTuruncuya Çevir.
                customrenk = table.Where(x => x.Gun1 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun1Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun2 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun2Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun3 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun3Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun4 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun4Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun5 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun5Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun6 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun6Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun7 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun7Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun8 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun8Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun9 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun9Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun10 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun10Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun11 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun11Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun12 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun12Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun13 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun13Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun13 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun13Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun14 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun14Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun15 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun15Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun16 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun16Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun17 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun17Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun18 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun18Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun19 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun19Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun20 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun20Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun21 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun21Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun22 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun22Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun23 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun23Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun24 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun24Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun25 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun25Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun26 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun26Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun27 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun27Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun28 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun28Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun29 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun29Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun30 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun30Index, Turuncu);
                }

                customrenk = table.Where(x => x.Gun31 == "RT").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun31Index, Turuncu);
                }

                #endregion

                #region Yıllık İzinleri Griye Çevir.
                customrenk = table.Where(x => x.Gun1 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun1Index, Gri);
                }

                customrenk = table.Where(x => x.Gun2 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun2Index, Gri);
                }

                customrenk = table.Where(x => x.Gun3 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun3Index, Gri);
                }

                customrenk = table.Where(x => x.Gun4 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun4Index, Gri);
                }

                customrenk = table.Where(x => x.Gun5 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun5Index, Gri);
                }

                customrenk = table.Where(x => x.Gun6 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun6Index, Gri);
                }

                customrenk = table.Where(x => x.Gun7 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun7Index, Gri);
                }

                customrenk = table.Where(x => x.Gun8 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun8Index, Gri);
                }

                customrenk = table.Where(x => x.Gun9 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun9Index, Gri);
                }

                customrenk = table.Where(x => x.Gun10 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun10Index, Gri);
                }

                customrenk = table.Where(x => x.Gun11 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun11Index, Gri);
                }

                customrenk = table.Where(x => x.Gun12 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun12Index, Gri);
                }

                customrenk = table.Where(x => x.Gun13 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun13Index, Gri);
                }

                customrenk = table.Where(x => x.Gun13 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun13Index, Gri);
                }

                customrenk = table.Where(x => x.Gun14 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun14Index, Gri);
                }

                customrenk = table.Where(x => x.Gun15 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun15Index, Gri);
                }

                customrenk = table.Where(x => x.Gun16 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun16Index, Gri);
                }

                customrenk = table.Where(x => x.Gun17 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun17Index, Gri);
                }

                customrenk = table.Where(x => x.Gun18 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun18Index, Gri);
                }

                customrenk = table.Where(x => x.Gun19 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun19Index, Gri);
                }

                customrenk = table.Where(x => x.Gun20 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun20Index, Gri);
                }

                customrenk = table.Where(x => x.Gun21 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun21Index, Gri);
                }

                customrenk = table.Where(x => x.Gun22 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun22Index, Gri);
                }

                customrenk = table.Where(x => x.Gun23 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun23Index, Gri);
                }

                customrenk = table.Where(x => x.Gun24 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun24Index, Gri);
                }

                customrenk = table.Where(x => x.Gun25 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun25Index, Gri);
                }

                customrenk = table.Where(x => x.Gun26 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun26Index, Gri);
                }

                customrenk = table.Where(x => x.Gun27 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun27Index, Gri);
                }

                customrenk = table.Where(x => x.Gun28 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun28Index, Gri);
                }

                customrenk = table.Where(x => x.Gun29 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun29Index, Gri);
                }

                customrenk = table.Where(x => x.Gun30 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun30Index, Gri);
                }

                customrenk = table.Where(x => x.Gun31 == "YI").ToList();


                foreach (var item in customrenk)
                {
                    oMatrix.CommonSetting.SetCellBackColor(item.INDEX, item.Gun31Index, Gri);
                }

                #endregion

                xmlColor = oMatrix.SerializeAsXML(SAPbouiCOM.BoMatrixXmlSelect.mxs_All);
            }
            catch (Exception)
            {
            }
            finally
            {
                frmGunlukPersonelPlanlama2.Freeze(false);
            }
        }

        public bool MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.MenuUID == "1282" && !pVal.BeforeAction)
            {
                kololariSifirla();
                CalisanListele();
                oComboYil.Select(DateTime.Now.Year.ToString());

                //try
                //{
                //    frmGunlukPersonelPlanlama.Close();
                //}
                //catch (Exception)
                //{
                //}

                //Handler.SAPApplication.ActivateMenuItem("GunPersPl");
                ////frmGunlukPersonelPlanlama.Mode = BoFormMode.fm_ADD_MODE;
                ////frmGunlukPersonelPlanlama.DataSources.DBDataSources.Item(0).Clear();
                ////frmGunlukPersonelPlanlama.DataSources.DBDataSources.Item(1).Clear();
                ////oEditDocEntry.Item.Click();
                ////CalisanListele();
                //renklendir(true);
            }

            return BubbleEvent;
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        List<GunlukPersonelPlanlama3> gunlukPersonelPlanlama3s = new List<GunlukPersonelPlanlama3>();
        public class GunlukPersonelPlanlama3
        {
            public int KullaniciId { get; set; }

            public string Tarih { get; set; }

            public string Durum { get; set; }

            public string Bolum1 { get; set; }

            public string Bolum2 { get; set; }

            public string Bolum3 { get; set; }
        }
    }
    public static class Extension
    {
        public static bool IsNumeric(this string s)
        {
            int output;
            return int.TryParse(s, out output);
        }
    }
}