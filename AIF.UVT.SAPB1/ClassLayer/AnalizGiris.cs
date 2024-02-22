using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.UVT.SAPB1.Models;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class AnalizGiris : IUserForm, IMenuEvents, IRightEvents
    {
        [ItemAtt(AIFConn.AnalizGirisUID)]
        public SAPbouiCOM.Form frmAnaliz;

        [ItemAtt("Item_3")]
        public SAPbouiCOM.Folder oFolderSonUrunKimyasal;

        [ItemAtt("Item_6")]
        public SAPbouiCOM.Matrix oMatrixSonUrunKimyasal;

        [ItemAtt("Item_8")]
        public SAPbouiCOM.Matrix oMatrixDuyusalFonskiyonel;

        [ItemAtt("Item_9")]
        public SAPbouiCOM.Matrix oMatrixMikrobiyolojik;

        [ItemAtt("Item_14")]
        public SAPbouiCOM.EditText oEditUrunKodu;

        [ItemAtt("Item_1")]
        public SAPbouiCOM.EditText oEditUrunAdi;


        [ItemAtt("Item_5")]
        public SAPbouiCOM.Folder oFolderMikrobiyolojik;

        private List<AnalizPartiler> analizPartilers = new List<AnalizPartiler>();
        private List<AnalizSecilmisPartiler> analizSecilmisPartilers = new List<AnalizSecilmisPartiler>();

        string secilmisUygunsuzlukNedeni = "";
        string matrixUID = "";
        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.AnalizGirisUID, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.AnalizGirisFrmXML));
            Functions.CreateUserOrSystemFormComponent<AnalizGiris>(AIFConn.AnalizGiris);

            InitForms();
        }

        public void InitForms()
        {
            try
            {
                frmAnaliz.EnableMenu("1283", false);
                frmAnaliz.EnableMenu("1284", false);
                frmAnaliz.EnableMenu("1286", false);

                oFolderMikrobiyolojik.Item.Visible = false;

                oFolderSonUrunKimyasal.Select();

                SAPbouiCOM.Column oCol = (SAPbouiCOM.Column)oMatrixMikrobiyolojik.Columns.Item("Col_4");

                string sql = "Select \"U_Id\", \"U_Bolum\" from \"@AIF_ANLZGIRISBOLUM\" ";
                ConstVariables.oRecordset.DoQuery(sql);

                if (ConstVariables.oRecordset.RecordCount > 0)
                {
                    while (!ConstVariables.oRecordset.EoF)
                    {
                        oCol.ValidValues.Add(ConstVariables.oRecordset.Fields.Item("U_Id").Value.ToString(), ConstVariables.oRecordset.Fields.Item("U_Bolum").Value.ToString());
                        ConstVariables.oRecordset.MoveNext();
                    }
                }

                oMatrixSonUrunKimyasal.AddRow();
                oMatrixDuyusalFonskiyonel.AddRow();
                oMatrixMikrobiyolojik.AddRow();

                oMatrixSonUrunKimyasal.AutoResizeColumns();
                oMatrixDuyusalFonskiyonel.AutoResizeColumns();
                oMatrixMikrobiyolojik.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Hata oluştu" + ex.Message);
            }
        }

        private void secilmispartilerimatristenal()
        {
            List<AnalizSecilmisPartiler> oncekiBelgelerdekiPartiler = new List<AnalizSecilmisPartiler>();
            analizSecilmisPartilers = new List<AnalizSecilmisPartiler>();
            string xml = "";
            string xmlTabloVerileri = "";
            XDocument xDoc = new XDocument();
            XNamespace ns = null;

            if (matrixUID != "" && matrixUID == "Item_6")
            {
                xml = oMatrixSonUrunKimyasal.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                #region Daha önceki belgelerdeki partileri getirir.
                if (ConstVariables.oRecordset.RecordCount > 0)
                {
                    string kmyasal = "SELECT \"U_Partino\"  FROM \"@AIF_ANALIZGIRIS1\"";
                    ConstVariables.oRecordset.DoQuery(kmyasal);

                    xmlTabloVerileri = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                    xDoc = XDocument.Parse(xmlTabloVerileri);
                    ns = "http://www.sap.com/SBO/SDK/DI";

                    oncekiBelgelerdekiPartiler = (from t in xDoc.Descendants(ns + "Row")
                                                  select new AnalizSecilmisPartiler
                                                  {
                                                      PartiNumarasi = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Partino" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                  }).ToList();
                    oncekiBelgelerdekiPartiler.RemoveAll(x => x.PartiNumarasi == "");  //satırdaki parti no boş olanı siler 
                }
                #endregion

            }
            else if (matrixUID != "" && matrixUID == "Item_8")
            {
                xml = oMatrixDuyusalFonskiyonel.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                #region Daha önceki belgelerdeki partileri getirir.
                string duyusal = "SELECT \"U_Partino\"  FROM \"@AIF_ANALIZGIRIS2\"";
                ConstVariables.oRecordset.DoQuery(duyusal);

                if (ConstVariables.oRecordset.RecordCount > 0)
                {
                    xmlTabloVerileri = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                    xDoc = XDocument.Parse(xmlTabloVerileri);
                    ns = "http://www.sap.com/SBO/SDK/DI";

                    oncekiBelgelerdekiPartiler = (from t in xDoc.Descendants(ns + "Row")
                                                  select new AnalizSecilmisPartiler
                                                  {
                                                      PartiNumarasi = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Partino" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                  }).ToList();
                    oncekiBelgelerdekiPartiler.RemoveAll(x => x.PartiNumarasi == "");  //satırdaki parti no boş olanı siler 
                }
                #endregion
            }
            else if (matrixUID != "" && matrixUID == "Item_9")
            {
                xml = oMatrixMikrobiyolojik.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                #region Daha önceki belgelerdeki partileri getirir.
                string mikrobiyolojik = "SELECT \"U_Partino\"  FROM \"@AIF_ANALIZGIRIS3\"";
                ConstVariables.oRecordset.DoQuery(mikrobiyolojik);

                if (ConstVariables.oRecordset.RecordCount > 0)
                {
                    xmlTabloVerileri = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                    xDoc = XDocument.Parse(xmlTabloVerileri);
                    ns = "http://www.sap.com/SBO/SDK/DI";

                    oncekiBelgelerdekiPartiler = (from t in xDoc.Descendants(ns + "Row")
                                                  select new AnalizSecilmisPartiler
                                                  {
                                                      PartiNumarasi = (from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Partino" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                  }).ToList();
                    oncekiBelgelerdekiPartiler.RemoveAll(x => x.PartiNumarasi == "");  //satırdaki parti no boş olanı siler
                }
                #endregion
            }

            analizSecilmisPartilers = (from x in XDocument.Parse(xml).Descendants("Row")
                                       select new AnalizSecilmisPartiler
                                       {
                                           PartiNumarasi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                       }).ToList();

            analizSecilmisPartilers.RemoveAll(x => x.PartiNumarasi == "");

            if (oncekiBelgelerdekiPartiler.Count > 0)
            {
                analizSecilmisPartilers.AddRange(oncekiBelgelerdekiPartiler);  //ikinci ekrana göndereceğimiz listeye ekranda olan matrixteki verileri doldururz. Bu satır ile daha önce başka belgelerde seçilmiş olanları ekleriz.
            }
            //foreach (var item in partilerMatrix)
            //{
            //    if (item.PartiNumarasi != "")
            //    {
            //        analizSecilmisPartilers.Add(item.PartiNumarasi);
            //    }
            //}
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
                    if (BusinessObjectInfo.BeforeAction)
                    {
                        string sonsatir1 = frmAnaliz.DataSources.DBDataSources.Item(1).GetValue("U_PartiNo", frmAnaliz.DataSources.DBDataSources.Item(1).Size - 1);
                        string sonsatir2 = frmAnaliz.DataSources.DBDataSources.Item(2).GetValue("U_PartiNo", frmAnaliz.DataSources.DBDataSources.Item(2).Size - 1);
                        string sonsatir3 = frmAnaliz.DataSources.DBDataSources.Item(3).GetValue("U_PartiNo", frmAnaliz.DataSources.DBDataSources.Item(3).Size - 1);

                        if (sonsatir1 == "")
                        {
                            frmAnaliz.DataSources.DBDataSources.Item(1).RemoveRecord(frmAnaliz.DataSources.DBDataSources.Item(1).Size - 1);
                        }

                        if (sonsatir2 == "")
                        {
                            frmAnaliz.DataSources.DBDataSources.Item(2).RemoveRecord(frmAnaliz.DataSources.DBDataSources.Item(2).Size - 1);
                        }

                        if (sonsatir3 == "")
                        {
                            frmAnaliz.DataSources.DBDataSources.Item(3).RemoveRecord(frmAnaliz.DataSources.DBDataSources.Item(3).Size - 1);
                        }
                    }
                    break;
                case BoEventTypes.et_FORM_DATA_UPDATE:
                    if (BusinessObjectInfo.BeforeAction)
                    {
                        string sonsatir1 = frmAnaliz.DataSources.DBDataSources.Item(1).GetValue("U_PartiNo", frmAnaliz.DataSources.DBDataSources.Item(1).Size - 1);
                        string sonsatir2 = frmAnaliz.DataSources.DBDataSources.Item(2).GetValue("U_PartiNo", frmAnaliz.DataSources.DBDataSources.Item(2).Size - 1);
                        string sonsatir3 = frmAnaliz.DataSources.DBDataSources.Item(3).GetValue("U_PartiNo", frmAnaliz.DataSources.DBDataSources.Item(3).Size - 1);

                        if (sonsatir1 == "")
                        {
                            frmAnaliz.DataSources.DBDataSources.Item(1).RemoveRecord(frmAnaliz.DataSources.DBDataSources.Item(1).Size - 1);
                        }

                        if (sonsatir2 == "")
                        {
                            frmAnaliz.DataSources.DBDataSources.Item(2).RemoveRecord(frmAnaliz.DataSources.DBDataSources.Item(2).Size - 1);
                        }

                        if (sonsatir3 == "")
                        {
                            frmAnaliz.DataSources.DBDataSources.Item(3).RemoveRecord(frmAnaliz.DataSources.DBDataSources.Item(3).Size - 1);
                        }
                    }
                    break;
                case BoEventTypes.et_FORM_DATA_DELETE:
                    break;
                case BoEventTypes.et_FORM_DATA_LOAD:
                    if (!BusinessObjectInfo.BeforeAction)
                    {
                        frmAnaliz.DataSources.DBDataSources.Item(1).Clear();
                        frmAnaliz.DataSources.DBDataSources.Item(2).Clear();
                        frmAnaliz.DataSources.DBDataSources.Item(3).Clear();

                        oMatrixSonUrunKimyasal.AddRow();
                        oMatrixDuyusalFonskiyonel.AddRow();
                        oMatrixMikrobiyolojik.AddRow();

                        oMatrixSonUrunKimyasal.AutoResizeColumns();
                        oMatrixDuyusalFonskiyonel.AutoResizeColumns();
                        oMatrixMikrobiyolojik.AutoResizeColumns();
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
                        frmAnaliz.DataSources.DBDataSources.Item(1).Clear();
                        frmAnaliz.DataSources.DBDataSources.Item(2).Clear();
                        frmAnaliz.DataSources.DBDataSources.Item(3).Clear();

                        oMatrixSonUrunKimyasal.AddRow();
                        oMatrixDuyusalFonskiyonel.AddRow();
                        oMatrixMikrobiyolojik.AddRow();

                        oMatrixSonUrunKimyasal.AutoResizeColumns();
                        oMatrixDuyusalFonskiyonel.AutoResizeColumns();
                        oMatrixMikrobiyolojik.AutoResizeColumns();
                    }
                    //if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    //{
                    //    if (itemUID == "Item" && pVal.ColUID == "")
                    //    {
                    //        if (silinecekler.Count > 0)
                    //        {
                    //            foreach (var item in silinecekler)
                    //            {
                    //                if (item != "")
                    //                {
                    //                    ConstVariables.oRecordset.DoQuery("Delete from \"@AIF_ANALIZGIRIS1\" where \"DocEntry\" = '" + item + "'");
                    //                }
                    //            }
                    //        }

                    //        silinecekler = new List<string>(); 
                    //    }
                    //}
                    break;
                case BoEventTypes.et_KEY_DOWN:
                    break;
                case BoEventTypes.et_GOT_FOCUS:
                    break;
                case BoEventTypes.et_LOST_FOCUS:
                  
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    if (pVal.ItemUID == "Item_9" && pVal.ColUID == "Col_4" && pVal.BeforeAction)
                    {
                        secilmisUygunsuzlukNedeni = ((SAPbouiCOM.ComboBox)oMatrixMikrobiyolojik.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific).Value.ToString();
                    }
                    else if (pVal.ItemUID == "Item_9" && pVal.ColUID == "Col_4" && !pVal.BeforeAction)
                    {
                        string uygunsuzlukNedeni = ((SAPbouiCOM.ComboBox)oMatrixMikrobiyolojik.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific).Value.ToString();

                        if (secilmisUygunsuzlukNedeni != "" && uygunsuzlukNedeni != secilmisUygunsuzlukNedeni)
                        {
                            ((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item("Col_5").Cells.Item(pVal.Row).Specific).Value = "";
                        }
                    }
                    break;
                case BoEventTypes.et_CLICK:
                    break;
                case BoEventTypes.et_DOUBLE_CLICK:
                    if (pVal.ItemUID == "Item_6" && pVal.ColUID == "Col_0" && !pVal.BeforeAction)
                    {
                        matrixUID = pVal.ItemUID;
                        if (oEditUrunKodu.Value.ToString() == "")
                        {
                            Handler.SAPApplication.MessageBox("Lütfen Ürün Kodu seçimi yapınız.");
                            return false;
                        }

                        if (oEditUrunKodu.Value.ToString() != "")
                        {
                            secilmispartilerimatristenal();

                            if (analizSecilmisPartilers.Count > 0)
                            {
                                AIFConn.AnlzGrsSec.LoadForms(oEditUrunKodu.Value.ToString(), "KimyasalAnaliz", analizPartilers, analizSecilmisPartilers);

                            }
                            else
                            {
                                AIFConn.AnlzGrsSec.LoadForms(oEditUrunKodu.Value.ToString(), "KimyasalAnaliz", analizPartilers, analizSecilmisPartilers);

                            }
                        }
                    }
                    else if (pVal.ItemUID == "Item_8" && pVal.ColUID == "Col_0" && !pVal.BeforeAction)
                    {
                        matrixUID = pVal.ItemUID;

                        if (oEditUrunKodu.Value.ToString() == "")
                        {
                            Handler.SAPApplication.MessageBox("Lütfen Ürün Kodu seçimi yapınız.");
                            return false;
                        }

                        secilmispartilerimatristenal();

                        if (oEditUrunKodu.Value.ToString() != "")
                        {
                            AIFConn.AnlzGrsSec.LoadForms(oEditUrunKodu.Value.ToString(), "DuyusalAnaliz", analizPartilers, analizSecilmisPartilers);
                        }
                    }
                    else if (pVal.ItemUID == "Item_9" && pVal.ColUID == "Col_0" && !pVal.BeforeAction)
                    {
                        matrixUID = pVal.ItemUID;

                        if (oEditUrunKodu.Value.ToString() == "")
                        {
                            Handler.SAPApplication.MessageBox("Lütfen Ürün Kodu seçimi yapınız.");
                            return false;
                        }
                        secilmispartilerimatristenal();

                        if (oEditUrunKodu.Value.ToString() != "")
                        {
                            AIFConn.AnlzGrsSec.LoadForms(oEditUrunKodu.Value.ToString(), "MikrobiyolojikAnaliz", analizPartilers, analizSecilmisPartilers);
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
                    if (pVal.ItemUID == "Item_14" && pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                            oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                            SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                            oCFL = frmAnaliz.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                            SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                            SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                            SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                            oCFL.SetConditions(oEmptyConts);
                            oCons = oCFL.GetConditions();

                            oCon = oCons.Add();
                            oCon.Alias = "validFor";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = "Y";

                            oCFL.SetConditions(oCons);
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_14" && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                        if (oDataTable == null)
                        {
                            break;
                        }
                        string val = "";
                        try
                        {
                            val = oDataTable.GetValue("ItemCode", 0).ToString();
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            oEditUrunKodu.Value = val;
                            //((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value = val;
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            val = oDataTable.GetValue("ItemName", 0).ToString();
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            oEditUrunAdi.Value = val;
                            //((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = val;
                        }
                        catch (Exception)
                        {
                        }

                    }
                    else if (pVal.ItemUID == "Item_9" && pVal.ColUID == "Col_5" && pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                            oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                            SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                            oCFL = frmAnaliz.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                            SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                            SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                            SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                            oCFL.SetConditions(oEmptyConts);
                            oCons = oCFL.GetConditions();

                            string bolum = "";

                            if (((SAPbouiCOM.ComboBox)oMatrixMikrobiyolojik.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific).Value != "")
                            {
                                bolum = ((SAPbouiCOM.ComboBox)oMatrixMikrobiyolojik.Columns.Item("Col_4").Cells.Item(pVal.Row).Specific).Value.ToString();
                            }

                            oCon = oCons.Add();
                            oCon.Alias = "U_BolumId";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = bolum;

                            //oCon.Relationship = BoConditionRelationship.cr_AND;

                            //oCon = oCons.Add();
                            //oCon.Alias = "U_Aktif";
                            //oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            //oCon.CondVal = "1";

                            oCFL.SetConditions(oCons);
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_9" && pVal.ColUID == "Col_5" && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                        if (oDataTable == null)
                        {
                            break;
                        }
                        string val = "";
                        try
                        {
                            val = oDataTable.GetValue("U_Numune", 0).ToString();
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            //oEditUrunKodu.Value = val;
                            ((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item("Col_5").Cells.Item(pVal.Row).Specific).Value = val;
                        }
                        catch (Exception)
                        {
                        }

                    }
                    else if (pVal.ItemUID == "Item_6" && pVal.ColUID == "Col_13" && pVal.BeforeAction) //son ürün kimyasal analiz
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                        oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                        SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                        oCFL = frmAnaliz.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                        SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                        SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                        SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                        oCFL.SetConditions(oEmptyConts);
                        oCons = oCFL.GetConditions();

                        oCon = oCons.Add();
                        oCon.Alias = "Active";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";

                        oCFL.SetConditions(oCons);
                    }
                    else if ((pVal.ItemUID == "Item_6" || pVal.ItemUID == "Item_8" || pVal.ItemUID == "Item_9") && (pVal.ColUID == "Col_13" || pVal.ColUID == "Col_16" || pVal.ColUID == "Col_13") && !pVal.BeforeAction) //son ürün kimyasal analiz - //duyusal analiz- //mikrobiyolojik analz
                    {
                        #region kimyasal matrisi analiz yapan kişi
                        if (pVal.ItemUID == "Item_6" && pVal.ColUID == "Col_13")
                        {
                            try
                            {
                                SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                                string Val = "";
                                if (oDataTable == null)
                                {
                                    return false;
                                }
                                Val = oDataTable.GetValue("empID", 0).ToString();
                                try
                                {
                                    //oEditSorumlu.Value = Val;
                                    ((SAPbouiCOM.EditText)oMatrixSonUrunKimyasal.Columns.Item("Col_13").Cells.Item(pVal.Row).Specific).Value = Val;

                                }
                                catch (Exception)
                                {
                                }
                                Val = "";
                                try
                                {
                                    if (oDataTable.GetValue("middleName", 0).ToString() == "")
                                    {
                                        Val = oDataTable.GetValue("firstName", 0).ToString() + " " + oDataTable.GetValue("lastName", 0).ToString();

                                    }
                                    else
                                    {
                                        Val = oDataTable.GetValue("firstName", 0).ToString() + " " + oDataTable.GetValue("middleName", 0).ToString() + " " + oDataTable.GetValue("lastName", 0).ToString();
                                    }

                                    ((SAPbouiCOM.EditText)oMatrixSonUrunKimyasal.Columns.Item("Col_14").Cells.Item(pVal.Row).Specific).Value = Val;

                                }
                                catch (Exception ex)
                                {
                                }

                                //try
                                //{


                                //    if (oDataTable.GetValue("middleName", 0).ToString() != "")
                                //    {
                                //        Val = oDataTable.GetValue("firstName", 0).ToString() + " " + oDataTable.GetValue("middleName", 0).ToString() + " " + oDataTable.GetValue("lastName", 0).ToString();
                                //    }

                                //    ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = Val;

                                //}
                                //catch (Exception ex)
                                //{
                                //}
                            }
                            catch (Exception)
                            {
                            }
                            oMatrixSonUrunKimyasal.AutoResizeColumns();
                        }
                        #endregion

                        #region duyusal analz matrisi panelist adı
                        if (pVal.ItemUID == "Item_8" && pVal.ColUID == "Col_16")
                        {
                            try
                            {
                                SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                                string Val = "";
                                if (oDataTable == null)
                                {
                                    return false;
                                }
                                Val = oDataTable.GetValue("empID", 0).ToString();
                                try
                                {
                                    //oEditSorumlu.Value = Val;
                                    ((SAPbouiCOM.EditText)oMatrixDuyusalFonskiyonel.Columns.Item("Col_16").Cells.Item(pVal.Row).Specific).Value = Val;

                                }
                                catch (Exception)
                                {
                                }
                                Val = "";
                                try
                                {
                                    if (oDataTable.GetValue("middleName", 0).ToString() == "")
                                    {
                                        Val = oDataTable.GetValue("firstName", 0).ToString() + " " + oDataTable.GetValue("lastName", 0).ToString();

                                    }
                                    else
                                    {
                                        Val = oDataTable.GetValue("firstName", 0).ToString() + " " + oDataTable.GetValue("middleName", 0).ToString() + " " + oDataTable.GetValue("lastName", 0).ToString();
                                    }

                                    ((SAPbouiCOM.EditText)oMatrixDuyusalFonskiyonel.Columns.Item("Col_2").Cells.Item(pVal.Row).Specific).Value = Val;

                                }
                                catch (Exception ex)
                                {
                                }

                                //try
                                //{


                                //    if (oDataTable.GetValue("middleName", 0).ToString() != "")
                                //    {
                                //        Val = oDataTable.GetValue("firstName", 0).ToString() + " " + oDataTable.GetValue("middleName", 0).ToString() + " " + oDataTable.GetValue("lastName", 0).ToString();
                                //    }

                                //    ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = Val;

                                //}
                                //catch (Exception ex)
                                //{
                                //}
                            }
                            catch (Exception)
                            {
                            }
                            oMatrixDuyusalFonskiyonel.AutoResizeColumns();
                        }
                        #endregion

                        #region mikrobiyolojik analz matrisi analiz yapan
                        if (pVal.ItemUID == "Item_9" && pVal.ColUID == "Col_13")
                        {
                            try
                            {
                                SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                                string Val = "";
                                if (oDataTable == null)
                                {
                                    return false;
                                }
                                Val = oDataTable.GetValue("empID", 0).ToString();
                                try
                                {
                                    //oEditSorumlu.Value = Val;
                                    ((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item("Col_13").Cells.Item(pVal.Row).Specific).Value = Val;

                                }
                                catch (Exception)
                                {
                                }
                                Val = "";
                                try
                                {
                                    if (oDataTable.GetValue("middleName", 0).ToString() == "")
                                    {
                                        Val = oDataTable.GetValue("firstName", 0).ToString() + " " + oDataTable.GetValue("lastName", 0).ToString();

                                    }
                                    else
                                    {
                                        Val = oDataTable.GetValue("firstName", 0).ToString() + " " + oDataTable.GetValue("middleName", 0).ToString() + " " + oDataTable.GetValue("lastName", 0).ToString();
                                    }

                                    ((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item("Col_14").Cells.Item(pVal.Row).Specific).Value = Val;

                                }
                                catch (Exception ex)
                                {
                                }

                                //try
                                //{


                                //    if (oDataTable.GetValue("middleName", 0).ToString() != "")
                                //    {
                                //        Val = oDataTable.GetValue("firstName", 0).ToString() + " " + oDataTable.GetValue("middleName", 0).ToString() + " " + oDataTable.GetValue("lastName", 0).ToString();
                                //    }

                                //    ((SAPbouiCOM.EditText)oMatrixDetay.Columns.Item("Col_1").Cells.Item(pVal.Row).Specific).Value = Val;

                                //}
                                //catch (Exception ex)
                                //{
                                //}
                            }
                            catch (Exception)
                            {
                            }
                            oMatrixMikrobiyolojik.AutoResizeColumns();
                        }
                        #endregion
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
        List<string> silinecekler = new List<string>();
        string itemUID = "";
        public void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.MenuUID == "1282" && !pVal.BeforeAction)
            {
                try
                {
                    oMatrixSonUrunKimyasal.AddRow();
                    oMatrixDuyusalFonskiyonel.AddRow();
                    oMatrixMikrobiyolojik.AddRow();

                    //oMatrixSonUrunKimyasal.AutoResizeColumns();
                    //oMatrixDuyusalFonskiyonel.AutoResizeColumns();
                    //oMatrixMikrobiyolojik.AutoResizeColumns();
                }
                catch (Exception)
                {
                }
            }

            try
            {
                if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && pVal.BeforeAction)
                {

                    if (itemUID == "Item_6")
                    {
                        int row = oMatrixSonUrunKimyasal.GetNextSelectedRow();
                        if (row != -1)
                        {
                            //if (((SAPbouiCOM.EditText)oMatrixSonUrunKimyasal.Columns.Item("Col_0").Cells.Item(row).Specific).Value != "")
                            //{
                            //    silinecekler.Add(((SAPbouiCOM.EditText)oMatrixSonUrunKimyasal.Columns.Item("Col_0").Cells.Item(row).Specific).Value);
                            //}

                            oMatrixSonUrunKimyasal.DeleteRow(row);

                            if (frmAnaliz.Mode == BoFormMode.fm_OK_MODE)
                            {
                                frmAnaliz.Mode = BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                    }
                    else if (itemUID == "Item_8")
                    {
                        int row = oMatrixDuyusalFonskiyonel.GetNextSelectedRow();
                        if (row != -1)
                        {
                            //if (((SAPbouiCOM.EditText)oMatrixDuyusalFonskiyonel.Columns.Item("Col_0").Cells.Item(row).Specific).Value != "")
                            //{
                            //    silinecekler.Add(((SAPbouiCOM.EditText)oMatrixDuyusalFonskiyonel.Columns.Item("Col_0").Cells.Item(row).Specific).Value);
                            //}

                            oMatrixDuyusalFonskiyonel.DeleteRow(row);

                            if (frmAnaliz.Mode == BoFormMode.fm_OK_MODE)
                            {
                                frmAnaliz.Mode = BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                    }
                    else if (itemUID == "Item_9")
                    {
                        int row = oMatrixMikrobiyolojik.GetNextSelectedRow();
                        if (row != -1)
                        {
                            //if (((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item("Col_0").Cells.Item(row).Specific).Value != "")
                            //{
                            //    silinecekler.Add(((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item("Col_0").Cells.Item(row).Specific).Value);
                            //}

                            oMatrixMikrobiyolojik.DeleteRow(row);

                            if (frmAnaliz.Mode == BoFormMode.fm_OK_MODE)
                            {
                                frmAnaliz.Mode = BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                    }
                }
                //else if (pVal.MenuUID == "AIFRGHTCLK_AddRow" && pVal.BeforeAction)
                //{
                //    frmBankaEslestirme.DataSources.DBDataSources.Item("@AIF_BNKESLESTIRME").Clear();
                //    oMatrix.AddRow();

                //}
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
                        itemUID = eventInfo.ItemUID;
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
                    //Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_AddRow");
                    return;
                }


                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_DeleteRow");
                    //Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_AddRow");
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

                    //try
                    //{

                    //    oCreationPackage.UniqueID = "AIFRGHTCLK_AddRow";

                    //    oCreationPackage.String = "Satır Ekle";

                    //    oCreationPackage.Enabled = true;

                    //    oMenuItem = Handler.SAPApplication.Menus.Item("1280");

                    //    oMenus = oMenuItem.SubMenus;

                    //    oMenus.AddEx(oCreationPackage);
                    //}
                    //catch (Exception)
                    //{
                    //}
                }
                catch (Exception ex)
                {
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void partileriGetir(List<AnalizPartiler> analizPartilers)
        {
            try
            {
                //frmAnaliz.Freeze(true);

                if (analizPartilers != null && analizPartilers.Count > 0)
                {

                    if (analizPartilers[0].AnalizAdi == "KimyasalAnaliz")
                    {
                        foreach (var item in analizPartilers)
                        {
                            ((SAPbouiCOM.EditText)oMatrixSonUrunKimyasal.Columns.Item("Col_0").Cells.Item(oMatrixSonUrunKimyasal.RowCount).Specific).Value = item.PartiNumarasi;

                            if (item.KabulTarihi != "")
                            {
                                DateTime dtUretim = Convert.ToDateTime(item.KabulTarihi);
                                ((SAPbouiCOM.EditText)oMatrixSonUrunKimyasal.Columns.Item("Col_1").Cells.Item(oMatrixSonUrunKimyasal.RowCount).Specific).Value = dtUretim.ToString("yyyyMMdd");
                            }

                            if (item.GecerlilikSonu != "")
                            {
                                DateTime dtSonKullanim = Convert.ToDateTime(item.GecerlilikSonu);
                                ((SAPbouiCOM.EditText)oMatrixSonUrunKimyasal.Columns.Item("Col_2").Cells.Item(oMatrixSonUrunKimyasal.RowCount).Specific).Value = dtSonKullanim.ToString("yyyyMMdd");
                            }

                            oMatrixSonUrunKimyasal.AddRow();
                        }

                        oMatrixSonUrunKimyasal.AutoResizeColumns();
                    }

                    if (analizPartilers[0].AnalizAdi == "DuyusalAnaliz")
                    {
                        foreach (var item in analizPartilers)
                        {
                            ((SAPbouiCOM.EditText)oMatrixDuyusalFonskiyonel.Columns.Item("Col_0").Cells.Item(oMatrixDuyusalFonskiyonel.RowCount).Specific).Value = item.PartiNumarasi;
                            if (item.KabulTarihi != "")
                            {
                                DateTime dtUretim = Convert.ToDateTime(item.KabulTarihi);
                                ((SAPbouiCOM.EditText)oMatrixDuyusalFonskiyonel.Columns.Item("Col_1").Cells.Item(oMatrixDuyusalFonskiyonel.RowCount).Specific).Value = dtUretim.ToString("yyyyMMdd");
                            }
                            if (item.GecerlilikSonu != "")
                            {
                                DateTime dtSonKullanim = Convert.ToDateTime(item.GecerlilikSonu);
                                ((SAPbouiCOM.EditText)oMatrixDuyusalFonskiyonel.Columns.Item("Col_15").Cells.Item(oMatrixDuyusalFonskiyonel.RowCount).Specific).Value = dtSonKullanim.ToString("yyyyMMdd");
                            }
                            oMatrixDuyusalFonskiyonel.AddRow();
                        }

                        oMatrixDuyusalFonskiyonel.AutoResizeColumns();
                    }

                    if (analizPartilers[0].AnalizAdi == "MikrobiyolojikAnaliz")
                    {
                        foreach (var item in analizPartilers)
                        {
                            //SetCellWithoutValidation bu method arkasına bir olay var ise yapmadan yazmaya yarar mesela bu kolon cfl olsaydı bütün olayları işletip yazmayacaktı direk yazıyor olacaktı.
                            //oMatrixMikrobiyolojik.SetCellWithoutValidation(oMatrixMikrobiyolojik.RowCount, "Col_0", item.PartiNumarasi); //Parti No kolonu
                            ((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item("Col_0").Cells.Item(oMatrixMikrobiyolojik.RowCount).Specific).Value = item.PartiNumarasi;
                            if (item.KabulTarihi != "")
                            {
                                DateTime dtUretim = Convert.ToDateTime(item.KabulTarihi);
                                //oMatrixMikrobiyolojik.SetCellWithoutValidation(oMatrixMikrobiyolojik.RowCount, "Col_3", dtUretim.ToString("yyyyMMdd")); //Kabul Tarihi kolonu
                                ((SAPbouiCOM.EditText)oMatrixMikrobiyolojik.Columns.Item("Col_3").Cells.Item(oMatrixMikrobiyolojik.RowCount).Specific).Value = dtUretim.ToString("yyyyMMdd");
                            }
                            oMatrixMikrobiyolojik.AddRow();
                        }

                        oMatrixMikrobiyolojik.AutoResizeColumns();
                    }

                    if (frmAnaliz.Mode == BoFormMode.fm_OK_MODE)
                    {
                        frmAnaliz.Mode = BoFormMode.fm_UPDATE_MODE;
                    }
                }


            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Hata oluştu." + ex.Message);
            }
            finally
            {
                //frmAnaliz.Freeze(false);
            }
        }
    }

}
