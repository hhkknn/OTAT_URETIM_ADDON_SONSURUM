using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using AIF.UVT.SAPB1.HelperClass;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing.Design;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class UretimSiparisCogalt
    {
        [ItemAtt(AIFConn.UretimSiparisCogaltUID)]
        public SAPbouiCOM.Form frmUretimSiparisCogalt;

        [ItemAtt("Item_5")]
        public SAPbouiCOM.EditText EdtUretimSiparisNo;

        [ItemAtt("Item_6")]
        public SAPbouiCOM.EditText EdtUretimSiparisTarihi;

        [ItemAtt("Item_7")]
        public SAPbouiCOM.EditText EdtUretimSiparisMiktari;

        [ItemAtt("Item_8")]
        public SAPbouiCOM.EditText EdtUrunKodu;

        [ItemAtt("Item_9")]
        public SAPbouiCOM.EditText EdtUrunTanimi;

        [ItemAtt("Item_10")]
        public SAPbouiCOM.Matrix oMatrix;

        [ItemAtt("Item_11")]
        public SAPbouiCOM.Button btnSiparisCogalt;

        [ItemAtt("Item_12")]
        public SAPbouiCOM.Button btnIptal;

        string hataverildi = "";

        int siparisNo = 0;
        string siparisTarihi = "";
        double planlananMik = 0;
        string urunKodu = "";
        string urunTanimi = "";
        public void LoadForms(int _siparisNo, string _siparisTarihi, double _siparisMik, string _urunKodu, string _urunTanimi)
        {
            if (hataverildi == "")
            {
                SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery("select \"U_UrtPrtSekli\" as \"UrtPrtSekli\" from \"@AIF_UVT_PARAM\"");

                string UrtPrtSekli = oRS.Fields.Item("UrtPrtSekli").Value.ToString();

                if (UrtPrtSekli == "1")
                {
                    Handler.SAPApplication.MessageBox("Üretim siparişi Yönetme Şekli Üretim Siparişini Grupla Olduğunda Bu Özellik Kullanılamaz.");
                    hataverildi = "Hata";
                    return;
                }

                ConstVariables.oFnc.LoadSAPXML(AIFConn.UretimSiparisCogaltXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.UretimSiparisCogaltXML));
                Functions.CreateUserOrSystemFormComponent<UretimSiparisCogalt>(AIFConn.UrtSipCog);

                siparisNo = _siparisNo;
                siparisTarihi = _siparisTarihi;
                planlananMik = _siparisMik;
                urunKodu = _urunKodu;
                urunTanimi = _urunTanimi;

                InitForms();


            }
            else
            {
                hataverildi = "";
            }
        }
        public void InitForms()
        {
            try
            {
                frmUretimSiparisCogalt.Freeze(true);
                frmUretimSiparisCogalt.EnableMenu("1283", false);
                frmUretimSiparisCogalt.EnableMenu("1284", false);
                frmUretimSiparisCogalt.EnableMenu("1286", false);


                //ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                //ConstVariables.oRecordset.DoQuery("Select \"DocEntry\" from \"@AIF_UVT_ACTPARAM\"");

                //if (ConstVariables.oRecordset.RecordCount > 0)
                //{
                //    frmAktivteParametre.Mode = BoFormMode.fm_FIND_MODE;
                //    EdtDocEntry.Value = "1";
                //    btnAddOrUpdate.Item.Click();

                //}

                EdtUretimSiparisNo.Value = siparisNo.ToString();
                EdtUretimSiparisTarihi.Value = siparisTarihi.ToString();
                EdtUretimSiparisMiktari.Value = parseNumber_Seperator.setDoubleVal(planlananMik.ToString()).ToString();
                EdtUrunKodu.Value = urunKodu.ToString();
                EdtUrunTanimi.Value = urunTanimi.ToString();

                oMatrix.AddRow();
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific).Value = parseNumber_Seperator.setDoubleVal(planlananMik.ToString()).ToString();

                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Hata oluştu" + ex.Message);
            }
            finally
            {
                frmUretimSiparisCogalt.Freeze(false);

            }
        }

        string header = @"<?xml version=""1.0"" encoding=""UTF-16"" ?><dbDataSources uid=""@AIF_URT_PART1""><rows>{0}</rows></dbDataSources>";

        public static string operasyonplaniTarihi = "";

        string row = "<row><cells><cell><uid>U_PartiNo</uid><value>{0}</value></cell><cell><uid>U_Miktar</uid><value>{1}</value></cell><cell><uid>U_PartiKatSayi</uid><value>{2}</value></cell><cell><uid>U_PartiBasSaati</uid><value>{3}</value></cell><cell><uid>U_PartiBitSaati</uid><value>{4}</value></cell><cell><uid>U_Durum</uid><value>{5}</value></cell></cells></row>";
        bool satirdakiMiktarOtomatikDEgisiyor = false;
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
                    //    XmlDocument key = new XmlDocument();
                    //    key.LoadXml(BusinessObjectInfo.ObjectKey);
                    //    string docentry = key.SelectNodes("AIF_URT_PARTParams").Item(0).SelectNodes("DocEntry").Item(0).InnerXml;

                    //    SAPbobsCOM.ProductionOrders oWOR = (SAPbobsCOM.ProductionOrders)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);

                    //    oWOR.GetByKey(Convert.ToInt32(EdtUretimSiparisNo.Value));

                    //    //oWOR.PlannedQuantity = parseNumber_Seperator.setDoubleVal(EdtUretimSiparisMiktari.Value);

                    //    oWOR.UserFields.Fields.Item("U_PartiBelgeNo").Value = docentry;

                    //    ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //    ConstVariables.oRecordset.DoQuery("Select ISNULL(\"U_ISTASYON\",'') from \"OITT\" where \"Code\" = '" + oWOR.ItemNo + "'");

                    //    string istasyon = ConstVariables.oRecordset.Fields.Item(0).Value.ToString();


                    //    if (istasyon != "")
                    //    {
                    //        oWOR.UserFields.Fields.Item("U_ISTASYON").Value = istasyon;
                    //    }

                    //    var resp = oWOR.Update();
                    //}
                    break;
                case BoEventTypes.et_FORM_DATA_UPDATE:
                    break;
                case BoEventTypes.et_FORM_DATA_DELETE:
                    break;
                case BoEventTypes.et_FORM_DATA_LOAD:
                    //try
                    //{
                    //    if (!BusinessObjectInfo.BeforeAction)
                    //    {
                    //        for (int i = 1; i <= oMatrix.RowCount - 1; i++)
                    //        {
                    //            oMatrix.CommonSetting.SetCellEditable(i, 2, false);
                    //        }
                    //    }
                    //    else if (BusinessObjectInfo.BeforeAction)
                    //    {
                    //        oMatrix.Columns.Item(2).Editable = true;
                    //    }
                    //}
                    //catch (Exception)
                    //{
                    //}
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
        public string NoktalamaIsaretiDegistir(string val1)
        {
            string Val = Program.SAPnfi == "," ? val1.ToString().Replace(",", ".") : Program.SAPnfi == "." ? val1.ToString().Replace(".", ",") : val1.ToString();

            return Val;
        }
        bool col2LostFocus = false;
        bool sonrakiSatir = false;
        bool oncekiSatir = false;
        double siparismiktari = 0;
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
                    if (pVal.ColUID == "Col_2" && !pVal.BeforeAction)
                    {
                        if (col2LostFocus && sonrakiSatir)
                        {
                            oMatrix.SetCellFocus(oMatrix.RowCount, 1);
                            col2LostFocus = false;
                            sonrakiSatir = false;
                        }
                        else if (col2LostFocus && oncekiSatir)
                        {
                            oMatrix.SetCellFocus(oMatrix.RowCount, 1);
                            col2LostFocus = false;
                            oncekiSatir = false;
                        }
                    }
                    break;
                case BoEventTypes.et_LOST_FOCUS:
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    break;
                case BoEventTypes.et_CLICK:
                    if (!pVal.BeforeAction && pVal.ItemUID == "Item_11")
                    {
                        if (btnSiparisCogalt.Item.Enabled == false)
                        {
                            return false;
                        }
                        SAPbobsCOM.Recordset orec = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string hata = "";

                        string siparisYonetimSekli = "";
                        string sql = "Select \"U_UrtPrtSekli\" from \"@AIF_UVT_PARAM\"";
                        ConstVariables.oRecordset.DoQuery(sql);

                        if (ConstVariables.oRecordset.RecordCount > 0)
                        {
                            siparisYonetimSekli = ConstVariables.oRecordset.Fields.Item("U_UrtPrtSekli").Value.ToString();
                        }

                        if (siparisYonetimSekli == "1")
                        {
                            return false;
                        }

                        ProductionOrders wo = (ProductionOrders)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.oProductionOrders);
                        ProductionOrders wo_Add = (ProductionOrders)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.oProductionOrders); //yeni belge
                        wo.GetByKey(Convert.ToInt32(EdtUretimSiparisNo.Value));
                        ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);

                        ConstVariables.oRecordset.DoQuery("Select ISNULL(\"U_PartiMiktari\",0) as \"U_PartiMiktari\",\"U_ISTASYON\" from \"OITT\" where \"Code\" = '" + wo.ItemNo + "'");

                        string istasyon = ConstVariables.oRecordset.Fields.Item("U_ISTASYON").Value.ToString();
                        double UrunAgaciPartiMiktari = parseNumber_Seperator.ConvertToDouble(ConstVariables.oRecordset.Fields.Item("U_PartiMiktari").Value.ToString());

                        string aciklama = "";

                        int ilkBelge = 0;

                        //double yazilacakMiktar = 0;
                        for (int i = 1; i <= oMatrix.RowCount; i++)
                        {
                            string sipnosatiri = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific).Value.ToString();

                            if (sipnosatiri != "")
                            {
                                continue;
                            }

                            double OworPlananMiktar = parseNumber_Seperator.ConvertToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value.ToString());

                            if (wo.UserFields.Fields.Item("U_PartDahilEtme").Value.ToString() == "Evet") //partilendirme yapamz,durumu onaylandı yapar
                            {
                                aciklama = EdtUretimSiparisNo.Value + " numaralı " + parseNumber_Seperator.ConvertToDouble(EdtUretimSiparisMiktari.Value.ToString()) + " adet ürüne göre partilendirilmiştir.";
                                wo.Remarks = aciklama;

                                wo.ProductionOrderStatus = BoProductionOrderStatusEnum.boposReleased;

                                wo.UserFields.Fields.Item("U_GrupSipNo").Value = wo.DocumentNumber;
                                if (istasyon != "")
                                {
                                    wo_Add.UserFields.Fields.Item("U_ISTASYON").Value = istasyon;
                                }

                                wo.PlannedQuantity = OworPlananMiktar;
                                wo.UserFields.Fields.Item("U_OrjMiktar").Value = OworPlananMiktar.ToString();
                                int ret = wo.Update();

                            }
                            else
                            {
                                #region ÜRETİM SİPARİŞİ YENİ OLUŞACAK BELGE
                                int ret = -1;
                                if (ilkBelge > 0)
                                {
                                    wo_Add = (ProductionOrders)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.oProductionOrders);
                                    wo_Add.ItemNo = wo.ItemNo;
                                    wo_Add.PlannedQuantity = OworPlananMiktar;
                                    wo_Add.UserFields.Fields.Item("U_OrjMiktar").Value = OworPlananMiktar.ToString();
                                    wo_Add.Warehouse = wo.Warehouse;

                                    wo_Add.StartDate = wo.StartDate;
                                    wo_Add.DueDate = wo.DueDate;

                                    wo_Add.UserFields.Fields.Item("U_GrupSipNo").Value = wo.DocumentNumber;

                                    wo_Add.Remarks = aciklama;

                                    wo_Add.ProductionOrderStatus = BoProductionOrderStatusEnum.boposPlanned;

                                    if (istasyon != "")
                                    {
                                        wo_Add.UserFields.Fields.Item("U_ISTASYON").Value = istasyon;
                                    }

                                   
                                     
                                    

                                    //for (int i = 0; i <= wo.Lines.Count - 1; i++)
                                    //{
                                    //    //wo_Add.Lines.SetCurrentLine(i);
                                    //    wo.Lines.SetCurrentLine(i);

                                    //    wo_Add.Lines.ItemType = wo.Lines.ItemType;
                                    //    wo_Add.Lines.ItemNo = wo.Lines.ItemNo;
                                    //    wo_Add.Lines.BaseQuantity = wo.Lines.BaseQuantity;
                                    //    wo_Add.Warehouse = wo.Lines.Warehouse;

                                    //    //wo_Add.Lines.Add();
                                    //}

                                    ret = wo_Add.Add();

                                    if (ret != 0)
                                    {
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(i).Specific).Value = "Hata." + ConstVariables.oCompanyObject.GetLastErrorDescription();
                                        Handler.SAPApplication.MessageBox("Hata oluştu." + ConstVariables.oCompanyObject.GetLastErrorDescription().ToString());
                                        return false;
                                    }

                                    wo_Add.GetByKey(Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey()));

                                    #region grup parti no üretme 15.04.2022 

                                    if (Program.mKod == "010OTATURVT")
                                    {
                                        int belgeno = Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey());

                                        SAPbobsCOM.Recordset orec2 = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                        sql = "SELECT Count(T0.\"ItemCode\") as \"SipSayi\" FROM OWOR T0 WHERE T0.\"ItemCode\" = '" + wo.ItemNo + "' and T0.\"StartDate\" = '" + wo.StartDate.ToString("yyyyMMdd") + "' GROUP BY T0.\"ItemCode\" ";

                                        orec2.DoQuery(sql);

                                        if (orec2.RecordCount > 0)
                                        {
                                            int partisayi = 0;
                                            //partisayi = ConstVariables.oRecordset.RecordCount;
                                            partisayi = Convert.ToInt32(orec2.Fields.Item("SipSayi").Value);

                                            wo_Add.UserFields.Fields.Item("U_GrupPartiNo").Value = wo.StartDate.ToString("yyyMMdd") + "-" + belgeno.ToString() + "-" + Convert.ToString(Convert.ToInt32(partisayi));
                                        }

                                        if (orec2.RecordCount == 0)
                                        {
                                            wo_Add.UserFields.Fields.Item("U_GrupPartiNo").Value = wo.StartDate.ToString("yyyMMdd") + "-" + belgeno.ToString() + "-" + Convert.ToString(Convert.ToInt32(1));
                                        }
                                    }
                                    #endregion parti no üretme 15.04.2022

                                    wo_Add.ProductionOrderStatus = BoProductionOrderStatusEnum.boposReleased;

                                    ret = wo_Add.Update();


                                    //oMatrix.Columns.Item("Col_1").Editable = true;
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific).Value = ConstVariables.oCompanyObject.GetNewObjectKey();
                                    //oMatrix.Columns.Item("Col_1").Editable = false;

                                    if (ret == 0)
                                    {
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(i).Specific).Value = "Başarılı";
                                    }
                                    else
                                    {
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(i).Specific).Value = "Hata." + ConstVariables.oCompanyObject.GetLastErrorDescription();
                                    }
                                    #endregion ÜRETİM SİPARİŞİ YENİ OLUŞACAK BELGE
                                }
                                else
                                {
                                    #region ÜRETİM SİPARİŞİ MEVCUT BELGE GÜNCELLEME
                                    aciklama = EdtUretimSiparisNo.Value + " numaralı " + parseNumber_Seperator.ConvertToDouble(EdtUretimSiparisMiktari.Value.ToString()) + " adet ürüne göre partilendirilmiştir.";
                                    wo.Remarks = aciklama;

                                    wo.ProductionOrderStatus = BoProductionOrderStatusEnum.boposReleased;

                                    wo.UserFields.Fields.Item("U_GrupSipNo").Value = wo.DocumentNumber;
                                    if (istasyon != "")
                                    {
                                        wo.UserFields.Fields.Item("U_ISTASYON").Value = istasyon;
                                    }

                                    wo.PlannedQuantity = OworPlananMiktar;
                                    wo.UserFields.Fields.Item("U_OrjMiktar").Value = OworPlananMiktar.ToString();
                                    #region grup parti no üretme 15.04.2022
                                    //if (Program.mKod == "010OTATURVT")
                                    //{

                                    //    string val = wo.UserFields.Fields.Item("U_GrupPartiNo").Value.ToString();

                                    //    var satir = val.Split('-');
                                    //    val = satir[satir.Length - 1];

                                    //    if (val == "")
                                    //    {
                                    //        val = "1";

                                    //        wo.UserFields.Fields.Item("U_GrupPartiNo").Value = wo.StartDate.ToString("yyyyMMdd") + "-" + wo.DocumentNumber + "-" + Convert.ToString(Convert.ToInt32(val));
                                    //    }
                                    //}
                                    #endregion grup parti no üretme 15.04.2022

                                    #region grup parti no üretme 31.05.2022 

                                    if (Program.mKod == "010OTATURVT")
                                    {
                                        int belgeno = Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey());

                                        SAPbobsCOM.Recordset orec2 = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                        sql = "SELECT Count(T0.\"ItemCode\") as \"SipSayi\" FROM OWOR T0 WHERE T0.\"ItemCode\" = '" + wo.ItemNo + "' and T0.\"StartDate\" = '" + wo.StartDate.ToString("yyyyMMdd") + "' GROUP BY T0.\"ItemCode\" ";

                                        orec2.DoQuery(sql);

                                        if (orec2.RecordCount > 0)
                                        {
                                            int partisayi = 0;
                                            //partisayi = ConstVariables.oRecordset.RecordCount;
                                            partisayi = Convert.ToInt32(orec2.Fields.Item("SipSayi").Value);

                                            wo.UserFields.Fields.Item("U_GrupPartiNo").Value = wo.StartDate.ToString("yyyMMdd") + "-" + belgeno.ToString() + "-" + Convert.ToString(Convert.ToInt32(partisayi));
                                        }

                                        if (orec2.RecordCount == 0)
                                        {
                                            wo.UserFields.Fields.Item("U_GrupPartiNo").Value = wo.StartDate.ToString("yyyMMdd") + "-" + belgeno.ToString() + "-" + Convert.ToString(Convert.ToInt32(1));
                                        }
                                    }
                                    #endregion parti no üretme 31.05.2022

                                    ret = wo.Update();


                                    if (ret == 0)
                                    {
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(i).Specific).Value = "Başarılı";
                                    }
                                    else
                                    {
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(i).Specific).Value = "Hata." + ConstVariables.oCompanyObject.GetLastErrorDescription();
                                        Handler.SAPApplication.MessageBox("Mevcut belge güncellenemediğinden işleme devam edilmemiştir.");
                                        return false;
                                    }


                                    //oMatrix.Columns.Item("Col_1").Editable = true;
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(i).Specific).Value = ConstVariables.oCompanyObject.GetNewObjectKey();
                                    //oMatrix.Columns.Item("Col_1").Cells.Item(i).Click();
                                    //oMatrix.Columns.Item("Col_1").Editable = false;


                                    ilkBelge++;
                                    #endregion ÜRETİM SİPARİŞİ MEVCUT BELGE GÜNCELLEME
                                }
                            }
                        }

                        if (hata == "")
                        {
                            Handler.SAPApplication.MessageBox("Sipariş çoğaltma işlemleri tamamlandı.");
                            btnSiparisCogalt.Item.Enabled = false;
                        }
                        else
                        {
                            string msg = "Aşağıdaki siparişler için hatalar oluştu.";
                            msg += hata;
                            Handler.SAPApplication.MessageBox(msg);
                        }


                    }
                    else if (pVal.ItemUID == "Item_12" && !pVal.BeforeAction)
                    {
                        try
                        {
                            frmUretimSiparisCogalt.Close();
                        }
                        catch (Exception)
                        {

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
                    if (pVal.ColUID == "Col_0" && pVal.ItemChanged && !pVal.BeforeAction)
                    {
                        try
                        {
                            //frmUretimSiparisCogalt.Freeze(true);
                            string a = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value.ToString();

                            double deger = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value.ToString() == "" ? 0 : parseNumber_Seperator.ConvertToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value.ToString());
                            if (deger > 0)
                            {
                                string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                            select new
                                            {
                                                Miktar = Convert.ToDouble((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value),
                                            }).ToList();


                                double siparisMiktari = parseNumber_Seperator.ConvertToDouble(EdtUretimSiparisMiktari.Value.ToString());
                                double matrixtoplammiktar = parseNumber_Seperator.ConvertToDouble(rows.Sum(x => x.Miktar).ToString());
                                double sonuc = siparisMiktari - matrixtoplammiktar;

                                if (matrixtoplammiktar > siparisMiktari)
                                {
                                    double sonsatirmik = parseNumber_Seperator.ConvertToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific).Value.ToString());
                                    double girilebilirmik = siparisMiktari - (matrixtoplammiktar - sonsatirmik);
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific).Value = parseNumber_Seperator.setDoubleVal(girilebilirmik.ToString()).ToString();
                                    Handler.SAPApplication.MessageBox("Sipariş miktarından fazla miktar girilemez.");
                                    return false;
                                }
                                if (sonuc > 0)
                                {
                                    oMatrix.AddRow();
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific).Value = parseNumber_Seperator.setDoubleVal(sonuc.ToString()).ToString();

                                    //EdtUrunKodu.Item.Click();
                                    for (int i = 1; i <= oMatrix.RowCount - 1; i++)
                                    {
                                        oMatrix.CommonSetting.SetCellEditable(i, 1, false);
                                    }

                                    col2LostFocus = true;
                                    sonrakiSatir = true;

                                }
                                else if (sonuc < 0)
                                {
                                    Handler.SAPApplication.MessageBox(string.Format("Kalan Miktar {0}'dır kalan kadar değer otomatik atanmıştır.", deger - (sonuc * -1)));
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value = parseNumber_Seperator.setDoubleVal((deger - (sonuc * -1)).ToString()).ToString();
                                }
                                //else if (matrixtoplammiktar == siparisMiktari)
                                //{
                                //    //oMatrix.AddRow();
                                //    //((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific).Value = parseNumber_Seperator.setDoubleVal(sonuc.ToString()).ToString();

                                //    //EdtUrunKodu.Item.Click();
                                //    for (int i = 1; i <= oMatrix.RowCount - 1; i++)
                                //    {
                                //        oMatrix.CommonSetting.SetCellEditable(i, 1, true);
                                //    }

                                //    col2LostFocus = true;
                                //    sonrakiSatir = true;

                                //    if (oMatrix.RowCount > 1)
                                //    {
                                //        //oMatrix.AddRow();
                                //        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value = parseNumber_Seperator.setDoubleVal((deger - (sonuc * -1)).ToString()).ToString();
                                //    }
                                //}
                            }
                            else if (deger == 0)
                            {
                                if (pVal.Row > 1)
                                {
                                    #region sıfır girildiğinde satır silme
                                    oMatrix.DeleteRow(pVal.Row);
                                    oMatrix.CommonSetting.SetCellEditable(pVal.Row - 1, 1, true);
                                    #endregion

                                    string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                                    var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                                select new
                                                {
                                                    Miktar = Convert.ToDouble((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value),
                                                }).ToList();

                                    double sonsatir = parseNumber_Seperator.ConvertToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific).Value.ToString());
                                    double siparisMiktari = parseNumber_Seperator.ConvertToDouble(EdtUretimSiparisMiktari.Value.ToString());
                                    double matrixtoplammiktar = parseNumber_Seperator.ConvertToDouble(rows.Sum(x => x.Miktar).ToString());
                                    double sonuc = siparisMiktari - (matrixtoplammiktar - sonsatir);

                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific).Value = parseNumber_Seperator.setDoubleVal(sonuc.ToString()).ToString();

                                    //Handler.SAPApplication.MessageBox(string.Format("Kalan Miktar 0 veya boş girilemez."));

                                }
                                else if (pVal.Row == 1)
                                {
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific).Value = parseNumber_Seperator.setDoubleVal(EdtUretimSiparisMiktari.Value.ToString()).ToString();
                                    //oMatrix.AutoResizeColumns();
                                }

                                col2LostFocus = true;
                                oncekiSatir = true;
                            }
                        }
                        catch (Exception ex)
                        {
                            Handler.SAPApplication.StatusBar.SetText("Hesaplama yapılırken hata oluştu. Hata Kodu A0018", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        }
                        finally
                        {
                            //frmUretimSiparisCogalt.Freeze(false);
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
