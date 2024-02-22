using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.UVT.SAPB1.HelperClass;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class SiparisOnaylama
    {
        [ItemAtt(AIFConn.SiparisOnaylamaUID)]
        public SAPbouiCOM.Form frmSiparisOnaylama;

        [ItemAtt("Item_5")]
        public SAPbouiCOM.Matrix oMatrix;
        [ItemAtt("Item_1")]
        public SAPbouiCOM.EditText oEdtBaslangic;
        [ItemAtt("Item_3")]
        public SAPbouiCOM.EditText oEdtBitis;


        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.SiparisOnaylamaXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.SiparisOnaylamaXML));
            Functions.CreateUserOrSystemFormComponent<SiparisOnaylama>(AIFConn.SipOnay);

            InitForms();
        }
        public void InitForms()
        {
            try
            {
                oDataTable = frmSiparisOnaylama.DataSources.DataTables.Add("DATA");
                oEdtBaslangic.Value = DateTime.Now.AddDays(-7).ToString("yyyyMMdd");
                oEdtBitis.Value = DateTime.Now.ToString("yyyyMMdd");
                Listele(oEdtBaslangic.Value, oEdtBitis.Value);
            }
            catch (Exception ex)
            {
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
                    if (!pVal.BeforeAction && pVal.ItemUID == "Item_6")
                    {
                        string xml = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                        var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                    select new
                                    {
                                        Secili = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "#" select new XElement(y.Element("Value"))).First().Value,
                                        SiparisNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value,
                                        SiparisTarihi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value,
                                        UrunKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value,
                                        PlanlananMiktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value),
                                    }).ToList();

                        SAPbobsCOM.Recordset orec = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string docentry = "";
                        var oCompanyServiceData = ConstVariables.oCompanyObject.GetCompanyService();
                        var oGeneralServiceData = oCompanyServiceData.GetGeneralService("AIF_URT_PART");
                        SAPbobsCOM.GeneralDataCollection oChildren = null;
                        SAPbobsCOM.GeneralData oChild = null;
                        var oGeneralDataData = ((SAPbobsCOM.GeneralData)(oGeneralServiceData.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                        var oGeneralDataParams = ((SAPbobsCOM.GeneralDataParams)(oGeneralServiceData.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)));
                        string hata = "";

                        string siparisYonetimSekli = "";
                        string sql = "Select \"U_UrtPrtSekli\" from \"@AIF_UVT_PARAM\"";
                        ConstVariables.oRecordset.DoQuery(sql);

                        if (ConstVariables.oRecordset.RecordCount > 0)
                        {
                            siparisYonetimSekli = ConstVariables.oRecordset.Fields.Item("U_UrtPrtSekli").Value.ToString();
                        }

                        foreach (var item in rows.Where(x => x.Secili == "Y"))
                        {

                            if (siparisYonetimSekli == "1")
                            {
                                //orec.DoQuery("Select MAX(\"DocEntry\")+1 from \"@AIF_URT_PART\"");

                                //docentry = orec.Fields.Item(0).Value.ToString();

                                //if (docentry == "")
                                //{
                                //    docentry = "1";
                                //}


                                ProductionOrders wo = (ProductionOrders)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.oProductionOrders);
                                wo.GetByKey(Convert.ToInt32(item.SiparisNo));

                                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
                                ConstVariables.oRecordset.DoQuery("Select ISNULL(\"U_PartiMiktari\",0),\"U_ISTASYON\" from \"OITT\" where \"Code\" = '" + wo.ItemNo + "'");

                                string istasyon = ConstVariables.oRecordset.Fields.Item(1).Value.ToString();
                                double UrunAgaciPartiMiktari = Convert.ToDouble(ConstVariables.oRecordset.Fields.Item(0).Value);
                                double OworPlananMiktar = item.PlanlananMiktar;
                                double yazilacakMiktar = 0;
                                double yazilanToplamMiktar = 0;
                                DateTime dttarih = new DateTime(Convert.ToInt32(item.SiparisTarihi.Substring(0, 4)), Convert.ToInt32(item.SiparisTarihi.Substring(4, 2)), Convert.ToInt32(item.SiparisTarihi.Substring(6, 2)));

                                //oGeneralDataData.SetProperty("DocNum", Convert.ToInt32(docentry));
                                oGeneralDataData.SetProperty("U_UretimSipNo", item.SiparisNo);
                                oGeneralDataData.SetProperty("U_UretimSipTar", dttarih);
                                oGeneralDataData.SetProperty("U_UretimSipMik", item.PlanlananMiktar);
                                oGeneralDataData.SetProperty("U_UrunKodu", item.UrunKodu);
                                oGeneralDataData.SetProperty("U_UrunTanimi", wo.ProductDescription);

                                oChildren = oGeneralDataData.Child("AIF_URT_PART1");


                                int sira = 1;

                                if (UrunAgaciPartiMiktari > 0)
                                {
                                tekrarla:
                                    if (OworPlananMiktar < UrunAgaciPartiMiktari)
                                    {
                                        yazilacakMiktar = OworPlananMiktar;
                                        OworPlananMiktar = OworPlananMiktar - yazilacakMiktar;
                                    }
                                    else
                                    {
                                        yazilacakMiktar = UrunAgaciPartiMiktari;
                                        OworPlananMiktar = Math.Round(OworPlananMiktar, 2) - yazilacakMiktar;
                                    }

                                    oChild = oChildren.Add();

                                    #region eski parti no üretme
                                    //oChild.SetProperty("U_PartiNo", DateTime.Now.ToString("yyyyMMdd") + "-" + item.SiparisNo + "-" + sira.ToString());
                                    #endregion

                                    #region startdate ile parti no üretme 18 şubat 2022
                                    oChild.SetProperty("U_PartiNo", dttarih.ToString("yyyyMMdd") + "-" + item.SiparisNo + "-" + sira.ToString()); //başlangıç tarihi ile parti üretme 
                                    #endregion

                                    oChild.SetProperty("U_Miktar", yazilacakMiktar);
                                    oChild.SetProperty("U_PartiKatSayi", Convert.ToDouble(item.PlanlananMiktar / yazilacakMiktar));


                                    if (OworPlananMiktar > 0)
                                    {
                                        yazilanToplamMiktar = yazilanToplamMiktar + yazilacakMiktar;
                                        sira++;
                                        goto tekrarla;

                                    }
                                }
                                else
                                {

                                    oChild = oChildren.Add();

                                    #region eski parti no üretme
                                    //oChild.SetProperty("U_PartiNo", DateTime.Now.ToString("yyyyMMdd") + "-" + item.SiparisNo + "-" + sira.ToString());
                                    #endregion

                                    #region startdate ile parti no üretme 18 şubat 2022
                                    oChild.SetProperty("U_PartiNo", dttarih.ToString("yyyyMMdd") + "-" + item.SiparisNo + "-" + sira.ToString()); //başlangıç tarihi ile parti üretme 
                                    #endregion

                                    oChild.SetProperty("U_Miktar", OworPlananMiktar);
                                    oChild.SetProperty("U_PartiKatSayi", Convert.ToDouble("1"));
                                }

                                SAPbobsCOM.GeneralDataParams result = oGeneralServiceData.Add(oGeneralDataData);

                                oGeneralDataData = ((SAPbobsCOM.GeneralData)(oGeneralServiceData.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                                oChildren = null;
                                oChild = null;

                                orec.DoQuery("Select MAX(\"DocEntry\") from \"@AIF_URT_PART\" where \"U_UretimSipNo\" = '" + item.SiparisNo + "'");

                                docentry = orec.Fields.Item(0).Value.ToString();


                                if (docentry != "")
                                {
                                    wo.ProductionOrderStatus = BoProductionOrderStatusEnum.boposReleased;
                                    wo.UserFields.Fields.Item("U_PartiBelgeNo").Value = docentry;
                                    if (istasyon != "")
                                    {
                                        wo.UserFields.Fields.Item("U_ISTASYON").Value = istasyon;
                                    }
                                    var ret = wo.Update();

                                    if (ret != 0)
                                    {
                                        oGeneralDataParams.SetProperty("DocEntry", docentry);
                                        oGeneralServiceData.Delete(oGeneralDataParams);
                                        hata += Environment.NewLine;
                                        hata += item.SiparisNo + " siparişi için onaylama sırasında hata alındı." + ConstVariables.oCompanyObject.GetLastErrorDescription();
                                    }

                                }
                            }
                            else if (siparisYonetimSekli == "2")
                            {
                                ProductionOrders wo = (ProductionOrders)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.oProductionOrders);
                                ProductionOrders wo_Add = (ProductionOrders)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.oProductionOrders);
                                wo.GetByKey(Convert.ToInt32(item.SiparisNo));


                                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);

                                ConstVariables.oRecordset.DoQuery("Select ISNULL(\"U_PartiMiktari\",0) as \"U_PartiMiktari\",\"U_ISTASYON\" from \"OITT\" where \"Code\" = '" + wo.ItemNo + "'");

                                string istasyon = ConstVariables.oRecordset.Fields.Item("U_ISTASYON").Value.ToString();

                                string aciklama = "";

                                int ilkBelge = 0;
                                double UrunAgaciPartiMiktari = Convert.ToDouble(ConstVariables.oRecordset.Fields.Item("U_PartiMiktari").Value);
                                double OworPlananMiktar = item.PlanlananMiktar;
                                double yazilacakMiktar = 0;

                                if (wo.UserFields.Fields.Item("U_PartDahilEtme").Value.ToString() == "Evet")
                                {
                                    aciklama = item.SiparisNo + " numaralı " + OworPlananMiktar.ToString() + " adet ürüne göre partilendirilmiştir.";
                                    wo.Remarks = aciklama;

                                    wo.ProductionOrderStatus = BoProductionOrderStatusEnum.boposReleased;

                                    wo.UserFields.Fields.Item("U_GrupSipNo").Value = wo.DocumentNumber;
                                    if (istasyon != "")
                                    {
                                        wo_Add.UserFields.Fields.Item("U_ISTASYON").Value = istasyon;
                                    }

                                    wo.PlannedQuantity = OworPlananMiktar;
                                    wo.UserFields.Fields.Item("U_OrjMiktar").Value = OworPlananMiktar;
                                    int ret = wo.Update();

                                }
                                else
                                {

                                    if (UrunAgaciPartiMiktari > 0)
                                    {
                                        if (OworPlananMiktar > UrunAgaciPartiMiktari)
                                        {
                                            yazilacakMiktar = UrunAgaciPartiMiktari;
                                        }
                                        else
                                        {
                                            yazilacakMiktar = OworPlananMiktar;
                                        }

                                        while (OworPlananMiktar > 0)
                                        {
                                            int ret = -1;
                                            if (ilkBelge > 0)
                                            {
                                                wo_Add = (ProductionOrders)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.oProductionOrders);
                                                wo_Add.ItemNo = wo.ItemNo;
                                                wo_Add.PlannedQuantity = yazilacakMiktar;
                                                wo_Add.UserFields.Fields.Item("U_OrjMiktar").Value = yazilacakMiktar.ToString();
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

                                                wo_Add.GetByKey(Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey()));

                                                wo_Add.ProductionOrderStatus = BoProductionOrderStatusEnum.boposReleased;

                                                #region grup parti no üretme 20.04.2022 
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
                                                #endregion grup parti no üretme 15.04.2022

                                                ret = wo_Add.Update();
                                            }
                                            else
                                            {
                                                aciklama = item.SiparisNo + " numaralı " + yazilacakMiktar.ToString() + " adet ürüne göre partilendirilmiştir.";
                                                wo.Remarks = aciklama;

                                                wo.ProductionOrderStatus = BoProductionOrderStatusEnum.boposReleased;

                                                wo.UserFields.Fields.Item("U_GrupSipNo").Value = wo.DocumentNumber;
                                                if (istasyon != "")
                                                {
                                                    wo.UserFields.Fields.Item("U_ISTASYON").Value = istasyon;
                                                }

                                                wo.PlannedQuantity = yazilacakMiktar;
                                                wo.UserFields.Fields.Item("U_OrjMiktar").Value = yazilacakMiktar.ToString();
                                                #region grup parti no üretme 20.04.2022 
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
                                                #endregion grup parti no üretme 15.04.2022

                                                ret = wo.Update();

                                                ilkBelge++;

                                            }

                                            #region Partilendirme olmayacak çünkü her belge 1 adet parti
                                            //var adsada = ConstVariables.oCompanyObject.GetLastErrorDescription();


                                            //Handler.SAPApplication.ActivateMenuItem("4369");

                                            //var oform = Handler.SAPApplication.Forms.ActiveForm;

                                            //oform.Mode = BoFormMode.fm_FIND_MODE;

                                            //((SAPbouiCOM.EditText)oform.Items.Item("18").Specific).Value = "7419";

                                            //oform.Items.Item("1").Click();

                                            //try
                                            //{
                                            //    try
                                            //    {
                                            //        oform.EnableMenu("1287", true);
                                            //    }
                                            //    catch (Exception)
                                            //    {

                                            //    }

                                            //    Handler.SAPApplication.ActivateMenuItem("1284");

                                            //    ((SAPbouiCOM.EditText)oform.Items.Item("12").Specific).Value = yazilacakMiktar.ToString();

                                            //    oform.Items.Item("1").Click();
                                            //}
                                            //catch (Exception)
                                            //{

                                            //} 


                                            //DateTime dttarih = new DateTime(Convert.ToInt32(item.SiparisTarihi.Substring(0, 4)), Convert.ToInt32(item.SiparisTarihi.Substring(4, 2)), Convert.ToInt32(item.SiparisTarihi.Substring(6, 2)));
                                            //if (ret == 0)
                                            //{
                                            //    int yeniOlusanUretimSipNo = Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey());
                                            //    oGeneralDataData.SetProperty("U_UretimSipNo", yeniOlusanUretimSipNo.ToString());
                                            //    oGeneralDataData.SetProperty("U_UretimSipTar", dttarih);
                                            //    oGeneralDataData.SetProperty("U_UretimSipMik", yazilacakMiktar);
                                            //    oGeneralDataData.SetProperty("U_UrunKodu", wo.ItemNo);
                                            //    oGeneralDataData.SetProperty("U_UrunTanimi", wo.ProductDescription);

                                            //    oChildren = oGeneralDataData.Child("AIF_URT_PART1");

                                            //    oChild = oChildren.Add();

                                            //    oChild.SetProperty("U_PartiNo", DateTime.Now.ToString("yyyyMMdd") + "-" + yeniOlusanUretimSipNo + "-" + "1");


                                            //    #region startdate ile parti no üretme şimdilik canlıya alınmayacak. 7 ocak 2022
                                            //    //oChild.SetProperty("U_PartiNo", dttarih.ToString("yyyyMMdd") + "-" + item.SiparisNo + "-" + sira.ToString()); //başlangıç tarihi ile parti üretme 
                                            //    #endregion

                                            //    oChild.SetProperty("U_Miktar", yazilacakMiktar);
                                            //    oChild.SetProperty("U_PartiKatSayi", Convert.ToDouble("1"));



                                            //    SAPbobsCOM.GeneralDataParams result = oGeneralServiceData.Add(oGeneralDataData);

                                            //    oGeneralDataData = ((SAPbobsCOM.GeneralData)(oGeneralServiceData.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                                            //    oChildren = null;
                                            //    oChild = null;

                                            //    orec.DoQuery("Select MAX(\"DocEntry\") from \"@AIF_URT_PART\" where \"U_UretimSipNo\" = '" + yeniOlusanUretimSipNo + "'");

                                            //    docentry = orec.Fields.Item(0).Value.ToString();


                                            //    if (docentry != "")
                                            //    {
                                            //        wo_Add.GetByKey(yeniOlusanUretimSipNo);

                                            //        wo_Add.ProductionOrderStatus = BoProductionOrderStatusEnum.boposReleased;
                                            //        wo_Add.UserFields.Fields.Item("U_PartiBelgeNo").Value = docentry;
                                            //        if (istasyon != "")
                                            //        {
                                            //            wo_Add.UserFields.Fields.Item("U_ISTASYON").Value = istasyon;
                                            //        }
                                            //        var ret2 = wo_Add.Update();

                                            //        if (ret2 != 0)
                                            //        {
                                            //            oGeneralDataParams.SetProperty("DocEntry", docentry);
                                            //            oGeneralServiceData.Delete(oGeneralDataParams);
                                            //            hata += Environment.NewLine;
                                            //            hata += docentry + " siparişi için onaylama sırasında hata alındı." + ConstVariables.oCompanyObject.GetLastErrorDescription();
                                            //        }

                                            //    }
                                            //} 
                                            #endregion

                                            OworPlananMiktar = OworPlananMiktar - UrunAgaciPartiMiktari;

                                            if (OworPlananMiktar > UrunAgaciPartiMiktari)
                                            {
                                                yazilacakMiktar = UrunAgaciPartiMiktari;
                                            }
                                            else
                                            {
                                                yazilacakMiktar = OworPlananMiktar;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        aciklama = item.SiparisNo + " numaralı " + OworPlananMiktar.ToString() + " adet ürüne göre partilendirilmiştir.";
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
                                        #endregion grup parti no üretme 15.04.2022

                                        int ret = wo.Update();

                                    }
                                }
                            }
                        }

                        Listele(oEdtBaslangic.Value, oEdtBitis.Value);
                        if (hata == "")
                        {
                            Handler.SAPApplication.MessageBox("Onaylama işlemleri tamamlandı.");
                        }
                        else
                        {
                            string msg = "Aşağıdaki siparişler için hatalar oluştu.";
                            msg += hata;
                            Handler.SAPApplication.MessageBox(msg);
                        }

                    }
                    else if (pVal.ItemUID == "Item_4" && !pVal.BeforeAction)
                    {
                        Listele(oEdtBaslangic.Value, oEdtBitis.Value);
                    }
                    break;
                case BoEventTypes.et_DOUBLE_CLICK:
                    if (!pVal.BeforeAction)
                    {
                        try
                        {
                            return false;
                            if (pVal.Row > 0)
                            {
                                var uretimSiparisNo = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific).Value.ToString();

                                if (uretimSiparisNo != "")
                                {
                                    AIFConn.UrtSipPart.LoadForms();
                                    SAPbouiCOM.Form activeForm = (SAPbouiCOM.Form)Handler.SAPApplication.Forms.ActiveForm;

                                    activeForm.Mode = BoFormMode.fm_FIND_MODE;
                                    ((SAPbouiCOM.EditText)activeForm.Items.Item("Item_8").Specific).Value = uretimSiparisNo;

                                    activeForm.Items.Item("1").Click();
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    break;
                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                    if (pVal.BeforeAction && pVal.ColUID == "Col_9")
                    {
                        var uretimSiparisNo = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_9").Cells.Item(pVal.Row).Specific).Value.ToString();

                        if (uretimSiparisNo != "")
                        {
                            AIFConn.UrtSipPart.LoadForms();
                            SAPbouiCOM.Form activeForm = (SAPbouiCOM.Form)Handler.SAPApplication.Forms.ActiveForm;

                            activeForm.Mode = BoFormMode.fm_FIND_MODE;
                            ((SAPbouiCOM.EditText)activeForm.Items.Item("Item_8").Specific).Value = uretimSiparisNo;

                            activeForm.Items.Item("1").Click();
                        }
                        BubbleEvent = false;
                    }
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
        SAPbouiCOM.DataTable oDataTable = null;
        private void Listele(string baslangicTarihi, string bitisTarihi)
        {
            try
            {
                frmSiparisOnaylama.Freeze(true);
                oDataTable.Clear();

                //oDataTable.LoadSerializedXML(BoDataTableXmlSelect.dxs_All, "");


                string sql = "";

                if (Program.mKod == "010OTATURVT")
                {
                    sql = "Select * from (Select \"DocNum\",\"StartDate\",\"ProdName\",T10.\"ItemCode\",\"PlannedQty\",CONVERT(varchar, \"StartDate\", 112) as \"StartDate1\",(Select Descr from (select TOP 1 FieldID from CUFD as T0 where T0.TableID = 'OITM' and T0.AliasID = 'ItemGroup2') as tbl INNER JOIN UFD1 as T1 ON tbl.FieldID = T1.FieldID and TableID = 'OITM' and FldValue = T10.\"U_ItemGroup2\") as \"Ürün Grubu\",T0.\"PlannedQty\" as \"Planlanan Miktar\",(Select ISNULL(Count(URT2.\"DocEntry\"),0) from \"@AIF_URT_PART\" as URT1 INNER JOIN \"@AIF_URT_PART1\" as URT2 ON URT1.\"DocEntry\" = URT2.\"DocEntry\" where URT1.\"U_UretimSipNo\" = T0.\"DocNum\") \"Üretim Parti Sayısı\",(Select TOP 1 ISNULL(URT2.\"DocEntry\",0) from \"@AIF_URT_PART\" as URT1 INNER JOIN \"@AIF_URT_PART1\" as URT2 ON URT1.\"DocEntry\" = URT2.\"DocEntry\" where URT1.\"U_UretimSipNo\" = T0.\"DocNum\") as \"Üretim Parti Numarası\",(select TUFD1.Descr from CUFD as TCUFD INNER JOIN UFD1 as TUFD1 ON TCUFD.FieldID = TUFD1.FieldID where TCUFD.TableID = 'OWOR' and TUFD1.TableID = 'OWOR' and TCUFD.AliasID='ISTASYON' and TUFD1.FldValue = T0.\"U_Istasyon\") as \"İstasyon\" from \"OWOR\" as T0 INNER JOIN OITM as T10 ON T0.\"ItemCode\" = T10.\"ItemCode\" where T0.\"StartDate\" >= '" + baslangicTarihi + "' and T0.\"StartDate\" <= '" + bitisTarihi + "' and \"Status\" = 'P') as tbl order by tbl.\"İstasyon\"";
                }
                else if (Program.mKod == "20URVT")
                {
                    sql = "Select * from (Select \"DocNum\",\"StartDate\",\"ProdName\",T10.\"ItemCode\",\"PlannedQty\",CONVERT(varchar, \"StartDate\", 112) as \"StartDate1\", T11.\"ItmsGrpNam\" as \"Ürün Grubu\",T0.\"PlannedQty\" as \"Planlanan Miktar\",(Select ISNULL(Count(URT2.\"DocEntry\"),0) from \"@AIF_URT_PART\" as URT1 INNER JOIN \"@AIF_URT_PART1\" as URT2 ON URT1.\"DocEntry\" = URT2.\"DocEntry\" where URT1.\"U_UretimSipNo\" = T0.\"DocNum\") \"Üretim Parti Sayısı\",(Select TOP 1 ISNULL(URT2.\"DocEntry\",0) from \"@AIF_URT_PART\" as URT1 INNER JOIN \"@AIF_URT_PART1\" as URT2 ON URT1.\"DocEntry\" = URT2.\"DocEntry\" where URT1.\"U_UretimSipNo\" = T0.\"DocNum\") as \"Üretim Parti Numarası\",(select TUFD1.Descr from CUFD as TCUFD INNER JOIN UFD1 as TUFD1 ON TCUFD.FieldID = TUFD1.FieldID where TCUFD.TableID = 'OWOR' and TUFD1.TableID = 'OWOR' and TCUFD.AliasID='ISTASYON' and TUFD1.FldValue = T0.\"U_Istasyon\") as \"İstasyon\" from \"OWOR\" as T0 INNER JOIN OITM as T10 ON T0.\"ItemCode\" = T10.\"ItemCode\" LEFT JOIN OITB as T11 ON T10.\"ItmsGrpCod\" = T11.\"ItmsGrpCod\" where T0.\"StartDate\" >= '" + baslangicTarihi + "' and T0.\"StartDate\" <= '" + bitisTarihi + "' and \"Status\" = 'P') as tbl order by tbl.\"İstasyon\"";
                }

                oDataTable.ExecuteQuery(sql);


                oMatrix.Clear();

                oMatrix.Columns.Item("Col_10").DataBind.Bind("DATA", "İstasyon");
                oMatrix.Columns.Item("Col_0").DataBind.Bind("DATA", "DocNum");
                oMatrix.Columns.Item("Col_1").DataBind.Bind("DATA", "StartDate");
                oMatrix.Columns.Item("Col_2").DataBind.Bind("DATA", "ProdName");
                oMatrix.Columns.Item("Col_3").DataBind.Bind("DATA", "ItemCode");
                oMatrix.Columns.Item("Col_4").DataBind.Bind("DATA", "PlannedQty");
                oMatrix.Columns.Item("Col_5").DataBind.Bind("DATA", "StartDate1");
                oMatrix.Columns.Item("Col_6").DataBind.Bind("DATA", "Ürün Grubu");
                oMatrix.Columns.Item("Col_7").DataBind.Bind("DATA", "Planlanan Miktar");
                oMatrix.Columns.Item("Col_8").DataBind.Bind("DATA", "Üretim Parti Sayısı");
                oMatrix.Columns.Item("Col_9").DataBind.Bind("DATA", "Üretim Parti Numarası");

                oMatrix.LoadFromDataSource();

                oMatrix.AutoResizeColumns();

            }
            catch (Exception)
            {
            }
            finally
            {
                frmSiparisOnaylama.Freeze(false);
            }
        }
    }
}
