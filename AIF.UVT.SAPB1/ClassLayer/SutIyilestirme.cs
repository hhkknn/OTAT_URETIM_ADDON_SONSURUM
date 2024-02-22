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
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class SutIyilestirme
    {
        [ItemAtt(AIFConn.SutIyilestirmeUID)]
        public SAPbouiCOM.Form frmSutIyilestirme;

        [ItemAtt("Item_1")]
        public SAPbouiCOM.ComboBox oComboAy;
        [ItemAtt("Item_3")]
        public SAPbouiCOM.ComboBox oComboYil;
        [ItemAtt("Item_4")]
        public SAPbouiCOM.Matrix oMatrix;
        [ItemAtt("Item_9")]
        public SAPbouiCOM.EditText oEditDocEntry;
        [ItemAtt("1")]
        public SAPbouiCOM.Button btnaddOrUpdate;

        private string header = @"<?xml version=""1.0"" encoding=""UTF-16"" ?><dbDataSources uid=""@AIF_SUTIYLSTR1""><rows>{0}</rows></dbDataSources>";

        private string row = "<row><cells><cell><uid>U_TedarikciAdi</uid><value>{0}</value></cell><cell><uid>U_TedarikciKodu</uid><value>{1}</value></cell><cell><uid>U_Nitelik</uid><value>{2}</value></cell><cell><uid>U_Sicaklik</uid><value>{3}</value></cell><cell><uid>U_Brix</uid><value>{4}</value></cell><cell><uid>U_Yag</uid><value>{5}</value></cell><cell><uid>U_YogurtPL</uid><value>{6}</value></cell><cell><uid>U_PH</uid><value>{7}</value></cell><cell><uid>U_SH</uid><value>{8}</value></cell><cell><uid>U_KaynatSH</uid><value>{9}</value></cell><cell><uid>U_Protein</uid><value>{10}</value></cell><cell><uid>U_Laktoz</uid><value>{11}</value></cell><cell><uid>U_KatilanSu</uid><value>{12}</value></cell><cell><uid>U_Somatik</uid><value>{13}</value></cell><cell><uid>U_MetilenMavisi</uid><value>{14}</value></cell><cell><uid>U_DonmaNoktasi</uid><value>{15}</value></cell><cell><uid>U_TKM</uid><value>{16}</value></cell><cell><uid>U_YKM</uid><value>{17}</value></cell><cell><uid>U_AntibiyotikCeza</uid><value>{18}</value></cell><cell><uid>U_SodaCeza</uid><value>{19}</value></cell><cell><uid>U_TemizlikCeza</uid><value>{20}</value></cell><cell><uid>U_AgirlikPuan</uid><value>{2}</value></cell><cell><uid>U_YagCeza</uid><value>{22}</value></cell><cell><uid>U_BrixCeza</uid><value>{23}</value></cell><cell><uid>U_SuCeza</uid><value>{24}</value></cell><cell><uid>U_AsitCeza</uid><value>{25}</value></cell><cell><uid>U_YagPrim</uid><value>{26}</value></cell><cell><uid>U_BrixPrim</uid><value>{27}</value></cell><cell><uid>U_ToplamCeza</uid><value>{28}</value></cell><cell><uid>U_ToplamPrim</uid><value>{29}</value></cell><cell><uid>U_IyilestirmeOncesiMik</uid><value>{30}</value></cell><cell><uid>U_Deger</uid><value>{31}</value></cell><cell><uid>U_IyilestirmeSonrasiMik</uid><value>{32}</value></cell></cells></row>";

        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.SutIyilestirmeXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.SutIyilestirmeXML));
            Functions.CreateUserOrSystemFormComponent<SutIyilestirme>(AIFConn.SutIylstr);

            InitForms();
        }
        public void InitForms()
        {
            try
            {
                frmSutIyilestirme.EnableMenu("1283", false);
                frmSutIyilestirme.EnableMenu("1284", false);
                frmSutIyilestirme.EnableMenu("1286", false);
                oComboYil.Select(DateTime.Now.Year.ToString(), BoSearchKey.psk_ByValue);
                frmSutIyilestirme.ActiveItem = "Item_1";


                oMatrix.CommonSetting.FixedColumnsCount = 3;

            }
            catch (Exception ex)
            {
            }
        }
        public string NoktalamaIsaretiDegistir(string val1)
        {

            string Val = Program.SAPnfi == "," ? val1.ToString().Replace(".", ",") : val1.ToString();

            return Val;

        }
        bool ekleme = false;
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
                        ekleme = true;
                    }
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
                    if (ekleme)
                    {
                        oComboYil.Select(DateTime.Now.Year.ToString(), BoSearchKey.psk_ByValue);
                        frmSutIyilestirme.ActiveItem = "Item_1";
                    }
                    break;
                case BoEventTypes.et_KEY_DOWN:
                    break;
                case BoEventTypes.et_GOT_FOCUS:
                    break;
                case BoEventTypes.et_LOST_FOCUS:
                    if (pVal.ColUID == "Col_0" && !pVal.BeforeAction)
                    {
                        var IyilestirmeOncesiNetMiktar = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_30").Cells.Item(pVal.Row).Specific).Value);
                        var IyilestirmeDegeri = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(pVal.Row).Specific).Value);

                        var sonuc = IyilestirmeOncesiNetMiktar - IyilestirmeDegeri;

                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_31").Cells.Item(pVal.Row).Specific).Value = sonuc.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    }
                    break;
                case BoEventTypes.et_COMBO_SELECT:
                    if (pVal.ItemUID == "Item_1" && !pVal.BeforeAction)
                    {

                    }
                    break;
                case BoEventTypes.et_CLICK:
                    if (pVal.ItemUID == "Item_7" && !pVal.BeforeAction)
                    {
                        try
                        {

                            #region Öncelikle SAP'De kaydı varsa getirilir.
                            ConstVariables.oRecordset.DoQuery("Select * from \"@AIF_SUTIYLSTR\" where \"U_Ay\" = '" + oComboAy.Value.Trim() + "' and \"U_Yil\" = '" + oComboYil.Value.Trim() + "'");

                            if (ConstVariables.oRecordset.RecordCount > 0)
                            {
                                try
                                {
                                    frmSutIyilestirme.Freeze(true);
                                    frmSutIyilestirme.Mode = BoFormMode.fm_FIND_MODE;
                                    oEditDocEntry.Value = ConstVariables.oRecordset.Fields.Item(0).Value.ToString();
                                    btnaddOrUpdate.Item.Click();
                                    oMatrix.AutoResizeColumns();

                                    //Handler.SAPApplication.StatusBar.SetText("Daha önceki girilmiş '" + oComboAy.Selected.Description + "' verisi getirilmiştir. Listeleme işlemi ile devam edebilirsiniz.");
                                }
                                catch (Exception)
                                {
                                }
                                finally
                                {
                                    frmSutIyilestirme.Freeze(false);
                                }
                            }


                            #endregion
                            frmSutIyilestirme.Freeze(true);

                            string sql = "";
                            #region eski süt kabul iyileştirme sorgusu
                            //sql = "WITH BOL AS (SELECT T0.DocNum,T0.[U_BelgeTarihi], AY = MONTH(T0.U_BelgeTarihi),YIL = YEAR(T0.U_BelgeTarihi),T1.[U_TedarikciAdi],T1.[U_Miktar],T1.[U_Sicaklik],T1.U_Miktar * T1.[U_Sicaklik] SCTP,T1.[U_Briks],         T1.U_Miktar * T1.[U_Briks] BRTP, T1.[U_Yag],T1.U_Miktar * T1.[U_Yag] YAGTPL, T1.[U_Yogunluk], T1.U_Miktar * T1.[U_Yogunluk] YOGTPL, T1.[U_pH], T1.U_Miktar * T1.[U_pH] UPHTPL, T1.[U_SH], T1.U_Miktar * T1.[U_SH] USHTPL, T1.[U_KaynatSh],T1.U_Miktar * T1.[U_KaynatSh] KYNTPL, T1.[U_Protein], T1.U_Miktar * T1.[U_Protein] PRTTPL, T1.[U_Laktoz], T1.U_Miktar * T1.[U_Laktoz] LAKTPL, T1.[U_KatilanSu],T1.U_Miktar * T1.[U_KatilanSu] KTLSU, T1.[U_SomatikHucre],T1.U_Miktar * T1.[U_SomatikHucre] SMTTPL,T1.[U_MetilenMavisi], T1.U_Miktar * T1.[U_MetilenMavisi] MVTPL, T1.U_DonNok,T1.U_Miktar * T1.U_DonNok DONTPL, T1.U_Tkm,T1.U_Miktar * T1.U_Tkm UTKMPL,T1.[U_Ykm], T1.U_Miktar * T1.[U_Ykm] YKMTPL, case when T1.U_Koku = 1 then 'İyi' when T1.U_Koku = 2 then 'Kötü' when T1.U_Koku = 3 then 'Orta' else 'Girilmemiş' end U_Koku,T1.[U_Depo], DEPO.WhsName DepoAdi, T2.U_AracPlakasi,T2.U_NetMikTon AS KantarKg, ISNULL(T2.U_NetMikLt,0) AS KantarLitre, T2.U_NetMikLtr1 as IrsaliyeLitre,T2.U_NetMikLtr2 as SayacLitre, CASE WHEN T1.[U_SutTuru] = 1 THEN 'Yağlı Süt' WHEN T1.[U_SutTuru] = 2 THEN 'Yağsız Süt' ELSE 'Boş' END AS 'SÜt Türü', CASE WHEN T1.[U_TemizlikKoku] = 1 THEN 'İyi' WHEN T1.[U_TemizlikKoku] = 2 THEN 'Kötü' WHEN T1.[U_TemizlikKoku] = 3 THEN 'Orta' ELSE 'Boş' END AS 'Temizlik Koku', CASE WHEN T1.[U_Soda] = 1 THEN 'Var' WHEN T1.[U_Soda] = 2 THEN 'Yok' ELSE 'Boş' END AS 'Soda', CASE WHEN T1.[U_Antibiyotik] = 1 THEN 'Var' WHEN T1.[U_Antibiyotik] = 2 THEN 'Yok' ELSE 'Boş' END AS 'Antibiyotik', CASE WHEN T1.[U_Alkol] = 1 THEN 'Kesiyor' WHEN T1.[U_Alkol] = 2 THEN 'Kesmiyor' ELSE 'Boş' END AS 'Alkol', UFD1.Descr, CASE WHEN T1.U_Antibiyotik= 1 THEN (SELECT T1.[U_Miktar]* U_Antibiyotikvar FROM[@AIF_SUTIYI_ANTCZ] WHERE U_BelgeTarihi>=U_GecerlilikBas and U_BelgeTarihi<=U_GecerlilikBit) ELSE 0 END AS AntibiyotikCeza,CASE WHEN(T1.[U_SH]-T1.[U_KaynatSh])>ISNULL(SD.U_OranBas,0)  THEN T1.[U_Miktar]* SD.U_Miktar ELSE 0 END SodaCeza, ISNULL(T1.U_Miktar,0)*ISNULL(TMZ1.U_Ceza,0) TemizlikCeza , ISNULL(T1.U_AgirlikliPuan,0) U_AgirlikliPuan, (ISNULL(T1.U_AgirlikliPuan,0)*T1.U_Miktar) AGRTPL,ISNULL(T1.U_Miktar,0)*ISNULL(YAG.U_Miktar,0) YagCeza,T1.U_Miktar* BRX.U_Miktar BrixCeza, IIF(T1.U_KatilanSu>0, (T1.U_Miktar*(100-T1.U_KatilanSu)),0)/100 SuCeza,ISNULL(T1.U_Miktar,0)*ISNULL(AST.U_Miktar,0) AsitCeza,ISNULL(T1.U_Miktar,0)*ISNULL(YGPR.U_Miktar,0) YagPrim,ISNULL(T1.U_Miktar,0)*ISNULL(BRXPR.U_Miktar,0) BRIXPRIM,ISNULL(CLS1.UnitPrice,0) SozlesmeFiyati,OCR.CardCode FROM[dbo].[@AIF_SUTKABUL] T0 INNER JOIN[dbo].[@AIF_SUTKABUL2] T1 ON T0.DocEntry = T1.DocEntry INNER JOIN OCRD OCR ON T1.U_Tedarikci=OCR.CardCode INNER JOIN[dbo].[@AIF_SUTKABUL1] T2 ON T0.DocEntry = T2.DocEntry INNER JOIN UFD1 ON UFD1.TableID='@AIF_SUTKABUL2' AND UFD1.FieldID= 11 AND OCR.U_Nitelik= UFD1.FldValue INNER JOIN OWHS DEPO ON DEPO.WhsCode= T1.U_Depo LEFT JOIN ( OOAT CLS INNER JOIN OAT1 CLS1 ON CLS.AbsID= CLS1.AgrNo AND CLS1.ItemCode= 'HAM-00001' ) on T0.U_BelgeTarihi>=CLS.StartDate and T0.U_BelgeTarihi<=CLS.EndDate AND CLS.BpCode= OCR.CardCode ";

                            //sql += "OUTER APPLY (SELECT SD1.U_Miktar, SD1.U_OranBas, SD1.U_OranBit, SODA.U_Aciklama FROM [@AIF_SUTIYI_SODA1] SD1 INNER JOIN[@AIF_SUTIYI_SODA] SODA ON T0.U_BelgeTarihi>=SODA.U_GecerlilikBas AND T0.U_BelgeTarihi<=SODA.U_GecerlilikBit WHERE (T1.[U_SH]-T1.[U_KaynatSh]) >=U_OranBas AND(T1.[U_SH]-T1.[U_KaynatSh]) <=U_OranBit AND SODA.DocEntry=SD1.DocEntry )SD ";



                            //sql += "OUTER APPLY (SELECT TM1.U_PuanBas, TM1.U_PuanBit, TM1.U_Ceza, TMZ.U_Aciklama FROM [@AIF_SUTIYI_TEMCZ1] TM1 INNER JOIN[@AIF_SUTIYI_TEMCZ] TMZ ON T0.U_BelgeTarihi>=TMZ.U_GecerlilikBas AND T0.U_BelgeTarihi<=TMZ.U_GecerlilikBit WHERE T1.U_AgirlikliPuan>=TM1.U_PuanBas AND T1.U_AgirlikliPuan<=TM1.U_PuanBit AND TMZ.DocEntry= TM1.DocEntry)TMZ1 ";


                            //sql += "OUTER APPLY (SELECT YGC.U_OranBas, YGC.U_OranBit, YGC.U_Miktar , YG.U_Aciklama FROM [@AIF_SUTIYI_YAGCZ2] YGC INNER JOIN[@AIF_SUTIYI_YAGCZ] YG ON T0.U_BelgeTarihi>=YG.U_GecerlilikBas AND T0.U_BelgeTarihi<=YG.U_GecerlilikBit WHERE YG.DocEntry= YGC.DocEntry AND T1.U_Yag>=YGC.U_OranBas AND T1.U_Yag<=YGC.U_OranBit AND OCR.U_Nitelik= YG.U_Nitelik )YAG ";

                            //sql += "OUTER APPLY(SELECT YGC.U_OranBas, YGC.U_OranBit, YGC.U_Miktar , YG.U_Aciklama FROM [@AIF_SUTIYI_YAGPR2]YGC INNER JOIN[@AIF_SUTIYI_YAGPR] YG ON T0.U_BelgeTarihi>= YG.U_GecerlilikBas AND T0.U_BelgeTarihi <= YG.U_GecerlilikBit WHERE YG.DocEntry = YGC.DocEntry AND T1.U_Yag >= YGC.U_OranBas AND T1.U_Yag <= YGC.U_OranBit AND OCR.U_Nitelik = YG.U_Nitelik )YGPR ";

                            //sql += "OUTER APPLY (SELECT BR1.U_OranBas, BR1.U_OranBit, BR1.U_Miktar, BR.U_Aciklama, BR.U_SutTuru FROM[@AIF_SUTIYI_BRIXCZ1] BR1 ";
                            //sql += "INNER JOIN[@AIF_SUTIYI_BRIXCZ] BR ON T0.U_BelgeTarihi >= BR.U_GecerlilikBas AND T0.U_BelgeTarihi <= BR.U_GecerlilikBit WHERE BR.DocEntry = BR1.DocEntry AND T1.U_Briks >= BR1.U_OranBas AND T1.U_Briks <= BR1.U_OranBit AND OCR.U_Nitelik = BR.U_Nitelik and T1.U_SutTuru = (case when BR.U_SutTuru = N'YAĞLI' then 1 when BR.U_SutTuru = N'YAĞSIZ' then 2 end ))BRX ";

                            //sql += "OUTER APPLY (SELECT BR1.U_OranBas, BR1.U_OranBit, BR1.U_Miktar, BR.U_Aciklama, BR.U_SutTuru FROM[@AIF_SUTIYI_BRIXPR1] BR1 INNER JOIN[@AIF_SUTIYI_BRIXPR] BR ON T0.U_BelgeTarihi >= BR.U_GecerlilikBas AND T0.U_BelgeTarihi <= BR.U_GecerlilikBit ";

                            //sql += "WHERE BR.DocEntry = BR1.DocEntry AND T1.U_Briks >= BR1.U_OranBas AND T1.U_Briks <= BR1.U_OranBit AND OCR.U_Nitelik = BR.U_Nitelik and T1.U_SutTuru = (case when BR.U_SutTuru = N'YAĞLI' then 1 when BR.U_SutTuru = N'YAĞSIZ' then 2 end ))BRXPR ";


                            //sql += "OUTER APPLY(SELECT  PH1.U_OranBas, PH1.U_OranBit, PH1.U_Miktar, PH.U_Aciklama FROM[@AIF_SUTIYI_PHCZ1] PH1 INNER JOIN[@AIF_SUTIYI_PHCZ] PH ON T0.U_BelgeTarihi >= PH.U_GecerlilikBas AND T0.U_BelgeTarihi <= PH.U_GecerlilikBit WHERE PH.DocEntry = PH1.DocEntry AND T1.U_pH >= PH1.U_OranBas AND T1.U_pH <= PH1.U_OranBit AND T1.U_Sicaklik >= PH.U_SicaklikMin AND T1.U_Sicaklik <= PH.U_SicaklikMax)AST) ";


                            //sql += "SELECT AY, YIL, CardCode, U_TedarikciAdi, Descr NITELIK,(SUM(SCTP) / SUM(U_Miktar)) SICAKLIK,(SUM(BRTP) / SUM(U_Miktar)) BRIX,(SUM(YAGTPL) / SUM(U_Miktar)) YAG,(SUM(YOGTPL) / SUM(U_Miktar)) YOGTPL,(SUM(UPHTPL) / SUM(U_Miktar)) PH,(SUM(USHTPL) / SUM(U_Miktar)) SH,(SUM(KYNTPL) / SUM(U_Miktar)) KAYNATSH,(SUM(PRTTPL) / SUM(U_Miktar)) PROTEIN,(SUM(LAKTPL) / SUM(U_Miktar)) LAKTOZ,(SUM(KTLSU) / SUM(U_Miktar)) KATILANSU,(SUM(SMTTPL) / SUM(U_Miktar)) SOMATIK,(SUM(MVTPL) / SUM(U_Miktar)) METILENMAVISI,(SUM(DONTPL) / SUM(U_Miktar)) DONNOK,(SUM(UTKMPL) / SUM(U_Miktar))TKM,(SUM(YKMTPL) / SUM(U_Miktar)) YKM,SUM(AntibiyotikCeza) AntibiyotikCeza,SUM(SodaCeza)SodaCeza,SUM(TemizlikCeza) TemizlikCeza,(SUM(AGRTPL) / SUM(U_Miktar)) AGIRLIKPUAN,SUM(YagCeza)YagCeza,SUM(BrixCeza)BrixCeza,SUM(SuCeza) SuCeza,SUM(AsitCeza) AsitCeza,sum(YagPrim)YagPrim,sum(BRIXPRIM) BRIXPRIM, SUM(AntibiyotikCeza) + SUM(SodaCeza) + SUM(TemizlikCeza) + SUM(YagCeza) + SUM(BrixCeza) + SUM(SuCeza) + SUM(AsitCeza) as ToplamCeza, SUM(YagPrim) + sum(BRIXPRIM) as PrimToplamı,SUM(U_Miktar) + (SUM(YagPrim) + sum(BRIXPRIM)) - (SUM(AntibiyotikCeza) + SUM(SodaCeza) + SUM(TemizlikCeza) + SUM(YagCeza) + SUM(BrixCeza) + SUM(SuCeza) + SUM(AsitCeza)) as 'IyilestirmeOncesiNetMiktar' FROM BOL where AY = '" + oComboAy.Value.Trim() + "' and YIL = '" + oComboYil.Value.Trim() + "' group by AY, YIL, CardCode, U_TedarikciAdi, Descr";
                            #endregion eski süt kabul iyileştirme sorgusu

                            #region Süt kabul iyileştirme sorgusu 11.01.2022
                            sql = "WITH BOL AS (SELECT T0.DocNum, T0.[U_BelgeTarihi], AY = MONTH(T0.U_BelgeTarihi), YIL = YEAR(T0.U_BelgeTarihi), T1.[U_TedarikciAdi], T1.[U_Miktar], T1.[U_Sicaklik], T1.U_Miktar * T1.[U_Sicaklik] SCTP, T1.[U_Briks], T1.U_Miktar * T1.[U_Briks] BRTP,T1.[U_Yag], T1.U_Miktar * T1.[U_Yag] YAGTPL, T1.[U_Yogunluk], T1.U_Miktar * T1.[U_Yogunluk] YOGTPL, T1.[U_pH], T1.U_Miktar * T1.[U_pH] UPHTPL, T1.[U_SH], T1.U_Miktar * T1.[U_SH] USHTPL, T1.[U_KaynatSh], T1.U_Miktar * T1.[U_KaynatSh] KYNTPL, T1.[U_Protein], T1.U_Miktar * T1.[U_Protein] PRTTPL, T1.[U_Laktoz], T1.U_Miktar * T1.[U_Laktoz] LAKTPL, T1.[U_KatilanSu], T1.U_Miktar * T1.[U_KatilanSu] KTLSU,T1.[U_SomatikHucre], T1.U_Miktar * T1.[U_SomatikHucre] SMTTPL, T1.[U_MetilenMavisi], T1.U_Miktar * T1.[U_MetilenMavisi] MVTPL, T1.U_DonNok, T1.U_Miktar * T1.U_DonNok DONTPL, T1.U_Tkm, T1.U_Miktar * T1.U_Tkm UTKMPL, T1.[U_Ykm], T1.U_Miktar * T1.[U_Ykm] YKMTPL, CASE WHEN T1.U_Koku = 1 THEN 'İyi' WHEN T1.U_Koku = 2 THEN 'Kötü' WHEN T1.U_Koku = 3 THEN 'Orta' ELSE 'Girilmemiş' END U_Koku, T1.[U_Depo], DEPO.WhsName DepoAdi, T2.U_AracPlakasi, T2.U_NetMikTon AS KantarKg, ISNULL(T2.U_NetMikLt, 0) AS KantarLitre, T2.U_NetMikLtr1 AS IrsaliyeLitre, T2.U_NetMikLtr2 AS SayacLitre, CASE WHEN T1.[U_SutTuru] = 1 THEN 'Yağlı Süt' WHEN T1.[U_SutTuru] = 2 THEN 'Yağsız Süt' ELSE 'Boş' END AS 'SÜt Türü', CASE WHEN T1.[U_TemizlikKoku] = 1 THEN 'İyi' WHEN T1.[U_TemizlikKoku] = 2 THEN 'Kötü' WHEN T1.[U_TemizlikKoku] = 3 THEN 'Orta' ELSE 'Boş' END AS 'Temizlik Koku', CASE WHEN T1.[U_Soda] = 1 THEN 'Var' WHEN T1.[U_Soda] = 2 THEN 'Yok' ELSE 'Boş' END AS 'Soda', CASE WHEN T1.[U_Antibiyotik] = 1 THEN 'Var' WHEN T1.[U_Antibiyotik] = 2 THEN 'Yok' ELSE 'Boş' END AS 'Antibiyotik', CASE WHEN T1.[U_Alkol] = 1 THEN 'Kesiyor' WHEN T1.[U_Alkol] = 2 THEN 'Kesmiyor' ELSE 'Boş' END AS 'Alkol', UFD1.Descr, CASE WHEN T1.U_Antibiyotik = 1 THEN (SELECT T1.[U_Miktar] * U_Antibiyotikvar FROM[@AIF_SUTIYI_ANTCZ] WHERE U_BelgeTarihi >= U_GecerlilikBas AND U_BelgeTarihi <= U_GecerlilikBit) ELSE 0 END AS AntibiyotikCeza, ISNULL(T1.[U_Miktar], 0) - CASE WHEN(T1.[U_SH] - T1.[U_KaynatSh]) > ISNULL(SD.U_OranBas, 0) THEN ISNULL(T1.[U_Miktar], 0) * ISNULL(SD.U_Miktar, 0) WHEN(T1.[U_SH] - T1.[U_KaynatSh]) < ISNULL(SD.U_OranBas, 0) THEN ISNULL(T1.[U_Miktar], 0) * ISNULL(SD.U_Miktar, 0) ELSE ISNULL(T1.[U_Miktar], 0) END SodaCeza, ISNULL(T1.U_Miktar, 0) - (ISNULL(T1.U_Miktar, 0) * ISNULL(TMZ1.U_Ceza, 0)) TemizlikCeza, ISNULL(T1.U_AgirlikliPuan, 0) U_AgirlikliPuan, ISNULL(ISNULL(T1.U_Miktar, 0) - (ISNULL(T1.U_AgirlikliPuan, 0) * NULLIF(T1.U_Miktar, 0)), 0) AGRTPL, ISNULL(ISNULL(ISNULL(T1.U_Miktar, 0) - (ISNULL(T1.U_Miktar, 0) * NULLIF(YAG.U_Miktar, 0)), 0), 0) YagCeza, ISNULL(ISNULL(T1.U_Miktar, 0) - (ISNULL(T1.U_Miktar, 0) * NULLIF(BRX.U_Miktar, 0)), 0) BrixCeza, ISNULL(T1.U_Miktar, 0) - IIF((T1.U_KatilanSu) > 0, ((100 - T1.U_KatilanSu) * T1.U_Miktar) / 100, T1.U_Miktar) SuCeza, ISNULL(ISNULL(T1.U_Miktar, 0) - (ISNULL(T1.U_Miktar, 0) * NULLIF(AST.U_Miktar, 0)), 0) AsitCeza, ISNULL(ISNULL(T1.U_Miktar, 0) - (ISNULL(T1.U_Miktar, 0) * NULLIF(YGPR.U_Miktar, 0)), 0) YagPrim, ISNULL(ISNULL(T1.U_Miktar, 0) - (ISNULL(T1.U_Miktar, 0) * NULLIF(BRXPR.U_Miktar, 0)), 0) BRIXPRIM, ISNULL(CLS1.UnitPrice, 0) SozlesmeFiyati, OCR.CardCode FROM[dbo].[@AIF_SUTKABUL] T0 INNER JOIN[dbo].[@AIF_SUTKABUL2] T1 ON T0.DocEntry = T1.DocEntry INNER JOIN OCRD OCR ON T1.U_Tedarikci = OCR.CardCode INNER JOIN[dbo].[@AIF_SUTKABUL1] T2 ON T0.DocEntry = T2.DocEntry INNER JOIN UFD1 ON UFD1.TableID = '@AIF_SUTKABUL2' AND UFD1.FieldID = 11 AND OCR.U_Nitelik = UFD1.FldValue INNER JOIN OWHS DEPO ON DEPO.WhsCode = T1.U_Depo LEFT JOIN(OOAT CLS INNER JOIN OAT1 CLS1 ON CLS.AbsID = CLS1.AgrNo AND CLS1.ItemCode = 'HAM-00001') ON T0.U_BelgeTarihi >= CLS.StartDate AND T0.U_BelgeTarihi <= CLS.EndDate AND CLS.BpCode = OCR.CardCode ";

                            sql += "OUTER APPLY (SELECT SD1.U_Miktar, SD1.U_OranBas, SD1.U_OranBit, SODA.U_Aciklama FROM[@AIF_SUTIYI_SODA1] SD1 INNER JOIN[@AIF_SUTIYI_SODA] SODA ON T0.U_BelgeTarihi >= SODA.U_GecerlilikBas AND T0.U_BelgeTarihi <= SODA.U_GecerlilikBit WHERE(T1.[U_SH] - T1.[U_KaynatSh]) >= U_OranBas AND(T1.[U_SH] - T1.[U_KaynatSh]) <= U_OranBit AND SODA.DocEntry = SD1.DocEntry )SD ";

                            sql += "OUTER APPLY (SELECT TM1.U_PuanBas, TM1.U_PuanBit, TM1.U_Ceza, TMZ.U_Aciklama FROM[@AIF_SUTIYI_TEMCZ1] TM1 INNER JOIN[@AIF_SUTIYI_TEMCZ] TMZ ON T0.U_BelgeTarihi >= TMZ.U_GecerlilikBas AND T0.U_BelgeTarihi <= TMZ.U_GecerlilikBit WHERE T1.U_AgirlikliPuan >= TM1.U_PuanBas AND T1.U_AgirlikliPuan <= TM1.U_PuanBit AND TMZ.DocEntry = TM1.DocEntry)TMZ1 ";

                            sql += "OUTER APPLY (SELECT YGC.U_OranBas, YGC.U_OranBit, YGC.U_Miktar, YG.U_Aciklama FROM[@AIF_SUTIYI_YAGCZ2] YGC INNER JOIN[@AIF_SUTIYI_YAGCZ] YG ON T0.U_BelgeTarihi >= YG.U_GecerlilikBas AND T0.U_BelgeTarihi <= YG.U_GecerlilikBit WHERE YG.DocEntry = YGC.DocEntry AND T1.U_Yag >= YGC.U_OranBas AND T1.U_Yag <= YGC.U_OranBit AND OCR.U_Nitelik = YG.U_Nitelik)YAG ";

                            sql += "OUTER APPLY (SELECT YGC.U_OranBas, YGC.U_OranBit, YGC.U_Miktar, YG.U_Aciklama FROM[@AIF_SUTIYI_YAGPR2]YGC INNER JOIN[@AIF_SUTIYI_YAGPR] YG ON T0.U_BelgeTarihi >= YG.U_GecerlilikBas AND T0.U_BelgeTarihi <= YG.U_GecerlilikBit WHERE YG.DocEntry = YGC.DocEntry AND T1.U_Yag >= YGC.U_OranBas AND T1.U_Yag <= YGC.U_OranBit AND OCR.U_Nitelik = YG.U_Nitelik)YGPR ";

                            sql += "OUTER APPLY (SELECT BR1.U_OranBas, BR1.U_OranBit, BR1.U_Miktar, BR.U_Aciklama, BR.U_SutTuru FROM[@AIF_SUTIYI_BRIXCZ1] BR1 INNER JOIN[@AIF_SUTIYI_BRIXCZ] BR ON T0.U_BelgeTarihi >= BR.U_GecerlilikBas AND T0.U_BelgeTarihi <= BR.U_GecerlilikBit WHERE BR.DocEntry = BR1.DocEntry AND T1.U_Briks >= BR1.U_OranBas AND T1.U_Briks <= BR1.U_OranBit AND OCR.U_Nitelik = BR.U_Nitelik AND T1.U_SutTuru = (CASE WHEN BR.U_SutTuru = N'YAĞLI' THEN 1 WHEN BR.U_SutTuru = N'YAĞSIZ' THEN 2 END))BRX ";

                            sql += "OUTER APPLY (SELECT BR1.U_OranBas, BR1.U_OranBit, BR1.U_Miktar, BR.U_Aciklama, BR.U_SutTuru FROM[@AIF_SUTIYI_BRIXPR1] BR1  INNER JOIN[@AIF_SUTIYI_BRIXPR] BR ON T0.U_BelgeTarihi >= BR.U_GecerlilikBas AND T0.U_BelgeTarihi <= BR.U_GecerlilikBit WHERE BR.DocEntry = BR1.DocEntry AND T1.U_Briks >= BR1.U_OranBas AND T1.U_Briks <= BR1.U_OranBit AND OCR.U_Nitelik = BR.U_Nitelik AND T1.U_SutTuru = (CASE WHEN BR.U_SutTuru = N'YAĞLI' THEN 1 WHEN BR.U_SutTuru = N'YAĞSIZ' THEN 2 END))BRXPR ";

                            sql += "OUTER APPLY (SELECT PH1.U_OranBas, PH1.U_OranBit, PH1.U_Miktar, PH.U_Aciklama FROM[@AIF_SUTIYI_PHCZ1] PH1 INNER JOIN[@AIF_SUTIYI_PHCZ] PH ON T0.U_BelgeTarihi >= PH.U_GecerlilikBas AND T0.U_BelgeTarihi <= PH.U_GecerlilikBit WHERE PH.DocEntry = PH1.DocEntry AND T1.U_pH >= PH1.U_OranBas AND T1.U_pH <= PH1.U_OranBit AND T1.U_Sicaklik >= PH.U_SicaklikMin AND T1.U_Sicaklik <= PH.U_SicaklikMax)AST) ";

                            sql += "SELECT AY, YIL, CardCode, U_TedarikciAdi, Descr NITELIK, (SUM(SCTP) / SUM(U_Miktar)) SICAKLIK, (SUM(BRTP) / SUM(U_Miktar)) BRIX, (SUM(YAGTPL) / SUM(U_Miktar)) YAG, (SUM(YOGTPL) / SUM(U_Miktar)) YOGTPL, (SUM(UPHTPL) / SUM(U_Miktar)) PH, (SUM(USHTPL) / SUM(U_Miktar)) SH, (SUM(KYNTPL) / SUM(U_Miktar)) KAYNATSH, (SUM(PRTTPL) / SUM(U_Miktar)) PROTEIN, (SUM(LAKTPL) / SUM(U_Miktar)) LAKTOZ, (SUM(KTLSU) / SUM(U_Miktar)) KATILANSU, (SUM(SMTTPL) / SUM(U_Miktar)) SOMATIK, (SUM(MVTPL) / SUM(U_Miktar)) METILENMAVISI, (SUM(DONTPL) / SUM(U_Miktar)) DONNOK, (SUM(UTKMPL) / SUM(U_Miktar))TKM, (SUM(YKMTPL) / SUM(U_Miktar)) YKM, SUM(AntibiyotikCeza) AntibiyotikCeza, SUM(SodaCeza)SodaCeza, SUM(TemizlikCeza) TemizlikCeza, (SUM(AGRTPL) / SUM(U_Miktar)) AGIRLIKPUAN, SUM(YagCeza)YagCeza, SUM(isnull(BrixCeza, 0))BrixCeza, SUM(SuCeza) SuCeza, SUM(AsitCeza) AsitCeza, sum(YagPrim)YagPrim, sum(BRIXPRIM) BRIXPRIM, SUM(isnull(AntibiyotikCeza, 0)) + SUM(isnull(SodaCeza, 0)) + SUM(isnull(TemizlikCeza, 0)) + SUM(isnull(YagCeza, 0)) + SUM(isnull(BrixCeza, 0)) + SUM(isnull(SuCeza, 0)) + SUM(isnull(AsitCeza, 0)) AS ToplamCeza, SUM(isnull(YagPrim, 0)) +sum(isnull(BRIXPRIM, 0)) AS PrimToplamı, SUM(isnull(U_Miktar, 0)) +(SUM(isnull(YagPrim, 0)) + sum(isnull(BRIXPRIM, 0))) - (SUM(isnull(AntibiyotikCeza, 0)) + SUM(isnull(SodaCeza, 0)) + SUM(isnull(TemizlikCeza, 0)) + SUM(isnull(YagCeza, 0)) + SUM(isnull(BrixCeza, 0)) + SUM(isnull(SuCeza, 0)) + SUM(isnull(AsitCeza, 0))) AS 'IyilestirmeOncesiNetMiktar' FROM BOL WHERE AY = '" + oComboAy.Value.Trim() + "' and YIL = '" + oComboYil.Value.Trim() + "' GROUP BY AY, YIL, CardCode, U_TedarikciAdi, Descr ";
                            #endregion Süt kabul iyileştirme sorgusu 11.01.2022

                            ConstVariables.oRecordset.DoQuery(sql);

                            string xmll = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                            XDocument xDoc = XDocument.Parse(xmll);
                            XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                            selectDetails = (from tx in xDoc.Descendants(ns + "Row")
                                             select new Details()
                                             {
                                                 SaticiKodu = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "CardCode" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                 SaticiAdi = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_TedarikciAdi" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                 Nitelik = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "NITELIK" select new XElement(y.Element(ns + "Value"))).First().Value,
                                                 Sicaklik = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "SICAKLIK" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 Brix = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "BRIX" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 Yag = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "YAG" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 YogTPL = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "YOGTPL" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 PH = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "PH" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 SH = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "SH" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 KaynatSH = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "KAYNATSH" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 Protein = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "PROTEIN" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 Laktoz = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "LAKTOZ" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 KatilanSu = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "KATILANSU" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 Somatik = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "SOMATIK" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 MetilenMavisi = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "METILENMAVISI" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 DonmaNoktasi = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "DONNOK" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 TKM = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "TKM" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 YKM = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "YKM" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 AntibiyotikCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "AntibiyotikCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 SodaCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "SodaCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 TemizlikCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "TemizlikCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 AgirlikPuan = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "AGIRLIKPUAN" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 YagCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "YagCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 BrixCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "BrixCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 SuCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "SuCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 AsitCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "AsitCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 YagPrim = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "YagPrim" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 BrixPrim = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "BRIXPRIM" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 ToplamCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "ToplamCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 ToplamPrim = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "PrimToplamı" select new XElement(y.Element(ns + "Value"))).First().Value),
                                                 IyilestirmeOncesiNetMiktar = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "IyilestirmeOncesiNetMiktar" select new XElement(y.Element(ns + "Value"))).First().Value),
                                             }).ToList();


                            sql = "Select \"U_TedarikciAdi\",\"U_TedarikciKodu\",\"U_Nitelik\",\"U_Sicaklik\",\"U_Brix\",\"U_Yag\",\"U_YogurtPL\",\"U_PH\",\"U_SH\",\"U_KaynatSH\",\"U_Protein\",\"U_Laktoz\",\"U_KatilanSu\",\"U_Somatik\",\"U_MetilenMavisi\",\"U_DonmaNoktasi\",\"U_TKM\",\"U_YKM\",\"U_AntibiyotikCeza\",\"U_SodaCeza\",\"U_TemizlikCeza\",\"U_AgirlikPuan\",\"U_YagCeza\",\"U_BrixCeza\",\"U_SuCeza\",\"U_AsitCeza\",\"U_YagPrim\",\"U_BrixPrim\",\"U_ToplamCeza\",\"U_ToplamPrim\",\"U_IyilestirmeOncesiMik\",\"U_Deger\" from \"@AIF_SUTIYLSTR\" as T0 INNER JOIN \"@AIF_SUTIYLSTR1\" as T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" where T0.\"U_Ay\" = '" + oComboAy.Value.Trim() + "' and T0.\"U_Yil\" = '" + oComboYil.Value.Trim() + "' ";

                            ConstVariables.oRecordset.DoQuery(sql);

                            xmll = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
                            xDoc = XDocument.Parse(xmll);
                            ns = "http://www.sap.com/SBO/SDK/DI";


                            sapDetails = (from tx in xDoc.Descendants(ns + "Row")
                                          select new Details()
                                          {
                                              SaticiKodu = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_TedarikciKodu" select new XElement(y.Element(ns + "Value"))).First().Value,
                                              SaticiAdi = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_TedarikciAdi" select new XElement(y.Element(ns + "Value"))).First().Value,
                                              Nitelik = (from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Nitelik" select new XElement(y.Element(ns + "Value"))).First().Value,
                                              Sicaklik = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Sicaklik" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              Brix = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Brix" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              Yag = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Yag" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              YogTPL = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YogurtPL" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              PH = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_PH" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              SH = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SH" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              KaynatSH = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_KaynatSH" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              Protein = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Protein" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              Laktoz = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Laktoz" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              KatilanSu = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_KatilanSu" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              Somatik = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Somatik" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              MetilenMavisi = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_MetilenMavisi" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              DonmaNoktasi = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_DonmaNoktasi" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              TKM = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_TKM" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              YKM = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YKM" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              AntibiyotikCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_AntibiyotikCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              SodaCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SodaCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              TemizlikCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_TemizlikCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              AgirlikPuan = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_AgirlikPuan" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              YagCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YagCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              BrixCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_BrixCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              SuCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SuCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              AsitCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SuCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              YagPrim = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YagPrim" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              BrixPrim = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_BrixPrim" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              ToplamCeza = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_ToplamCeza" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              ToplamPrim = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_ToplamPrim" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              IyilestirmeOncesiNetMiktar = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_IyilestirmeOncesiMik" select new XElement(y.Element(ns + "Value"))).First().Value),
                                              Deger = parseNumber.parservalues<double>((from y in tx.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_Deger" select new XElement(y.Element(ns + "Value"))).First().Value),
                                          }).ToList();


                            if (ConstVariables.oRecordset.RecordCount == 0)
                            {
                                string data = string.Join("", selectDetails.Select(s => string.Format(row, s.SaticiAdi, s.SaticiKodu, s.Nitelik.ToString(System.Globalization.CultureInfo.InvariantCulture), s.Sicaklik.ToString(System.Globalization.CultureInfo.InvariantCulture), s.Brix.ToString(System.Globalization.CultureInfo.InvariantCulture), s.Yag.ToString(System.Globalization.CultureInfo.InvariantCulture), s.YogTPL.ToString(System.Globalization.CultureInfo.InvariantCulture), s.PH.ToString(System.Globalization.CultureInfo.InvariantCulture), s.SH.ToString(System.Globalization.CultureInfo.InvariantCulture), s.KaynatSH.ToString(System.Globalization.CultureInfo.InvariantCulture), s.Protein.ToString(System.Globalization.CultureInfo.InvariantCulture), s.Laktoz.ToString(System.Globalization.CultureInfo.InvariantCulture), s.KatilanSu.ToString(System.Globalization.CultureInfo.InvariantCulture), s.Somatik.ToString(System.Globalization.CultureInfo.InvariantCulture), s.MetilenMavisi.ToString(System.Globalization.CultureInfo.InvariantCulture), s.DonmaNoktasi.ToString(System.Globalization.CultureInfo.InvariantCulture), s.TKM.ToString(System.Globalization.CultureInfo.InvariantCulture), s.YKM.ToString(System.Globalization.CultureInfo.InvariantCulture), s.AntibiyotikCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.SodaCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.TemizlikCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.AgirlikPuan.ToString(System.Globalization.CultureInfo.InvariantCulture), s.YagCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.BrixCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.SuCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.AsitCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.YagPrim.ToString(System.Globalization.CultureInfo.InvariantCulture), s.BrixPrim.ToString(System.Globalization.CultureInfo.InvariantCulture), s.ToplamCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.ToplamPrim.ToString(System.Globalization.CultureInfo.InvariantCulture), s.IyilestirmeOncesiNetMiktar.ToString(System.Globalization.CultureInfo.InvariantCulture), 0, s.IyilestirmeOncesiNetMiktar.ToString(System.Globalization.CultureInfo.InvariantCulture))));

                                string s1 = string.Format(header, data);
                                frmSutIyilestirme.DataSources.DBDataSources.Item("@AIF_SUTIYLSTR1").LoadFromXML(string.Format(header, data));

                                oMatrix.AutoResizeColumns();
                            }
                            else
                            {
                                List<Details> tempdetails = new List<Details>();
                                foreach (var item in selectDetails)
                                {
                                    var exist = sapDetails.Where(x => x.SaticiKodu.Contains(item.SaticiKodu)).ToList();

                                    if (exist.Count == 0)
                                    {
                                        tempdetails.AddRange(selectDetails.Where(x => x.SaticiKodu == item.SaticiKodu).ToList());
                                    }
                                }


                                if (tempdetails.Count > 0)
                                {
                                    frmSutIyilestirme.Mode = BoFormMode.fm_UPDATE_MODE;
                                    sapDetails.AddRange(tempdetails);
                                }

                                string data = string.Join("", sapDetails.Select(s => string.Format(row, s.SaticiAdi, s.SaticiKodu, s.Nitelik.ToString(System.Globalization.CultureInfo.InvariantCulture), s.Sicaklik.ToString(System.Globalization.CultureInfo.InvariantCulture), s.Brix.ToString(System.Globalization.CultureInfo.InvariantCulture), s.Yag.ToString(System.Globalization.CultureInfo.InvariantCulture), s.YogTPL.ToString(System.Globalization.CultureInfo.InvariantCulture), s.PH.ToString(System.Globalization.CultureInfo.InvariantCulture), s.SH.ToString(System.Globalization.CultureInfo.InvariantCulture), s.KaynatSH.ToString(System.Globalization.CultureInfo.InvariantCulture), s.Protein.ToString(System.Globalization.CultureInfo.InvariantCulture), s.Laktoz.ToString(System.Globalization.CultureInfo.InvariantCulture), s.KatilanSu.ToString(System.Globalization.CultureInfo.InvariantCulture), s.Somatik.ToString(System.Globalization.CultureInfo.InvariantCulture), s.MetilenMavisi.ToString(System.Globalization.CultureInfo.InvariantCulture), s.DonmaNoktasi.ToString(System.Globalization.CultureInfo.InvariantCulture), s.TKM.ToString(System.Globalization.CultureInfo.InvariantCulture), s.YKM.ToString(System.Globalization.CultureInfo.InvariantCulture), s.AntibiyotikCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.SodaCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.TemizlikCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.AgirlikPuan.ToString(System.Globalization.CultureInfo.InvariantCulture), s.YagCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.BrixCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.SuCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.AsitCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.YagPrim.ToString(System.Globalization.CultureInfo.InvariantCulture), s.BrixPrim.ToString(System.Globalization.CultureInfo.InvariantCulture), s.ToplamCeza.ToString(System.Globalization.CultureInfo.InvariantCulture), s.ToplamPrim.ToString(System.Globalization.CultureInfo.InvariantCulture), s.IyilestirmeOncesiNetMiktar.ToString(System.Globalization.CultureInfo.InvariantCulture), s.Deger.ToString(System.Globalization.CultureInfo.InvariantCulture), s.IyilestirmeOncesiNetMiktar - s.Deger)));

                                string s1 = string.Format(header, data);
                                frmSutIyilestirme.DataSources.DBDataSources.Item("@AIF_SUTIYLSTR1").LoadFromXML(string.Format(header, data));

                                oMatrix.AutoResizeColumns();

                            }
                        }
                        catch (Exception ex)
                        {
                        }
                        finally
                        {
                            frmSutIyilestirme.Freeze(false);
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
                oComboYil.Select(DateTime.Now.Year.ToString(), BoSearchKey.psk_ByValue);
                frmSutIyilestirme.ActiveItem = "Item_1";
            }
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }

        List<Details> selectDetails = new List<Details>();
        List<Details> sapDetails = new List<Details>();

        public class Details
        {
            public string SaticiKodu { get; set; }

            public string SaticiAdi { get; set; }

            public string Nitelik { get; set; }

            public double Sicaklik { get; set; }

            public double Brix { get; set; }

            public double Yag { get; set; }

            public double YogTPL { get; set; }

            public double PH { get; set; }

            public double SH { get; set; }

            public double KaynatSH { get; set; }

            public double Protein { get; set; }

            public double Laktoz { get; set; }

            public double KatilanSu { get; set; }

            public double Somatik { get; set; }

            public double MetilenMavisi { get; set; }

            public double DonmaNoktasi { get; set; }

            public double TKM { get; set; }

            public double YKM { get; set; }

            public double AntibiyotikCeza { get; set; }

            public double SodaCeza { get; set; }

            public double TemizlikCeza { get; set; }

            public double AgirlikPuan { get; set; }

            public double YagCeza { get; set; }

            public double BrixCeza { get; set; }

            public double SuCeza { get; set; }

            public double AsitCeza { get; set; }

            public double YagPrim { get; set; }

            public double BrixPrim { get; set; }

            public double ToplamCeza { get; set; }

            public double ToplamPrim { get; set; }

            public double IyilestirmeOncesiNetMiktar { get; set; }

            public double Deger { get; set; }
        }
    }
}
