using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using AIF.UVT.SAPB1.HelperClass;
using AIF.UVT.SAPB1.Models;
using DataTables;
using Microsoft.Win32;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing.Design;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class SutKabul : IUserForm, IMenuEvents, IRightEvents
    {
        [ItemAtt(AIFConn.SutKabulUID)]
        public SAPbouiCOM.Form frmSutKabul;

        #region Değişkenler

        [ItemAtt("Item_1")]
        public SAPbouiCOM.Folder oFolderAracBilgileri;

        [ItemAtt("Item_77")]
        public SAPbouiCOM.Matrix oMatrixAnalizBilgileri;

        [ItemAtt("Item_124")]
        public SAPbouiCOM.Matrix oMatrixAracTankVerileri;

        [ItemAtt("Item_25")]
        public SAPbouiCOM.Button btnKantarAnaveri;

        [ItemAtt("Item_74")]
        public SAPbouiCOM.Button btnAnalizKaydet;

        [ItemAtt("Item_75")]
        public SAPbouiCOM.Button btnAnalizTemizle;

        [ItemAtt("Item_128")]
        public SAPbouiCOM.EditText oEditDocEntry;

        [ItemAtt("Item_166")]
        public SAPbouiCOM.ComboBox oComboBelgeDurumu;

        [ItemAtt("Item_159")]
        public SAPbouiCOM.EditText oEditSatici;

        [ItemAtt("Item_161")]
        public SAPbouiCOM.EditText oEditSaticiAdi;

        [ItemAtt("Item_32")]
        public SAPbouiCOM.EditText edtAnalizMiktar;

        [ItemAtt("Item_34")]
        public SAPbouiCOM.EditText edtAnalizTedarikci;

        [ItemAtt("Item_134")]
        public SAPbouiCOM.EditText edtAnalizTedarikciAdi;

        [ItemAtt("Item_35")]
        public SAPbouiCOM.EditText edtAnalizSicaklik;

        [ItemAtt("Item_39")]
        public SAPbouiCOM.EditText edtAnalizBriks;

        [ItemAtt("Item_41")]
        public SAPbouiCOM.EditText edtAnalizYag;

        [ItemAtt("Item_43")]
        public SAPbouiCOM.EditText edtAnalizYogunluk;

        [ItemAtt("Item_45")]
        public SAPbouiCOM.EditText edtAnalizpH;

        [ItemAtt("Item_47")]
        public SAPbouiCOM.EditText edtAnalizSH;

        [ItemAtt("Item_137")]
        public SAPbouiCOM.ComboBox oComboAnalizSutTuru;

        [ItemAtt("Item_157")]
        public SAPbouiCOM.ComboBox oComboAnalizTemizkikKoku;

        [ItemAtt("Item_152")]
        public SAPbouiCOM.ComboBox oComboAnalizSoda;

        [ItemAtt("Item_153")]
        public SAPbouiCOM.ComboBox oComboAnalizAntibiyotik;

        [ItemAtt("Item_57")]
        public SAPbouiCOM.EditText edtAnalizYkm;

        [ItemAtt("Item_150")]
        public SAPbouiCOM.ComboBox oComboAnalizAlkol;

        [ItemAtt("Item_61")]
        public SAPbouiCOM.EditText edtAnalizKaynatSh;

        [ItemAtt("Item_63")]
        public SAPbouiCOM.EditText edtAnalizDepo;

        [ItemAtt("Item_65")]
        public SAPbouiCOM.EditText edtAnalizProtein;

        [ItemAtt("Item_67")]
        public SAPbouiCOM.EditText edtAnalizLaktoz;

        [ItemAtt("Item_69")]
        public SAPbouiCOM.EditText edtAnalizKatilanSu;

        [ItemAtt("Item_71")]
        public SAPbouiCOM.EditText edtAnalizSomatikHucre;

        [ItemAtt("Item_73")]
        public SAPbouiCOM.EditText edtAnalizMetilenMavisi;

        [ItemAtt("Item_36")]
        public SAPbouiCOM.EditText edtAnalizRow;



        [ItemAtt("Item_74")]
        public SAPbouiCOM.Button btnKaydet;

        [ItemAtt("Item_27")]
        public SAPbouiCOM.CheckBox oCheckKantaraGore;

        [ItemAtt("Item_28")]
        public SAPbouiCOM.CheckBox oCheckBelgeTipineGore;

        [ItemAtt("Item_29")]
        public SAPbouiCOM.CheckBox oCheckSayacaGore;

        [ItemAtt("Item_145")]
        public SAPbouiCOM.CheckBox oCheckIrsaliye;

        [ItemAtt("Item_146")]
        public SAPbouiCOM.CheckBox oCheckBelgesiz;

        [ItemAtt("Item_21")]
        public SAPbouiCOM.EditText oedtIrsaliyeNetSutMiktari;

        [ItemAtt("Item_164")]
        public SAPbouiCOM.EditText oEditBelgeNo;

        [ItemAtt("Item_130")]
        public SAPbouiCOM.EditText oEditBelgeTarihi;

        [ItemAtt("1")]
        public SAPbouiCOM.Button oButtonAddOrUpdate;

        [ItemAtt("Item_63")]
        public SAPbouiCOM.EditText oEditDepo;

        [ItemAtt("Item_26")]
        public SAPbouiCOM.CheckBox oChkTekilGiris;

        [ItemAtt("Item_49")]
        public SAPbouiCOM.CheckBox oChkCogulGiris;

        [ItemAtt("Item_144")]
        public SAPbouiCOM.EditText oEditKantarNo;

        [ItemAtt("Item_164")]
        public SAPbouiCOM.EditText oEditOlusanBelgeNo;

        [ItemAtt("Item_24")]
        public SAPbouiCOM.EditText oEditSayacBilgileriToplam;

        [ItemAtt("Item_121")]
        public SAPbouiCOM.EditText edtAracTedarikci;

        [ItemAtt("Item_132")]
        public SAPbouiCOM.EditText edtAracTedarikciAdi;

        [ItemAtt("Item_99")]
        public SAPbouiCOM.ComboBox oComboAnalizKoku;

        [ItemAtt("Item_141")]
        public SAPbouiCOM.EditText edtAnalizDonmaNoktasi;

        [ItemAtt("Item_148")]
        public SAPbouiCOM.EditText edtAnalizTKM;

        [ItemAtt("Item_167")]
        public SAPbouiCOM.ComboBox oComboAracTankSoda;

        [ItemAtt("Item_168")]
        public SAPbouiCOM.ComboBox oComboNitelik;

        [ItemAtt("Item_144")]
        public SAPbouiCOM.EditText edtFisNo;

        [ItemAtt("Item_5")]
        public SAPbouiCOM.EditText edtAracPlakasi;

        [ItemAtt("Item_7")]
        public SAPbouiCOM.EditText edtSurucu;

        [ItemAtt("Item_12")]
        public SAPbouiCOM.EditText edtDoluKantarMiktar;

        [ItemAtt("Item_15")]
        public SAPbouiCOM.EditText edtBosKantarMiktar;

        [ItemAtt("Item_18")]
        public SAPbouiCOM.EditText edtKantarNetSutMiktari;

        [ItemAtt("Item_180")]
        public SAPbouiCOM.EditText edtKantarNetSutMiktariLitre;

        [ItemAtt("Item_170")]
        public SAPbouiCOM.EditText edtSutTemizligi;

        [ItemAtt("Item_172")]
        public SAPbouiCOM.EditText edtTankTemizligi;

        [ItemAtt("Item_174")]
        public SAPbouiCOM.EditText edtEkipmanTemizligi;

        [ItemAtt("Item_176")]
        public SAPbouiCOM.EditText edtAracKasasiTemizligi;

        [ItemAtt("Item_178")]
        public SAPbouiCOM.EditText edtAgirlikliPuan;

        [ItemAtt("Item_183")]
        public SAPbouiCOM.EditText edtIrsaliyeNo;

        [ItemAtt("Item_186")]
        public SAPbouiCOM.EditText edtSutKabulTarihi;

        [ItemAtt("Item_187")]
        public SAPbouiCOM.EditText edtSutKabulSaati;

        [ItemAtt("Item_147")]
        public SAPbouiCOM.Button oBtnSutKabulIsleminiTamamla;

        #region Araç Tank

        [ItemAtt("Item_79")]
        public SAPbouiCOM.EditText edtAracTankTankNumarasi;

        [ItemAtt("Item_81")]
        public SAPbouiCOM.EditText edtAracTankLabMiktar;

        [ItemAtt("Item_83")]
        public SAPbouiCOM.EditText edtAracTankDerece;

        [ItemAtt("Item_85")]
        public SAPbouiCOM.EditText edtAracTankBrix;

        [ItemAtt("Item_87")]
        public SAPbouiCOM.EditText edtAracTankYag;

        [ItemAtt("Item_89")]
        public SAPbouiCOM.EditText edtAracTankDans;

        [ItemAtt("Item_91")]
        public SAPbouiCOM.EditText edtAracTankpH;

        [ItemAtt("Item_93")]
        public SAPbouiCOM.EditText edtAracTankSH;

        [ItemAtt("Item_95")]
        public SAPbouiCOM.EditText edtAracTankKaynamisSH;

        [ItemAtt("Item_154")]
        public SAPbouiCOM.ComboBox oComboAracTankAntibiyotik;

        [ItemAtt("Item_151")]
        public SAPbouiCOM.ComboBox oComboAracTankAlkol;

        [ItemAtt("Item_103")]
        public SAPbouiCOM.EditText edtAracTankYagsiz;

        [ItemAtt("Item_105")]
        public SAPbouiCOM.EditText edtAracTankProtein;

        [ItemAtt("Item_107")]
        public SAPbouiCOM.EditText edtAracTankDonNok;

        [ItemAtt("Item_109")]
        public SAPbouiCOM.EditText edtAracTankLaktoz;

        [ItemAtt("Item_111")]
        public SAPbouiCOM.EditText edtAracTankSu;

        [ItemAtt("Item_155")]
        public SAPbouiCOM.ComboBox oComboAracTankKoku;

        [ItemAtt("Item_115")]
        public SAPbouiCOM.EditText edtAracTankTkm;

        [ItemAtt("Item_117")]
        public SAPbouiCOM.EditText edtAracTankSomatikHucre;

        [ItemAtt("Item_119")]
        public SAPbouiCOM.EditText edtAracTankMetilenMavisi;

        [ItemAtt("Item_156")]
        public SAPbouiCOM.ComboBox oComboAracTankTemizlikDurum;

        [ItemAtt("Item_126")]
        public SAPbouiCOM.EditText edtAracTankRow;


        [ItemAtt("Item_122")]
        public SAPbouiCOM.Button btnAracTankKaydet;

        #endregion

        [ItemAtt("Item_188")]
        public SAPbouiCOM.Button btnKantarVerisiniManuelGir;




        //[ItemAtt("Item_157")]
        //public SAPbouiCOM.ComboBox oComboTemiz;

        #endregion Değişkenler
        //string mKod = System.Configuration.ConfigurationManager.AppSettings["MusteriKodu"];

        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.SutKabulUID, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.SutKabulFrmXML));
            Functions.CreateUserOrSystemFormComponent<SutKabul>(AIFConn.SutKabul);

            InitForms();
        }
        private bool kantarVerisiManuelGiris = false;
        public void InitForms()
        {
            try
            {
                frmSutKabul.EnableMenu("1283", false);
                frmSutKabul.EnableMenu("1284", false);
                frmSutKabul.EnableMenu("1286", false);

                KantarVerisiOk = false;
                oFolderAracBilgileri.Select();
                oMatrixAnalizBilgileri.AutoResizeColumns();
                oMatrixAracTankVerileri.AutoResizeColumns();
                //frmSutKabul.DataSources.UserDataSources.Item("UD_9").ValueEx = "0";
                oComboBelgeDurumu.Select("O");
                oEditBelgeTarihi.String = "A";
                oEditDepo.Value = "200";
                //frmSutKabul.DataSources.UserDataSources.Item("UD_9").Value = "1";


                //frmSutKabul.DataSources.UserDataSources.Item("UD_10").Value = "1";
                //frmSutKabul.DataSources.UserDataSources.Item("UD_11").Value = "2";
                //frmSutKabul.DataSources.UserDataSources.Item("UD_12").Value = "2";
                //frmSutKabul.DataSources.UserDataSources.Item("UD_47").Value = "2";

                //frmSutKabul.DataSources.UserDataSources.Item("UD_43").Value = "1";
                //frmSutKabul.DataSources.UserDataSources.Item("UD_39").Value = "1";
                //frmSutKabul.DataSources.UserDataSources.Item("UD_31").Value = "2";
                //frmSutKabul.DataSources.UserDataSources.Item("UD_32").Value = "2";

                #region kantar verisini manuel gimek için süper kullanıcıya bağlandı  
                string user = ConstVariables.oCompanyObject.UserName;
                ConstVariables.oRecordset.DoQuery("SELECT \"SUPERUSER\" FROM \"OUSR\" where \"USER_CODE\" = '" + user + "' and \"SUPERUSER\" = 'Y'");

                if (ConstVariables.oRecordset.RecordCount > 0)
                {
                    btnKantarVerisiniManuelGir.Item.Visible = true;
                    kantarVerisiManuelGiris = true;
                }
                #endregion

                ((SAPbouiCOM.ComboBox)frmSutKabul.Items.Item("Item_137").Specific).Select("1", BoSearchKey.psk_ByValue);
                string Sutturu = "1";
                SutTuruneGoreListDoldurma(Sutturu);

                #region eski yanlış 19.10.2022
                //ConstVariables.oRecordset.DoQuery("Select FldValue,Descr from UFD1 where FieldID =(select FieldID from CUFD where TableID = '@AIF_SUTKABUL2'and AliasID = 'Nitelik')"); 
                #endregion

                ConstVariables.oRecordset.DoQuery("Select FldValue,Descr from UFD1 where TableID = '@AIF_SUTKABUL2' and FieldID = (select FieldID from CUFD where TableID = '@AIF_SUTKABUL2' and AliasID = 'Nitelik')");

                while (!ConstVariables.oRecordset.EoF)
                {
                    oComboNitelik.ValidValues.Add(ConstVariables.oRecordset.Fields.Item(0).Value.ToString(), ConstVariables.oRecordset.Fields.Item(1).Value.ToString());
                    ConstVariables.oRecordset.MoveNext();
                }
            }
            catch (Exception ex)
            {
            }
        }

        private List<Models.SutKabulUyarlama> SutKabulUyarlama = new List<Models.SutKabulUyarlama>();
        private bool KantarVerisiOk = false;

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
                        string msg = "";

                        ConstVariables.oRecordset.DoQuery("Select \"AutoKey\" from ONNM where \"ObjectCode\" = 'AIF_SUTKABUL'");

                        if (ConstVariables.oRecordset.RecordCount > 0)
                        {
                            msg = Environment.NewLine;
                            msg += "İşlem zamanı:" + DateTime.Now + "--" + "süt kabul belgesi eklenmeye tıklandı ve SAP sıra numarası " + ConstVariables.oRecordset.Fields.Item(0).Value.ToString();
                            AIFConn.SutKabul.WriteToFile(msg);
                        }

                        string KayitNo = frmSutKabul.DataSources.DBDataSources.Item(2).GetValue("U_KayitNo", frmSutKabul.DataSources.DBDataSources.Item(2).Size - 1).ToString();
                        string Tedarikci = frmSutKabul.DataSources.DBDataSources.Item(2).GetValue("U_Tedarikci", frmSutKabul.DataSources.DBDataSources.Item(2).Size - 1).ToString();

                        if (KayitNo == "" && Tedarikci == "")
                        {
                            frmSutKabul.DataSources.DBDataSources.Item(2).RemoveRecord(frmSutKabul.DataSources.DBDataSources.Item(2).Size - 1);
                        }

                        KayitNo = frmSutKabul.DataSources.DBDataSources.Item(3).GetValue("U_KayitNo", frmSutKabul.DataSources.DBDataSources.Item(3).Size - 1).ToString();
                        string TankNumarasi = frmSutKabul.DataSources.DBDataSources.Item(3).GetValue("U_TankNumarasi", frmSutKabul.DataSources.DBDataSources.Item(3).Size - 1).ToString();

                        if (KayitNo == "" && TankNumarasi == "")
                        {
                            frmSutKabul.DataSources.DBDataSources.Item(3).RemoveRecord(frmSutKabul.DataSources.DBDataSources.Item(3).Size - 1);
                        }
                        ekleme = true;
                    }
                    break;

                case BoEventTypes.et_FORM_DATA_UPDATE:
                    break;

                case BoEventTypes.et_FORM_DATA_DELETE:
                    break;

                case BoEventTypes.et_FORM_DATA_LOAD:
                    if (!BusinessObjectInfo.BeforeAction)
                    {
                        KantarVerisiOk = true;

                        string val = oComboBelgeDurumu.Value.Trim();

                        if (val == "C")
                        {
                            statuKapaliIslemleri();
                        }
                        else
                        {
                            statusAcikOnaylandi();
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

        public string NoktalamaIsaretiDegistir(string val1)
        {
            string Val = Program.SAPnfi == "," ? val1.ToString().Replace(".", ",") : val1.ToString();

            return Val;
        }

        private bool ekleme = false;

        public bool SAP_ItemEvent(string FormUID, ref ItemEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;

            switch (pVal.EventType)
            {
                case BoEventTypes.et_ALL_EVENTS:
                    break;

                case BoEventTypes.et_ITEM_PRESSED:
                    if (pVal.ItemUID == "Item_27" && !pVal.BeforeAction)
                    {
                        frmSutKabul.DataSources.DBDataSources.Item(1).SetValue("U_BelgeTipineGore", 0, "N");
                        frmSutKabul.DataSources.DBDataSources.Item(1).SetValue("U_SayacaGore", 0, "N");
                    }
                    else if (pVal.ItemUID == "Item_28" && !pVal.BeforeAction)
                    {
                        frmSutKabul.DataSources.DBDataSources.Item(1).SetValue("U_KantaraGore", 0, "N");
                        frmSutKabul.DataSources.DBDataSources.Item(1).SetValue("U_SayacaGore", 0, "N");
                    }
                    else if (pVal.ItemUID == "Item_29" && !pVal.BeforeAction)
                    {
                        frmSutKabul.DataSources.DBDataSources.Item(1).SetValue("U_KantaraGore", 0, "N");
                        frmSutKabul.DataSources.DBDataSources.Item(1).SetValue("U_BelgeTipineGore", 0, "N");
                    }
                    else if (pVal.ItemUID == "Item_145" && !pVal.BeforeAction)
                    {
                        frmSutKabul.DataSources.DBDataSources.Item(1).SetValue("U_BelgesizGiris", 0, "N");
                    }
                    else if (pVal.ItemUID == "Item_146" && !pVal.BeforeAction)
                    {
                        frmSutKabul.DataSources.DBDataSources.Item(1).SetValue("U_IrsaliyeliGiris", 0, "N");
                    }
                    else if (pVal.ItemUID == "Item_49" && !pVal.BeforeAction)
                    {
                        frmSutKabul.DataSources.DBDataSources.Item(1).SetValue("U_TekilGiris", 0, "N");
                        oEditSatici.Value = "";
                        oEditKantarNo.Item.Click();
                        oEditSatici.Item.Enabled = false;
                        oEditOlusanBelgeNo.Value = "";
                        oEditKantarNo.Item.Click();
                        oEditOlusanBelgeNo.Item.Enabled = false;
                        string msg = "";
                        if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                        {
                            msg = Environment.NewLine;
                            msg += "İşlem zamanı:" + DateTime.Now + "--" + " çoklu giriş checkbox tıklandı.";
                            AIFConn.SutKabul.WriteToFile(msg);
                        }
                        else
                        {
                            msg = Environment.NewLine;
                            msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " çoklu giriş checkbox tıklandı.";
                            AIFConn.SutKabul.WriteToFile(msg);
                        }
                    }
                    else if (pVal.ItemUID == "Item_26" && !pVal.BeforeAction)
                    {
                        string msg = "";
                        if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                        {
                            msg = Environment.NewLine;
                            msg += "İşlem zamanı:" + DateTime.Now + "--" + " tekil giriş checkbox tıklandı.";
                            AIFConn.SutKabul.WriteToFile(msg);
                        }
                        else
                        {
                            msg = Environment.NewLine;
                            msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " tekil giriş checkbox tıklandı.";
                            AIFConn.SutKabul.WriteToFile(msg);
                        }

                        oEditSatici.Item.Enabled = true;
                        frmSutKabul.DataSources.DBDataSources.Item(1).SetValue("U_CokluGiris", 0, "N");
                    }
                    else if (pVal.ItemUID == "1" && ekleme)
                    {
                        KantarVerisiOk = false;
                        ekleme = false;
                        oComboBelgeDurumu.Select("O");
                        oEditBelgeTarihi.String = "A";
                        oEditDepo.Value = "200";
                        AnalizTemizle();
                        AracTankTemizle();

                        string msg = "";

                        ConstVariables.oRecordset.DoQuery("Select \"DocEntry\" from \"@AIF_SUTKABUL\"  order by \"DocEntry\" desc");

                        if (ConstVariables.oRecordset.RecordCount > 0)
                        {
                            msg = Environment.NewLine;
                            msg += "İşlem zamanı:" + DateTime.Now + "--" + "süt kabul belgesi eklenmeye tıklandıktan sonraki süt kabuldeki en son duran belge numarası " + ConstVariables.oRecordset.Fields.Item(0).Value.ToString();
                            AIFConn.SutKabul.WriteToFile(msg);
                        }
                    }
                    else if (pVal.ItemUID == "1" && !pVal.BeforeAction)
                    {
                        try
                        {
                            edtAracPlakasi.Item.Click();
                            oEditDocEntry.Item.Enabled = false;
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
                    if (pVal.ItemUID == "Item_170" || pVal.ItemUID == "Item_172" || pVal.ItemUID == "Item_174" || pVal.ItemUID == "Item_176")
                    {
                        if (pVal.ItemUID == "Item_170")
                        {
                            if (parseNumber.parservalues<double>(edtSutTemizligi.Value) < 0)
                            {
                                frmSutKabul.DataSources.UserDataSources.Item("UD_61").Value = "0";
                                Handler.SAPApplication.MessageBox("Analiz bilgileri süt temizliği alanı 0'dan küçük olamaz.");
                                return false;
                            }
                            else if (parseNumber.parservalues<double>(edtSutTemizligi.Value) > 100)
                            {
                                frmSutKabul.DataSources.UserDataSources.Item("UD_61").Value = "0";
                                Handler.SAPApplication.MessageBox("Analiz bilgileri süt temizliği alanı 100'den büyük olamaz.");
                                return false;
                            }
                        }

                        if (pVal.ItemUID == "Item_172")
                        {
                            if (parseNumber.parservalues<double>(edtTankTemizligi.Value) < 0)
                            {
                                frmSutKabul.DataSources.UserDataSources.Item("UD_62").Value = "0";
                                Handler.SAPApplication.MessageBox("Analiz bilgileri tank temizliği alanı 0'dan küçük olamaz.");
                                return false;
                            }
                            else if (parseNumber.parservalues<double>(edtSutTemizligi.Value) > 100)
                            {
                                frmSutKabul.DataSources.UserDataSources.Item("UD_62").Value = "0";
                                Handler.SAPApplication.MessageBox("Analiz bilgileri tank temizliği alanı 100'den büyük olamaz.");
                                return false;
                            }
                        }

                        if (pVal.ItemUID == "Item_174")
                        {
                            if (parseNumber.parservalues<double>(edtEkipmanTemizligi.Value) < 0)
                            {
                                frmSutKabul.DataSources.UserDataSources.Item("UD_63").Value = "0";
                                Handler.SAPApplication.MessageBox("Analiz bilgileri ekipman temizliği alanı 0'dan küçük olamaz.");
                                return false;
                            }
                            else if (parseNumber.parservalues<double>(edtEkipmanTemizligi.Value) > 100)
                            {
                                frmSutKabul.DataSources.UserDataSources.Item("UD_63").Value = "0";
                                Handler.SAPApplication.MessageBox("Analiz bilgileri ekipman temizliği alanı 100'den büyük olamaz.");
                                return false;
                            }
                        }

                        if (pVal.ItemUID == "Item_176")
                        {
                            if (parseNumber.parservalues<double>(edtAracKasasiTemizligi.Value) < 0)
                            {
                                frmSutKabul.DataSources.UserDataSources.Item("UD_64").Value = "0";
                                Handler.SAPApplication.MessageBox("Analiz bilgileri araç kasası temizliği alanı 0'dan küçük olamaz.");
                                return false;
                            }
                            else if (parseNumber.parservalues<double>(edtEkipmanTemizligi.Value) > 100)
                            {
                                frmSutKabul.DataSources.UserDataSources.Item("UD_64").Value = "0";
                                Handler.SAPApplication.MessageBox("Analiz bilgileri araç kasası temizliği alanı 100'den büyük olamaz.");
                                return false;
                            }
                        }

                        AgirlikliPuanHesapla();
                    }
                    break;

                case BoEventTypes.et_COMBO_SELECT:
                    if (pVal.ItemUID == "Item_137" && !pVal.BeforeAction)
                    {
                        AnalizTemizle(true);

                        string Sutturu = oComboAnalizSutTuru.Value.Trim();

                        SutTuruneGoreListDoldurma(Sutturu);

                    }
                    else if (pVal.ItemUID == "Item_166" && !pVal.BeforeAction)
                    {
                        if (oComboBelgeDurumu.Value.Trim() == "C")
                        {
                            if (!SutDepoSecim.sutkabultamamlaniyor)
                            {
                                frmSutKabul.DataSources.DBDataSources.Item("@AIF_SUTKABUL").SetValue("U_BelgeDurumu", 0, comboBelgeDurumuTiklanmaDegeri);
                                Handler.SAPApplication.MessageBox("Kapalı statüsü yalnızca Süt kabul işlemi tamamlandığında otomatik olarak seçilebilir.");
                                return false;
                            }
                        }

                        string msg = "";
                        if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                        {
                            msg = Environment.NewLine;
                            msg += "İşlem zamanı:" + DateTime.Now + "--" + " Belge durumu değiştirildi. Seçilen Değer " + oComboBelgeDurumu.Value.Trim();
                            AIFConn.SutKabul.WriteToFile(msg);
                        }
                        else
                        {
                            msg = Environment.NewLine;
                            msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " Belge durumu değiştirildi. Seçilen Değer " + oComboBelgeDurumu.Value.Trim();
                            AIFConn.SutKabul.WriteToFile(msg);
                        }
                    }
                    break;

                case BoEventTypes.et_CLICK:
                    if (pVal.ItemUID == "Item_74" && pVal.BeforeAction)
                    {
                        string msg = "";

                        if (btnKaydet.Item.Enabled == false)
                        {
                            BubbleEvent = false;
                            return false;
                        }

                        var row = edtAnalizRow.Value.ToString();

                        if (row == "")
                        {
                            //if (frmSutKabul.Mode != BoFormMode.fm_ADD_MODE)
                            if (oComboBelgeDurumu.Value.Trim() == "C" || oComboBelgeDurumu.Value.Trim() == "A")
                            {
                                if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                                {
                                    msg = Environment.NewLine;
                                    msg += "İşlem zamanı:" + DateTime.Now + "--" + " Güncelleme modunda satır ekleme işlemi yapılamaz hatası alındı.";
                                    AIFConn.SutKabul.WriteToFile(msg);
                                }
                                else
                                {
                                    msg = Environment.NewLine;
                                    msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " Güncelleme modunda satır ekleme işlemi yapılamaz hatası alındı.";
                                    AIFConn.SutKabul.WriteToFile(msg);
                                }
                                Handler.SAPApplication.MessageBox("Belge durumu kapalı ve Onaylandı olduğunda satır ekleme işlemi yapılamaz.");

                                BubbleEvent = false;
                                return false;
                            }
                        }

                        if (oComboBelgeDurumu.Value.Trim() == "C" || oComboBelgeDurumu.Value.Trim() == "A")
                        {
                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + " Durumu açık olmayan belge için işlem yapılamaz hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " Durumu açık olmayan belge için işlem yapılamaz hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            Handler.SAPApplication.MessageBox("Durumu açık olmayan belge için işlem yapılamaz.");
                            BubbleEvent = false;
                            return false;
                        }

                        string sutTuru = oComboAnalizSutTuru.Value.Trim();
                        if (sutTuru == "0")
                        {
                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + " Süt Türü seçilmeden işleme devam edilemez hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " Süt Türü seçilmeden işleme devam edilemez hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            Handler.SAPApplication.MessageBox("Süt Türü seçilmeden işleme devam edilemez.");
                            BubbleEvent = false;
                            return false;
                        }

                        if (edtAnalizTedarikci.Value == "")
                        {
                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + " Tedarikçi seçimi yapılmadan işleme devam edilemez hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " Tedarikçi seçimi yapılmadan işleme devam edilemez hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            Handler.SAPApplication.MessageBox("Tedarikçi seçimi yapılmadan işleme devam edilemez.");
                            BubbleEvent = false;
                            return false;
                        }

                        double sutMiktar = edtAnalizMiktar.Value.ToString() == "" ? 0 : Convert.ToDouble(edtAnalizMiktar.Value);
                        if (sutMiktar == 0)
                        {
                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + " Süt Miktarı girilmeden işleme devam edilemez hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " Süt Miktarı girilmeden işleme devam edilemez hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            Handler.SAPApplication.MessageBox("Süt Miktarı girilmeden işleme devam edilemez.");
                            BubbleEvent = false;
                            return false;
                        }

                        bool check = AnalizBilgileriAlanlariniKontrolEt();

                        if (!check)
                        {
                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + " AnalizBilgileriAlanlariniKontrolEt methounda hata alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " AnalizBilgileriAlanlariniKontrolEt methounda hata alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            BubbleEvent = false;
                            return false;
                        }
                    }
                    else if (pVal.ItemUID == "Item_74" && !pVal.BeforeAction)
                    {
                        try
                        {
                            string msg = "";
                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + " analiz bilgilerini kaydetme işlemi yapıldı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " analiz bilgilerini kaydetme işlemi yapıldı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }

                            frmSutKabul.Freeze(true);
                            int row = edtAnalizRow.Value.ToString() == "" ? 0 : Convert.ToInt32(edtAnalizRow.Value);

                            if (row == 0)
                            {
                                frmSutKabul.DataSources.DBDataSources.Item(2).Clear();
                                oMatrixAnalizBilgileri.AddRow();

                                if (oMatrixAnalizBilgileri.RowCount == 0)
                                    row = 1;
                                else
                                    row = oMatrixAnalizBilgileri.RowCount;
                            }

                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_1", edtAnalizTedarikci.Value);
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_22", edtAnalizTedarikciAdi.Value.ToString());
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_2", edtAnalizMiktar.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_3", edtAnalizSicaklik.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_4", edtAnalizBriks.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_5", edtAnalizYag.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_6", edtAnalizYogunluk.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_7", edtAnalizpH.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_8", edtAnalizSH.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_9", oComboAnalizSutTuru.Value.ToString());
                            //oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_10", oComboAnalizTemizkikKoku.Value.Trim());
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_11", oComboAnalizSoda.Value.Trim());
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_12", oComboAnalizAntibiyotik.Value.Trim());
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_13", edtAnalizYkm.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_14", oComboAnalizAlkol.Value.Trim());
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_15", edtAnalizKaynatSh.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_16", edtAnalizDepo.Value.ToString());
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_17", edtAnalizProtein.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_18", edtAnalizLaktoz.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_19", edtAnalizKatilanSu.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_20", edtAnalizSomatikHucre.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_21", edtAnalizMetilenMavisi.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_24", edtAnalizDonmaNoktasi.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_25", edtAnalizTKM.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_26", oComboAnalizKoku.Value.ToString().Trim());
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_27", oComboNitelik.Value.ToString().Trim());
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_28", edtSutTemizligi.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_29", edtTankTemizligi.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_30", edtEkipmanTemizligi.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_31", edtAracKasasiTemizligi.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_32", edtAgirlikliPuan.Value.ToString(Program.nfi));
                            oMatrixAnalizBilgileri.SetCellWithoutValidation(row, "Col_10", edtIrsaliyeNo.Value.ToString());

                            oMatrixAnalizBilgileri.AutoResizeColumns();

                            oMatrixAnalizBilgileri.Columns.Item("#").TitleObject.Click();

                            AnalizTemizle();

                            string xml = oMatrixAnalizBilgileri.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                            var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                        select new AnalizBilgileri
                                        {
                                            Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value),
                                        }).ToList();

                            double total = rows.Sum(x => x.Miktar);

                            oedtIrsaliyeNetSutMiktari.Value = total.ToString();

                            oEditDepo.Value = "200";
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmSutKabul.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "Item_75" && !pVal.BeforeAction)
                    {
                        AnalizTemizle();
                    }
                    else if (pVal.ItemUID == "Item_122" && !pVal.BeforeAction)
                    {
                        try
                        {
                            string msg = "";
                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + " araç tank verilerini kaydetme işlemi yapıldı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " araç tank verilerini kaydetme işlemi yapıldı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }

                            frmSutKabul.Freeze(true);
                            int row = edtAracTankRow.Value.ToString() == "" ? 0 : Convert.ToInt32(edtAracTankRow.Value);

                            if (row == 0)
                            {
                                oMatrixAracTankVerileri.AddRow();
                                row = oMatrixAracTankVerileri.RowCount;
                            }

                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_1", edtAracTankTankNumarasi.Value);
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_12", edtAracTedarikci.Value);
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_23", edtAracTedarikciAdi.Value);
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_2", edtAracTankLabMiktar.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_3", edtAracTankDerece.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_4", edtAracTankBrix.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_5", edtAracTankYag.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_6", edtAracTankDans.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_7", edtAracTankpH.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_8", edtAracTankSH.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_9", edtAracTankKaynamisSH.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_10", oComboAracTankAntibiyotik.Value.Trim());
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_11", oComboAracTankAlkol.Value.Trim());
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_13", edtAracTankYagsiz.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_14", edtAracTankProtein.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_15", edtAracTankDonNok.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_16", edtAracTankLaktoz.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_17", edtAracTankSu.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_18", oComboAracTankKoku.Value.Trim());
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_19", edtAracTankTkm.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_20", edtAracTankSomatikHucre.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_21", edtAracTankMetilenMavisi.Value.ToString(Program.nfi));
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_22", oComboAracTankTemizlikDurum.Value.Trim());
                            oMatrixAracTankVerileri.SetCellWithoutValidation(row, "Col_24", oComboAracTankSoda.Value.Trim());

                            oMatrixAracTankVerileri.AutoResizeColumns();

                            oMatrixAracTankVerileri.Columns.Item("#").TitleObject.Click();

                            AracTankTemizle();

                            string xml = oMatrixAracTankVerileri.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                            var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                        select new AnalizBilgileri
                                        {
                                            Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value),
                                        }).ToList();

                            double total = rows.Sum(x => x.Miktar);

                            oEditSayacBilgileriToplam.Value = total.ToString();
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmSutKabul.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "Item_122" && pVal.BeforeAction)
                    {
                        string msg = "";

                        string aracTankNumarasi = edtAracTankTankNumarasi.Value.ToString();
                        if (aracTankNumarasi == "")
                        {
                            Handler.SAPApplication.MessageBox("Araç tank numarası numarası girilmeden işleme devam edilemez.");
                            BubbleEvent = false;
                            return false;
                        }

                        if (edtAracTedarikci.Value == "")
                        {
                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + " Tedarikçi seçimi yapılmadan işleme devam edilemez hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " Tedarikçi seçimi yapılmadan işleme devam edilemez hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }

                            Handler.SAPApplication.MessageBox("Tedarikçi seçimi yapılmadan işleme devam edilemez.");
                            BubbleEvent = false;
                            return false;
                        }

                        double LabMiktar = edtAracTankLabMiktar.Value.ToString() == "" ? 0 : Convert.ToDouble(edtAracTankLabMiktar.Value);
                        if (LabMiktar == 0)
                        {
                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + " Lab miktarı girilmeden işleme devam edilemez hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " Lab miktarı girilmeden işleme devam edilemez alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            Handler.SAPApplication.MessageBox("Lab miktarı girilmeden işleme devam edilemez.");
                            BubbleEvent = false;
                            return false;
                        }

                        bool check = TankVerileriAlanlariniKontroEt();

                        if (!check)
                        {
                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + " TankVerileriAlanlariniKontroEt methodunda hata alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " TankVerileriAlanlariniKontroEt methodunda hata alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            BubbleEvent = false;
                            return false;
                        }
                    }
                    else if (pVal.ItemUID == "Item_123" && !pVal.BeforeAction)
                    {
                        AracTankTemizle();
                    }
                    else if (pVal.ColUID == "#" && pVal.ItemUID == "Item_77" && !pVal.BeforeAction)
                    {
                        try
                        {
                            frmSutKabul.Freeze(true);
                            string xml = oMatrixAnalizBilgileri.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                            var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                        select new AnalizBilgileri
                                        {
                                            Tedarikci = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                                            TedarikciAdi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_22" select new XElement(y.Element("Value"))).First().Value,
                                            Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value),
                                            Sicaklik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value),
                                            Briks = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value),
                                            Yag = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value),
                                            Yogunluk = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value),
                                            pH = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value),
                                            SH = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value),
                                            SutTuru = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_9" select new XElement(y.Element("Value"))).First().Value,
                                            //TemizlikKoku = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_10" select new XElement(y.Element("Value"))).First().Value,
                                            Soda = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_11" select new XElement(y.Element("Value"))).First().Value,
                                            Antibiyotik = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_12" select new XElement(y.Element("Value"))).First().Value,
                                            Ykm = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_13" select new XElement(y.Element("Value"))).First().Value),
                                            Alkol = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_14" select new XElement(y.Element("Value"))).First().Value,
                                            KaynatSh = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_15" select new XElement(y.Element("Value"))).First().Value),
                                            Depo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_16" select new XElement(y.Element("Value"))).First().Value,
                                            Protein = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_17" select new XElement(y.Element("Value"))).First().Value),
                                            Laktoz = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_18" select new XElement(y.Element("Value"))).First().Value),
                                            KatilanSu = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_19" select new XElement(y.Element("Value"))).First().Value),
                                            SomatikHucre = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_20" select new XElement(y.Element("Value"))).First().Value),
                                            MetilenMavisi = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_21" select new XElement(y.Element("Value"))).First().Value),
                                            Koku = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_26" select new XElement(y.Element("Value"))).First().Value,
                                            DonmaNoktasi = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_24" select new XElement(y.Element("Value"))).First().Value),
                                            TKM = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_25" select new XElement(y.Element("Value"))).First().Value),
                                            Nitelik = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_27" select new XElement(y.Element("Value"))).First().Value,
                                            row = Convert.ToString(x.ElementsBeforeSelf().Count() + 1),
                                            SutTemizlik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_28" select new XElement(y.Element("Value"))).First().Value),
                                            TankTemizlik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_29" select new XElement(y.Element("Value"))).First().Value),
                                            EkipmanTemizlik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_30" select new XElement(y.Element("Value"))).First().Value),
                                            AracKasasiTemizlik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_31" select new XElement(y.Element("Value"))).First().Value),
                                            AgirlikliPuan = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_32" select new XElement(y.Element("Value"))).First().Value),
                                            IrsaliyeNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_10" select new XElement(y.Element("Value"))).First().Value,
                                        }).ToList();

                            string pvalrow = pVal.Row.ToString();
                            foreach (var item in rows.Where(x => x.row == pvalrow))
                            {
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_1").Value = item.Miktar.ToString(Program.nfi);
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_2").Value = item.Tedarikci.ToString();
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_46").Value = item.TedarikciAdi.ToString();
                                //edtAnalizSicaklik.Value = item.Sicaklik.ToString(Program.nfi);
                                //edtAnalizBriks.Value = item.Briks.ToString(Program.nfi);
                                //edtAnalizYag.Value = item.Yag.ToString(Program.nfi);
                                //edtAnalizYogunluk.Value = item.Yogunluk.ToString(Program.nfi);
                                //edtAnalizpH.Value = item.pH.ToString(Program.nfi);
                                //edtAnalizSH.Value = item.SH.ToString(Program.nfi);
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_9").Value = item.SutTuru.ToString();
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_10").Value = item.TemizlikKoku.ToString();
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_11").Value = item.Soda.ToString();
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_12").Value = item.Antibiyotik.ToString();
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_13").Value = item.Ykm.ToString();
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_47").Value = item.Alkol.ToString();
                                //edtAnalizKaynatSh.Value = item.KaynatSh.ToString(Program.nfi);
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_16").Value = item.Depo.ToString();
                                //edtAnalizProtein.Value = item.Protein.ToString(Program.nfi);
                                //edtAnalizLaktoz.Value = item.Laktoz.ToString(Program.nfi);
                                //edtAnalizKatilanSu.Value = item.KatilanSu.ToString(Program.nfi);
                                //edtAnalizSomatikHucre.Value = item.SomatikHucre.ToString(Program.nfi);
                                //edtAnalizMetilenMavisi.Value = item.MetilenMavisi.ToString(Program.nfi);
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_44").Value = pVal.Row.ToString();

                                frmSutKabul.DataSources.UserDataSources.Item("UD_1").Value = item.Miktar.ToString(Program.nfi);
                                frmSutKabul.DataSources.UserDataSources.Item("UD_2").Value = item.Tedarikci.ToString();
                                frmSutKabul.DataSources.UserDataSources.Item("UD_46").Value = item.TedarikciAdi.ToString(); frmSutKabul.DataSources.UserDataSources.Item("UD_3").Value = NoktalamaIsaretiDegistir(item.Sicaklik.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_4").Value = NoktalamaIsaretiDegistir(item.Briks.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_5").Value = NoktalamaIsaretiDegistir(item.Yag.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_6").Value = NoktalamaIsaretiDegistir(item.Yogunluk.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_7").Value = NoktalamaIsaretiDegistir(item.pH.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_8").Value = NoktalamaIsaretiDegistir(item.SH.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_9").Value = item.SutTuru.ToString();
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_10").Value = item.TemizlikKoku.ToString();
                                frmSutKabul.DataSources.UserDataSources.Item("UD_11").Value = item.Soda.ToString();
                                frmSutKabul.DataSources.UserDataSources.Item("UD_12").Value = item.Antibiyotik.ToString();
                                frmSutKabul.DataSources.UserDataSources.Item("UD_13").Value = item.Ykm.ToString();
                                frmSutKabul.DataSources.UserDataSources.Item("UD_47").Value = item.Alkol.ToString();

                                frmSutKabul.DataSources.UserDataSources.Item("UD_15").Value = NoktalamaIsaretiDegistir(item.KaynatSh.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_13").Value = NoktalamaIsaretiDegistir(item.Ykm.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_17").Value = NoktalamaIsaretiDegistir(item.Protein.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_18").Value = NoktalamaIsaretiDegistir(item.Laktoz.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_19").Value = NoktalamaIsaretiDegistir(item.KatilanSu.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_20").Value = NoktalamaIsaretiDegistir(item.SomatikHucre.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_21").Value = NoktalamaIsaretiDegistir(item.MetilenMavisi.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_44").Value = pVal.Row.ToString();
                                frmSutKabul.DataSources.UserDataSources.Item("UD_56").Value = item.Koku;
                                frmSutKabul.DataSources.UserDataSources.Item("UD_57").Value = NoktalamaIsaretiDegistir(item.DonmaNoktasi.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_58").Value = NoktalamaIsaretiDegistir(item.TKM.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_60").Value = item.Nitelik;
                                frmSutKabul.DataSources.UserDataSources.Item("UD_61").Value = NoktalamaIsaretiDegistir(item.SutTemizlik.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_62").Value = NoktalamaIsaretiDegistir(item.TankTemizlik.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_63").Value = NoktalamaIsaretiDegistir(item.EkipmanTemizlik.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_64").Value = NoktalamaIsaretiDegistir(item.AracKasasiTemizlik.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_65").Value = NoktalamaIsaretiDegistir(item.AgirlikliPuan.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_66").Value = item.IrsaliyeNo;
                            }

                            if (pvalrow == "0")
                            {
                                AnalizTemizle();
                            }
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmSutKabul.Freeze(false);
                        }
                    }
                    else if (pVal.ColUID == "#" && pVal.ItemUID == "Item_124" && !pVal.BeforeAction)
                    {
                        try
                        {
                            frmSutKabul.Freeze(true);
                            string xml = oMatrixAracTankVerileri.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                            var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                        select new AracTankVerileri
                                        {
                                            TankNumarasi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
                                            Tedarikci = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_12" select new XElement(y.Element("Value"))).First().Value,
                                            TedarikciAdi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_23" select new XElement(y.Element("Value"))).First().Value,
                                            LabMiktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value),
                                            Derece = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value),
                                            Brix = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value),
                                            Yag = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value),
                                            Dans = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value),
                                            pH = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value),
                                            SH = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value),
                                            KaynamisSH = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_9" select new XElement(y.Element("Value"))).First().Value),
                                            Antibiyotik = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_10" select new XElement(y.Element("Value"))).First().Value,
                                            Alkol = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_11" select new XElement(y.Element("Value"))).First().Value,
                                            YagiszKm = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_13" select new XElement(y.Element("Value"))).First().Value),
                                            Protein = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_14" select new XElement(y.Element("Value"))).First().Value),
                                            DON_NOK = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_15" select new XElement(y.Element("Value"))).First().Value),
                                            Laktoz = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_16" select new XElement(y.Element("Value"))).First().Value),
                                            Su = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_17" select new XElement(y.Element("Value"))).First().Value),
                                            Koku = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_18" select new XElement(y.Element("Value"))).First().Value,
                                            Tkm = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_19" select new XElement(y.Element("Value"))).First().Value),
                                            SomatikHucre = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_20" select new XElement(y.Element("Value"))).First().Value),
                                            MetilenMavisi = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_21" select new XElement(y.Element("Value"))).First().Value),
                                            TemizlikDurum = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_22" select new XElement(y.Element("Value"))).First().Value,
                                            row = Convert.ToString(x.ElementsBeforeSelf().Count() + 1),
                                            Soda = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_24" select new XElement(y.Element("Value"))).First().Value,
                                        }).ToList();

                            string pvalrow = pVal.Row.ToString();
                            foreach (var item in rows.Where(x => x.row == pvalrow))
                            {
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_22").Value = item.TankNumarasi;
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_53").Value = item.Tedarikci;
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_54").Value = item.TedarikciAdi;
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_23").Value = item.LabMiktar.ToString(Program.nfi);
                                //edtAracTankDerece.Value = item.Derece.ToString(Program.nfi);
                                //edtAracTankBrix.Value = item.Brix.ToString(Program.nfi);
                                //edtAracTankYag.Value = item.Yag.ToString(Program.nfi);
                                //edtAracTankDans.Value = item.Dans.ToString(Program.nfi);
                                //edtAracTankpH.Value = item.pH.ToString(Program.nfi);
                                //edtAracTankSH.Value = item.SH.ToString(Program.nfi);
                                //edtAracTankKaynamisSH.Value = item.KaynamisSH.ToString();
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_31").Value = item.Antibiyotik.ToString();
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_32").Value = item.Alkol.ToString();
                                ////frmSutKabul.DataSources.UserDataSources.Item("UD_33").Value = item.Yag1.ToString();
                                //edtAracTankYagsiz.Value = item.YagiszKm.ToString(Program.nfi);
                                //edtAracTankProtein.Value = item.Protein.ToString(Program.nfi);
                                //edtAracTankDonNok.Value = item.DON_NOK.ToString(Program.nfi);
                                //edtAracTankLaktoz.Value = item.Laktoz.ToString(Program.nfi);
                                //edtAracTankSu.Value = item.Su.ToString(Program.nfi);
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_39").Value = item.Koku.ToString();
                                //edtAracTankTkm.Value = item.Tkm.ToString(Program.nfi);
                                //edtAracTankSomatikHucre.Value = item.SomatikHucre.ToString(Program.nfi);
                                //edtAracTankMetilenMavisi.Value = item.MetilenMavisi.ToString(Program.nfi);
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_43").Value = item.TemizlikDurum;
                                //frmSutKabul.DataSources.UserDataSources.Item("UD_45").Value = pVal.Row.ToString();

                                frmSutKabul.DataSources.UserDataSources.Item("UD_22").Value = item.TankNumarasi;
                                frmSutKabul.DataSources.UserDataSources.Item("UD_53").Value = item.Tedarikci;
                                frmSutKabul.DataSources.UserDataSources.Item("UD_54").Value = item.TedarikciAdi;
                                frmSutKabul.DataSources.UserDataSources.Item("UD_23").Value = NoktalamaIsaretiDegistir(item.LabMiktar.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_24").Value = NoktalamaIsaretiDegistir(item.Derece.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_25").Value = NoktalamaIsaretiDegistir(item.Brix.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_26").Value = NoktalamaIsaretiDegistir(item.Yag.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_27").Value = NoktalamaIsaretiDegistir(item.Dans.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_28").Value = NoktalamaIsaretiDegistir(item.pH.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_29").Value = NoktalamaIsaretiDegistir(item.SH.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_30").Value = NoktalamaIsaretiDegistir(item.KaynamisSH.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_31").Value = item.Antibiyotik;
                                frmSutKabul.DataSources.UserDataSources.Item("UD_32").Value = item.Alkol;
                                frmSutKabul.DataSources.UserDataSources.Item("UD_43").Value = item.TemizlikDurum;
                                frmSutKabul.DataSources.UserDataSources.Item("UD_39").Value = item.Koku;
                                frmSutKabul.DataSources.UserDataSources.Item("UD_34").Value = NoktalamaIsaretiDegistir(item.YagiszKm.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_35").Value = NoktalamaIsaretiDegistir(item.Protein.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_37").Value = NoktalamaIsaretiDegistir(item.Laktoz.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_38").Value = NoktalamaIsaretiDegistir(item.Su.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_41").Value = NoktalamaIsaretiDegistir(item.SomatikHucre.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_42").Value = NoktalamaIsaretiDegistir(item.MetilenMavisi.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_36").Value = NoktalamaIsaretiDegistir(item.DON_NOK.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_40").Value = NoktalamaIsaretiDegistir(item.Tkm.ToString());
                                frmSutKabul.DataSources.UserDataSources.Item("UD_59").Value = item.Soda;
                            }

                            if (pvalrow == "0")
                            {
                                AracTankTemizle();
                            }
                        }
                        catch (Exception)
                        {
                        }
                        finally
                        {
                            frmSutKabul.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "1" && pVal.BeforeAction)
                    {
                        if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            if (!KantarVerisiOk)
                            {
                                string msg = "";
                                if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                                {
                                    msg = Environment.NewLine;
                                    msg += "İşlem zamanı:" + DateTime.Now + "--" + " Kantar verileri getirilmeden işlem yapılamaz hatası alındı.";
                                    AIFConn.SutKabul.WriteToFile(msg);
                                }
                                else
                                {
                                    msg = Environment.NewLine;
                                    msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " Kantar verileri getirilmeden işlem yapılamaz hatası alındı.";
                                    AIFConn.SutKabul.WriteToFile(msg);
                                }

                                Handler.SAPApplication.MessageBox("Kantar verileri getirilmeden işlem yapılamaz.");
                                BubbleEvent = false;
                                return false;
                            }
                            else
                            {
                                #region old
                                //if (oEditKantarNo.Value != "")
                                //{
                                //    btnKantarAnaveri.Item.Click();
                                //}
                                if (!kantarVerisiManuelGiris && oEditKantarNo.Value != "")
                                {
                                    btnKantarAnaveri.Item.Click();
                                }
                                #endregion
                            }
                        }

                        if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE || frmSutKabul.Mode == BoFormMode.fm_UPDATE_MODE)
                        {
                            double irsaliyeNetSutMikLtFazla = Convert.ToDouble(oedtIrsaliyeNetSutMiktari.Value) + (Convert.ToDouble(oedtIrsaliyeNetSutMiktari.Value) * 10) / 100;
                            double irsaliyeNetSutMikLtEksik = Convert.ToDouble(oedtIrsaliyeNetSutMiktari.Value) - (Convert.ToDouble(oedtIrsaliyeNetSutMiktari.Value) * 10) / 100;

                            if (Convert.ToDouble(edtKantarNetSutMiktariLitre.Value) > irsaliyeNetSutMikLtFazla)
                            {
                                Handler.SAPApplication.MessageBox("Net süt miktarı geçersiz.");
                                BubbleEvent = false;
                                return false;
                            }

                            if (Convert.ToDouble(edtKantarNetSutMiktariLitre.Value) < irsaliyeNetSutMikLtEksik)
                            {
                                Handler.SAPApplication.MessageBox("Net süt miktarı geçersiz.");
                                BubbleEvent = false;
                                return false;
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Item_147" && pVal.BeforeAction)
                    {
                        if (oBtnSutKabulIsleminiTamamla.Item.Enabled == false)
                        {
                            return false;
                        }
                        string msg = "";

                        if (oComboBelgeDurumu.Value.Trim() != "A")
                        {
                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + " Belge oluşturulabilmesi için belge durumu Onaylandı olmalıdır hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " Belge oluşturulabilmesi için belge durumu Onaylandı olmalıdır hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }

                            Handler.SAPApplication.MessageBox("Belge oluşturulabilmesi için belge durumu Onaylandı olmalıdır.");
                            BubbleEvent = false;
                        }

                        var tekilGiris = oChkTekilGiris.Checked;
                        var cogulGiris = oChkCogulGiris.Checked;

                        if (tekilGiris == false && cogulGiris == false)
                        {
                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + " Tekil Giriş veya Çoğul Giriş seçimi yapılmalıdır hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " Tekil Giriş veya Çoğul Giriş seçimi yapılmalıdır hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }

                            Handler.SAPApplication.MessageBox("Tekil Giriş veya Çoğul Giriş seçimi yapılmalıdır.");
                            BubbleEvent = false;
                            return false;
                        }

                        if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + " Eklenmemiş olan belge üzerinde işleme devam edilemez hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " Eklenmemiş olan belge üzerinde işleme devam edilemez hatası alındı.";
                                AIFConn.SutKabul.WriteToFile(msg);
                            }
                            Handler.SAPApplication.MessageBox("Eklenmemiş olan belge üzerinde işleme devam edilemez.");
                            BubbleEvent = false;
                        }
                    }
                    else if (pVal.ItemUID == "Item_147" && !pVal.BeforeAction)
                    {
                        try
                        {
                            if (oBtnSutKabulIsleminiTamamla.Item.Enabled == false)
                            {
                                return false;
                            }
                            string msg = "";
                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value.ToString() == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + " süt kabul işlemini tamamlama için butona tıklandı.";
                                WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + "süt kabul işlemini tamamlama için butona tıklandı.";
                                WriteToFile(msg);
                            }

                            double sutmiktari = parseNumber.parservalues<double>(oedtIrsaliyeNetSutMiktari.Value.ToString());

                            AIFConn.SutDepo.LoadForms(oEditDocEntry.Value, sutmiktari, urunKodu, partiNo, oEditBelgeTarihi.Value, frmSutKabul);
                            //ConstVariables.oRecordset.DoQuery("Select \"U_SutKabulId\" from \"OPDN\" where \"U_SutKabulId\" = '" + oEditDocEntry.Value + "'");

                            //if (ConstVariables.oRecordset.RecordCount == 0)
                            //{
                            //    string tekilCogul = oChkTekilGiris.Checked == true ? "1" : "2";
                            //    bool check = SatinAlmaMalGirisiOlustur(tekilCogul);
                            //    if (check)
                            //    {
                            //    }
                            //}
                            //else
                            //{
                            //    Handler.SAPApplication.MessageBox("Süt girişi daha önceden yapıldığı için tekrar giriş yapılamaz.");
                            //    BubbleEvent = false;
                            //    return false;
                            //}
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_25" && !pVal.BeforeAction)
                    {
                        if (Program.mKod == "010OTATURVT")
                        {
                            if (btnKantarAnaveri.Item.Enabled == false)
                            {
                                return false;
                            }
                            string msg = "";

                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value.ToString() == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditKantarNo.Value + " numarası ile kantar verisi getirme işlemi yapıldı.";
                                WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + oEditKantarNo.Value + " numarası ile kantar verisi getirme işlemi yapıldı.";
                                WriteToFile(msg);
                            }

                            SqlCommand cmd = null;
                            SqlDataAdapter sda = null;
                            System.Data.DataTable dt = new System.Data.DataTable();
                            string sql = "Select FISNO,GIRISMIKTAR,CIKISMIKTAR,NET,PLAKA,SOFOR from OTAT_KANTAR";

                            cmd = new SqlCommand(sql, Connection.sql);

                            if (Connection.sql.State != ConnectionState.Open)
                                Connection.sql.Open();

                            sda = new SqlDataAdapter(cmd);
                            sda.Fill(dt);

                            int val = edtFisNo.Value == "" ? -1 : Convert.ToInt32(edtFisNo.Value);

                            if (val == -1)
                            {
                                if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value.ToString() == "")
                                {
                                    msg = Environment.NewLine;
                                    msg += "İşlem zamanı:" + DateTime.Now + "--" + " Fiş numarası girilmeden işlem yapılamaz hatası alındı.";
                                    WriteToFile(msg);
                                }
                                else
                                {
                                    msg = Environment.NewLine;
                                    msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + "Fiş numarası girilmeden işlem yapılamaz hatası alındı.";
                                    WriteToFile(msg);
                                }

                                Handler.SAPApplication.MessageBox("Fiş numarası girilmeden işlem yapılamaz.");
                                BubbleEvent = false;
                                return false;
                            }

                            var exist = dt.AsEnumerable().Where(x => x.Field<int>("FISNO") == val).ToList();

                            if (exist.Count == 0)
                            {
                                if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value.ToString() == "")
                                {
                                    msg = Environment.NewLine;
                                    msg += "İşlem zamanı:" + DateTime.Now + "--" + " Eşleşen kayıt bulunamadı. Lütfen iletişime geçiniz hatası alındı.";
                                    WriteToFile(msg);
                                }
                                else
                                {
                                    msg = Environment.NewLine;
                                    msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + "Eşleşen kayıt bulunamadı. Lütfen iletişime geçiniz hatası alındı.";
                                    WriteToFile(msg);
                                }
                                Handler.SAPApplication.MessageBox("Eşleşen kayıt bulunamadı. Lütfen iletişime geçiniz.");
                                BubbleEvent = false;
                                return false;
                            }
                            else
                            {
                                foreach (DataRow dr in exist)
                                {
                                    edtAracPlakasi.Value = dr["PLAKA"].ToString();
                                    edtSurucu.Value = dr["SOFOR"].ToString();
                                    edtDoluKantarMiktar.Value = NoktalamaIsaretiDegistir(dr["GIRISMIKTAR"].ToString());
                                    edtBosKantarMiktar.Value = NoktalamaIsaretiDegistir(dr["CIKISMIKTAR"].ToString());
                                    edtKantarNetSutMiktari.Value = NoktalamaIsaretiDegistir(dr["NET"].ToString());

                                    var kantarnetSut = parseNumber.parservalues<double>(NoktalamaIsaretiDegistir(dr["NET"].ToString()));
                                    var girilenNetSutMik = parseNumber.parservalues<double>(oedtIrsaliyeNetSutMiktari.Value);
                                    double netSutLitre = 0;
                                    double miktar = 0;
                                    double sonuc1 = 0;
                                    double sonuc2 = 0;
                                    double yogunluk = 0;
                                    //for (int i = 1; i <= oMatrixAnalizBilgileri.RowCount; i++)
                                    //{
                                    //    miktar = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixAnalizBilgileri.Columns.Item("Col_2").Cells.Item(i).Specific).Value);
                                    //    yogunluk = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrixAnalizBilgileri.Columns.Item("Col_6").Cells.Item(i).Specific).Value);

                                    //    sonuc1 += Math.Round(miktar * yogunluk, 6);
                                    //    sonuc2 += parseNumber.parservalues<double>(miktar.ToString());
                                    //}

                                    //if (sonuc1 > 0)
                                    //{

                                    //    #region net süt miktarı hesaplaması
                                    //    //netSutLitre = Math.Round(girilenNetSutMik / Math.Round((sonuc1 / sonuc2), 6), 6);


                                    //    //edtKantarNetSutMiktariLitre.Value = parseNumber.parservalues<double>(netSutLitre.ToString()).ToString(System.Globalization.CultureInfo.InvariantCulture);
                                    //    #endregion

                                    //    #region net süt miktarı hesaplaması değiştirilmişti geri alındı 2022_01_25
                                    //    netSutLitre = Math.Round(parseNumber.parservalues<double>(edtKantarNetSutMiktari.Value) * (0.103), 6);

                                    //    edtKantarNetSutMiktariLitre.Value = parseNumber.parservalues<double>(netSutLitre.ToString()).ToString(System.Globalization.CultureInfo.InvariantCulture);
                                    //    #endregion

                                    //}
                                    netSutLitre = Math.Round(parseNumber.parservalues<double>(edtKantarNetSutMiktari.Value) / (1.030), 6);

                                    edtKantarNetSutMiktariLitre.Value = parseNumber.parservalues<double>(netSutLitre.ToString()).ToString(System.Globalization.CultureInfo.InvariantCulture);

                                    //edtKantarNetSutMiktari.Value = NoktalamaIsaretiDegistir(dr["NET"].ToString());
                                }
                            }

                            if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value.ToString() == "")
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditKantarNo.Value + " numarası ile kantar verisi getirme işlemi başarılı.";
                                WriteToFile(msg);
                            }
                            else
                            {
                                msg = Environment.NewLine;
                                msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + "numarası ile kantar verisi getirme işlemi başarılı.";
                                WriteToFile(msg);
                            }

                            KantarVerisiOk = true;
                            kantarVerisiManuelGiris = false;
                        }
                    }
                    else if (pVal.ItemUID == "Item_166" && pVal.BeforeAction)
                    {
                        comboBelgeDurumuTiklanmaDegeri = oComboBelgeDurumu.Value.Trim();
                    }
                    else if (pVal.ItemUID == "Item_188" && !pVal.BeforeAction)
                    {
                        //[ItemAtt("Item_12")]  edtDoluKantarMiktar; 
                        //[ItemAtt("Item_15")]  edtBosKantarMiktar; 
                        //[ItemAtt("Item_18")]  edtKantarNetSutMiktari; 
                        //[ItemAtt("Item_180")]   edtKantarNetSutMiktariLitre;
                        if (btnKantarVerisiniManuelGir.Item.Enabled == false)
                        {
                            return false;
                        }
                        edtDoluKantarMiktar.Item.Enabled = true;
                        edtBosKantarMiktar.Item.Enabled = true;
                        edtKantarNetSutMiktari.Item.Enabled = true;
                        edtKantarNetSutMiktariLitre.Item.Enabled = true;
                        kantarVerisiManuelGiris = true;
                        KantarVerisiOk = true;
                    }
                    break;

                case BoEventTypes.et_DOUBLE_CLICK:
                    break;

                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                    break;

                case BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
                    break;

                case BoEventTypes.et_VALIDATE:
                    if (pVal.ItemUID == "Item_35" && pVal.BeforeAction) // Sıcaklık
                    {
                        try
                        {
                            double sicaklik = edtAnalizSicaklik.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizSicaklik.Value);
                            if (sicaklik != 0)
                            {
                                double sicaklikMin = SutKabulUyarlama.Select(x => x.SicaklikMin).FirstOrDefault();
                                double sicaklikMax = SutKabulUyarlama.Select(x => x.SicaklikMax).FirstOrDefault();

                                if (sicaklikMin != 0 || sicaklikMax != 0)
                                {
                                    if (sicaklik < sicaklikMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Sıcaklık değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (sicaklik > sicaklikMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Sıcaklık değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_39" && pVal.BeforeAction) // Briks
                    {
                        try
                        {
                            double briks = edtAnalizBriks.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizBriks.Value);
                            if (briks != 0)
                            {
                                double briksMin = SutKabulUyarlama.Select(x => x.BriksMin).FirstOrDefault();
                                double briksMax = SutKabulUyarlama.Select(x => x.BriksMax).FirstOrDefault();

                                if (briksMin != 0 || briksMax != 0)
                                {
                                    if (briks < briksMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Briks değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (briks > briksMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Briks değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_41" && pVal.BeforeAction)// Yağ
                    {
                        try
                        {
                            double yag = edtAnalizYag.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizYag.Value);
                            if (yag != 0)
                            {
                                double YagMin = SutKabulUyarlama.Select(x => x.YagMin).FirstOrDefault();
                                double YagMax = SutKabulUyarlama.Select(x => x.YagMax).FirstOrDefault();

                                if (YagMin != 0 || YagMax != 0)
                                {
                                    if (yag < YagMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Yağ değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (yag > YagMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Yağ değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_43" && pVal.BeforeAction) // Yoğunluk 
                    {
                        try
                        {
                            double yogunluk = edtAnalizYogunluk.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizYogunluk.Value);
                            if (yogunluk != 0)
                            {
                                double YogunlukMin = SutKabulUyarlama.Select(x => x.YogunlukMin).FirstOrDefault();
                                double YogunlukMax = SutKabulUyarlama.Select(x => x.YogunlukMax).FirstOrDefault();

                                if (YogunlukMin != 0 || YogunlukMax != 0)
                                {
                                    if (yogunluk < YogunlukMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Yoğunluk değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (yogunluk > YogunlukMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Yoğunluk değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_45" && pVal.BeforeAction) // ph 
                    {
                        try
                        {
                            double pH = edtAnalizpH.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizpH.Value);
                            if (pH != 0)
                            {
                                double pHMin = SutKabulUyarlama.Select(x => x.phMin).FirstOrDefault();
                                double pHMax = SutKabulUyarlama.Select(x => x.phMax).FirstOrDefault();

                                if (pHMin != 0 || pHMax != 0)
                                {
                                    if (pH < pHMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("PH değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (pH > pHMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("PH değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_47" && pVal.BeforeAction) // sh 
                    {
                        try
                        {
                            double SH = edtAnalizSH.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizSH.Value);
                            if (SH != 0)
                            {
                                double SHMin = SutKabulUyarlama.Select(x => x.shMin).FirstOrDefault();
                                double SHMax = SutKabulUyarlama.Select(x => x.shMax).FirstOrDefault();

                                if (SHMin != 0 || SHMax != 0)
                                {
                                    if (SH < SHMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("SH değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (SH > SHMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("SH değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_51" && pVal.BeforeAction)
                    {
                        try
                        {
                            //double TemizlikKoku = edtAnalizTemizkikKoku.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizTemizkikKoku.Value);
                            //if (TemizlikKoku != 0)
                            //{
                            //    double TemizlikKokuMin = SutKabulUyarlama.Select(x => x.TemizlikKokuMin).FirstOrDefault();
                            //    double TemizlikKokuMax = SutKabulUyarlama.Select(x => x.TemizlikKokuMax).FirstOrDefault();

                            //    if (TemizlikKokuMin != 0 || TemizlikKokuMax != 0)
                            //    {
                            //        if (TemizlikKoku < TemizlikKokuMin)
                            //        {
                            //            btnKaydet.Item.Enabled = false;
                            //            Handler.SAPApplication.MessageBox("Temizlik Koku değeri minimum değerden düşük girilemez.");
                            //            BubbleEvent = false;
                            //            return false;
                            //        }
                            //        else if (TemizlikKoku > TemizlikKokuMax)
                            //        {
                            //            btnKaydet.Item.Enabled = false;
                            //            Handler.SAPApplication.MessageBox("Temizlik Koku değeri maksimum değerden yüksek girilemez.");
                            //            BubbleEvent = false;
                            //            return false;
                            //        }

                            //        btnKaydet.Item.Enabled = true;
                            //    }
                            //}
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_53" && pVal.BeforeAction)
                    {
                        try
                        {
                            //double Soda = edtAnalizSoda.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizSoda.Value);
                            //if (Soda != 0)
                            //{
                            //    double SodaMin = SutKabulUyarlama.Select(x => x.SodaMin).FirstOrDefault();
                            //    double SodaMax = SutKabulUyarlama.Select(x => x.SodaMax).FirstOrDefault();

                            //    if (SodaMin != 0 || SodaMax != 0)
                            //    {
                            //        if (Soda < SodaMin)
                            //        {
                            //            btnKaydet.Item.Enabled = false;
                            //            Handler.SAPApplication.MessageBox("Soda değeri minimum değerden düşük girilemez.");
                            //            BubbleEvent = false;
                            //            return false;
                            //        }
                            //        else if (Soda > SodaMax)
                            //        {
                            //            btnKaydet.Item.Enabled = false;
                            //            Handler.SAPApplication.MessageBox("Soda değeri maksimum değerden yüksek girilemez.");
                            //            BubbleEvent = false;
                            //            return false;
                            //        }

                            //        btnKaydet.Item.Enabled = true;
                            //    }
                            //}
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_55" && pVal.BeforeAction)
                    {
                        try
                        {
                            //double Antibiyotik = edtAnalizAntibiyotik.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizAntibiyotik.Value);
                            //if (Antibiyotik != 0)
                            //{
                            //    double AntibiyotikMin = SutKabulUyarlama.Select(x => x.AntibiyotikMin).FirstOrDefault();
                            //    double AntibiyotikMax = SutKabulUyarlama.Select(x => x.AntibiyotikMax).FirstOrDefault();

                            //    if (AntibiyotikMin != 0 || AntibiyotikMax != 0)
                            //    {
                            //        if (Antibiyotik < AntibiyotikMin)
                            //        {
                            //            btnKaydet.Item.Enabled = false;
                            //            Handler.SAPApplication.MessageBox("Antibiyotik değeri minimum değerden düşük girilemez.");
                            //            BubbleEvent = false;
                            //            return false;
                            //        }
                            //        else if (Antibiyotik > AntibiyotikMax)
                            //        {
                            //            btnKaydet.Item.Enabled = false;
                            //            Handler.SAPApplication.MessageBox("Antibiyotik değeri maksimum değerden yüksek girilemez.");
                            //            BubbleEvent = false;
                            //            return false;
                            //        }

                            //        btnKaydet.Item.Enabled = true;
                            //    }
                            //}
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_57" && pVal.BeforeAction) // Yağsız Kuru Madde 
                    {
                        try
                        {
                            double Ykm = edtAnalizYkm.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizYkm.Value);
                            if (Ykm != 0)
                            {
                                double YkmMin = SutKabulUyarlama.Select(x => x.YkmMin).FirstOrDefault();
                                double YkmMax = SutKabulUyarlama.Select(x => x.YkmMax).FirstOrDefault();

                                if (YkmMin != 0 || YkmMax != 0)
                                {
                                    if (Ykm < YkmMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Ykm değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (Ykm > YkmMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Ykm değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_59" && pVal.BeforeAction) // alkol 
                    {
                        try
                        {
                            double Alkol = oComboAnalizAlkol.Value == "" ? 0 : parseNumber.parservalues<double>(oComboAnalizAlkol.Value);
                            if (Alkol != 0)
                            {
                                double AlkolMin = SutKabulUyarlama.Select(x => x.AlkolMin).FirstOrDefault();
                                double AlkolMax = SutKabulUyarlama.Select(x => x.AlkolMax).FirstOrDefault();

                                if (AlkolMin != 0 || AlkolMax != 0)
                                {
                                    if (Alkol < AlkolMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Alkol değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (Alkol > AlkolMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Alkol değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_61" && pVal.BeforeAction) // Kaynamış sh
                    {
                        try
                        {
                            double KaynatSh = edtAnalizKaynatSh.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizKaynatSh.Value);
                            if (KaynatSh != 0)
                            {
                                double KaynatShMin = SutKabulUyarlama.Select(x => x.KaynatsHMin).FirstOrDefault();
                                double KaynatShMax = SutKabulUyarlama.Select(x => x.KaynatsHMax).FirstOrDefault();

                                if (KaynatShMin != 0 || KaynatShMax != 0)
                                {
                                    if (KaynatSh < KaynatShMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Kaynat SH değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (KaynatSh > KaynatShMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Kaynat SH değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_65" && pVal.BeforeAction) // Protein 
                    {
                        try
                        {
                            double Protein = edtAnalizProtein.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizProtein.Value);
                            if (Protein != 0)
                            {
                                double ProteinMin = SutKabulUyarlama.Select(x => x.ProteinMin).FirstOrDefault();
                                double ProteinMax = SutKabulUyarlama.Select(x => x.ProteinMax).FirstOrDefault();

                                if (ProteinMin != 0 || ProteinMax != 0)
                                {
                                    if (Protein < ProteinMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Protein değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (Protein > ProteinMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Protein değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_67" && pVal.BeforeAction) // Laktoz 
                    {
                        try
                        {
                            double Laktoz = edtAnalizLaktoz.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizLaktoz.Value);
                            if (Laktoz != 0)
                            {
                                double LaktozMin = SutKabulUyarlama.Select(x => x.LaktozMin).FirstOrDefault();
                                double LaktozMax = SutKabulUyarlama.Select(x => x.LaktozMax).FirstOrDefault();

                                if (LaktozMin != 0 || LaktozMax != 0)
                                {
                                    if (Laktoz < LaktozMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Laktoz değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (Laktoz > LaktozMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Laktoz değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_69" && pVal.BeforeAction) // Katılan Su
                    {
                        try
                        {
                            double KatilanSu = edtAnalizKatilanSu.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizKatilanSu.Value);
                            if (KatilanSu != 0)
                            {
                                double KatilanSuMin = SutKabulUyarlama.Select(x => x.KatilanSuMin).FirstOrDefault();
                                double KatilanSuMax = SutKabulUyarlama.Select(x => x.KatilanSuMax).FirstOrDefault();

                                if (KatilanSuMin != 0 || KatilanSuMax != 0)
                                {
                                    if (KatilanSu < KatilanSuMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Katılan Su değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (KatilanSu > KatilanSuMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Katılan Su değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_71" && pVal.BeforeAction) // Somatik Hücre 
                    {
                        try
                        {
                            double SomatikHucre = edtAnalizSomatikHucre.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizSomatikHucre.Value);
                            if (SomatikHucre != 0)
                            {
                                double SomatikHucreMin = SutKabulUyarlama.Select(x => x.SomatikHucreMin).FirstOrDefault();
                                double SomatikHucreMax = SutKabulUyarlama.Select(x => x.SomatikHucreMax).FirstOrDefault();

                                if (SomatikHucreMin != 0 || SomatikHucreMax != 0)
                                {
                                    if (SomatikHucre < SomatikHucreMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Somatik Hücre değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (SomatikHucre > SomatikHucreMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Somatik Hücre değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_73" && pVal.BeforeAction) // Metilen Mavisi 
                    {
                        try
                        {
                            double MetilenMavisi = edtAnalizMetilenMavisi.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizMetilenMavisi.Value);
                            if (MetilenMavisi != 0)
                            {
                                double MetilenMavisiMin = SutKabulUyarlama.Select(x => x.MetilenMaviMin).FirstOrDefault();
                                double MetilenMavisiMax = SutKabulUyarlama.Select(x => x.MetilenMaviMax).FirstOrDefault();

                                if (MetilenMavisiMin != 0 || MetilenMavisiMax != 0)
                                {
                                    if (MetilenMavisi < MetilenMavisiMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Metilen Mavisi değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (MetilenMavisi > MetilenMavisiMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Metilen Mavisi değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_141" && pVal.BeforeAction) // Donma Noktası
                    {
                        try
                        {

                            double DonmaNoktasi = edtAnalizDonmaNoktasi.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizDonmaNoktasi.Value);
                            if (DonmaNoktasi != 0)
                            {
                                double DonmaNoktasiMin = SutKabulUyarlama.Select(x => x.DonmaNoktasıMin).FirstOrDefault();
                                double DonmaNoktasiMax = SutKabulUyarlama.Select(x => x.DonmaNoktasıMax).FirstOrDefault();

                                if (DonmaNoktasiMin != 0 || DonmaNoktasiMax != 0)
                                {
                                    if (DonmaNoktasi < DonmaNoktasiMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Donma Noktasi değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (DonmaNoktasi > DonmaNoktasiMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Donma Noktasi değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_148" && pVal.BeforeAction) // TKM
                    {
                        try
                        {
                            double TKM = edtAnalizTKM.Value == "" ? 0 : parseNumber.parservalues<double>(edtAnalizTKM.Value);
                            if (TKM != 0)
                            {
                                double TKMMin = SutKabulUyarlama.Select(x => x.TKMMin).FirstOrDefault();
                                double TKMMax = SutKabulUyarlama.Select(x => x.TKMMax).FirstOrDefault();

                                if (TKMMin != 0 || TKMMax != 0)
                                {
                                    if (TKM < TKMMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("TKM değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (TKM > TKMMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("TKM değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_170" && pVal.BeforeAction) // Süt Temizliği
                    {
                        try
                        {
                            double SutTemizlik = edtSutTemizligi.Value == "" ? 0 : parseNumber.parservalues<double>(edtSutTemizligi.Value);
                            if (SutTemizlik != 0)
                            {
                                double SutTemizligiMin = SutKabulUyarlama.Select(x => x.SutTemizligiMin).FirstOrDefault();
                                double SutTemizligiMax = SutKabulUyarlama.Select(x => x.SutTemizligiMax).FirstOrDefault();

                                if (SutTemizligiMin != 0 || SutTemizligiMax != 0)
                                {
                                    if (SutTemizlik < SutTemizligiMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Süt Temizliği değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (SutTemizlik > SutTemizligiMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Süt Temizliği değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_172" && pVal.BeforeAction) // Tank Temizliği
                    {
                        try
                        {
                            double TankTemizligi = edtTankTemizligi.Value == "" ? 0 : parseNumber.parservalues<double>(edtTankTemizligi.Value);
                            if (TankTemizligi != 0)
                            {
                                double TankTemizligiMin = SutKabulUyarlama.Select(x => x.TankTemizligiMin).FirstOrDefault();
                                double TankTemizligiMax = SutKabulUyarlama.Select(x => x.TankTemizligiMax).FirstOrDefault();

                                if (TankTemizligiMin != 0 || TankTemizligiMax != 0)
                                {
                                    if (TankTemizligi < TankTemizligiMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Tank Temizliği değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (TankTemizligi > TankTemizligiMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Tank Temizliği değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_174" && pVal.BeforeAction) // Ekipman Temizliği
                    {
                        try
                        {
                            double EkipmanTemizligi = edtEkipmanTemizligi.Value == "" ? 0 : parseNumber.parservalues<double>(edtEkipmanTemizligi.Value);
                            if (EkipmanTemizligi != 0)
                            {
                                double EkipmanTemizligiMin = SutKabulUyarlama.Select(x => x.EkipmanTemizligiMin).FirstOrDefault();
                                double EkipmanTemizligiMax = SutKabulUyarlama.Select(x => x.EkipmanTemizligiMax).FirstOrDefault();

                                if (EkipmanTemizligiMin != 0 || EkipmanTemizligiMax != 0)
                                {
                                    if (EkipmanTemizligi < EkipmanTemizligiMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Ekipman Temizliği değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (EkipmanTemizligi > EkipmanTemizligiMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Ekipman Temizliği değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_176" && pVal.BeforeAction) // Araç Kasası Temizliği
                    {
                        try
                        {
                            double AracKasaTemizligi = edtAracKasasiTemizligi.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracKasasiTemizligi.Value);
                            if (AracKasaTemizligi != 0)
                            {
                                double AracKasaTemizligiMin = SutKabulUyarlama.Select(x => x.AracKasaTemizligiMin).FirstOrDefault();
                                double AracKasaTemizligiMax = SutKabulUyarlama.Select(x => x.AracKasaTemizligiMax).FirstOrDefault();

                                if (AracKasaTemizligiMin != 0 || AracKasaTemizligiMax != 0)
                                {
                                    if (AracKasaTemizligi < AracKasaTemizligiMin)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Kasa Temizliği değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracKasaTemizligi > AracKasaTemizligiMax)
                                    {
                                        btnKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Kasa Temizliği değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_83" && pVal.BeforeAction) // Araç Tank Sıcaklık
                    {
                        try
                        {
                            double AracTankDerece = edtAracTankDerece.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankDerece.Value);
                            if (AracTankDerece != 0)
                            {
                                double AracTankSicaklikMin = SutKabulUyarlama.Select(x => x.AracTankSicaklikMin).FirstOrDefault();
                                double AracTankSicaklikMax = SutKabulUyarlama.Select(x => x.AracTankSicaklikMax).FirstOrDefault();

                                if (AracTankSicaklikMin != 0 || AracTankSicaklikMax != 0)
                                {
                                    if (AracTankDerece < AracTankSicaklikMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Sıcaklık değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTankDerece > AracTankSicaklikMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Sıcaklık değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_85" && pVal.BeforeAction) // Araç Tank Briks
                    {
                        try
                        {
                            double AracTankBriks = edtAracTankBrix.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankBrix.Value);
                            if (AracTankBriks != 0)
                            {
                                double AracTankBriksMin = SutKabulUyarlama.Select(x => x.AracTankBriksMin).FirstOrDefault();
                                double AracTankBriksMax = SutKabulUyarlama.Select(x => x.AracTankBriksMax).FirstOrDefault();

                                if (AracTankBriksMin != 0 || AracTankBriksMax != 0)
                                {
                                    if (AracTankBriks < AracTankBriksMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Briks değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTankBriks > AracTankBriksMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Briks değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_87" && pVal.BeforeAction) // Araç Tank Yağ
                    {
                        try
                        {
                            double AracTankYag = edtAracTankYag.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankYag.Value);
                            if (AracTankYag != 0)
                            {
                                double AracTankYagMin = SutKabulUyarlama.Select(x => x.AracTankYagMin).FirstOrDefault();
                                double AracTankYagMax = SutKabulUyarlama.Select(x => x.AracTankYagMax).FirstOrDefault();

                                if (AracTankYagMin != 0 || AracTankYagMax != 0)
                                {
                                    if (AracTankYag < AracTankYagMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Yağ değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTankYag > AracTankYagMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Yağ değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_89" && pVal.BeforeAction) // Araç Tank Yoğunluk
                    {
                        try
                        {
                            double AracTankYogunluk = edtAracTankDans.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankDans.Value);
                            if (AracTankYogunluk != 0)
                            {
                                double AracTankYogunlukMin = SutKabulUyarlama.Select(x => x.AracTankYogunlukMin).FirstOrDefault();
                                double AracTankYogunlukMax = SutKabulUyarlama.Select(x => x.AracTankYogunlukMax).FirstOrDefault();

                                if (AracTankYogunlukMin != 0 || AracTankYogunlukMax != 0)
                                {
                                    if (AracTankYogunluk < AracTankYogunlukMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Yoğunluk değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTankYogunluk > AracTankYogunlukMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Yoğunluk değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_91" && pVal.BeforeAction) // Araç Tank ph
                    {
                        try
                        {
                            double AracTankpH = edtAracTankpH.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankpH.Value);
                            if (AracTankpH != 0)
                            {
                                double AracTankpHMin = SutKabulUyarlama.Select(x => x.AracTankphMin).FirstOrDefault();
                                double AracTankpHMax = SutKabulUyarlama.Select(x => x.AracTankphMax).FirstOrDefault();

                                if (AracTankpHMin != 0 || AracTankpHMax != 0)
                                {
                                    if (AracTankpH < AracTankpHMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank pH değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTankpH > AracTankpHMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank pH değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_93" && pVal.BeforeAction) // Araç Tank sh
                    {
                        try
                        {
                            double AracTanksH = edtAracTankSH.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankSH.Value);
                            if (AracTanksH != 0)
                            {
                                double AracTanksHMin = SutKabulUyarlama.Select(x => x.AracTankshMin).FirstOrDefault();
                                double AracTanksHMax = SutKabulUyarlama.Select(x => x.AracTankshMax).FirstOrDefault();

                                if (AracTanksHMin != 0 || AracTanksHMax != 0)
                                {
                                    if (AracTanksH < AracTanksHMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank sH değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTanksH > AracTanksHMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank sH değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_95" && pVal.BeforeAction) // Araç Tank Kaynamış sH
                    {
                        try
                        {
                            double AracTankKaynamissH = edtAracTankKaynamisSH.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankKaynamisSH.Value);
                            if (AracTankKaynamissH != 0)
                            {
                                double AracTankKaynamissHMin = SutKabulUyarlama.Select(x => x.AracTankKaynamisShMin).FirstOrDefault();
                                double AracTankKaynamissHMax = SutKabulUyarlama.Select(x => x.AracTankKaynamisShMax).FirstOrDefault();

                                if (AracTankKaynamissHMin != 0 || AracTankKaynamissHMax != 0)
                                {
                                    if (AracTankKaynamissH < AracTankKaynamissHMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Kaynamış sH değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTankKaynamissH > AracTankKaynamissHMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Kaynamış sH değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_103" && pVal.BeforeAction) // Araç Tank Yağsız Kuru madde
                    {
                        try
                        {
                            double AracTankYagsizKuruMadde = edtAracTankYagsiz.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankYagsiz.Value);
                            if (AracTankYagsizKuruMadde != 0)
                            {
                                double AracTankYagsizKuruMaddeMin = SutKabulUyarlama.Select(x => x.AracTankYagsizKuruMaddeMin).FirstOrDefault();
                                double AracTankYagsizKuruMaddeMax = SutKabulUyarlama.Select(x => x.AracTankYagsizKuruMaddeMax).FirstOrDefault();

                                if (AracTankYagsizKuruMaddeMin != 0 || AracTankYagsizKuruMaddeMax != 0)
                                {
                                    if (AracTankYagsizKuruMadde < AracTankYagsizKuruMaddeMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Yağsız Kuru Madde değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTankYagsizKuruMadde > AracTankYagsizKuruMaddeMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Yağsız Kuru Madde değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_105" && pVal.BeforeAction) // Araç Tank Protein
                    {
                        try
                        {
                            double AracTankProtein = edtAracTankProtein.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankProtein.Value);
                            if (AracTankProtein != 0)
                            {
                                double AracTankProteinMin = SutKabulUyarlama.Select(x => x.AracTankProteinMin).FirstOrDefault();
                                double AracTankProteinMax = SutKabulUyarlama.Select(x => x.AracTankProteinMax).FirstOrDefault();

                                if (AracTankProteinMin != 0 || AracTankProteinMax != 0)
                                {
                                    if (AracTankProtein < AracTankProteinMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Protein değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTankProtein > AracTankProteinMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Protein değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_109" && pVal.BeforeAction) // Araç Tank Laktoz
                    {
                        try
                        {
                            double AracTankLaktoz = edtAracTankLaktoz.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankLaktoz.Value);
                            if (AracTankLaktoz != 0)
                            {
                                double AracTankLaktozMin = SutKabulUyarlama.Select(x => x.AracTankLaktozMin).FirstOrDefault();
                                double AracTankLaktozMax = SutKabulUyarlama.Select(x => x.AracTankLaktozMax).FirstOrDefault();

                                if (AracTankLaktozMin != 0 || AracTankLaktozMax != 0)
                                {
                                    if (AracTankLaktoz < AracTankLaktozMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Laktoz değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTankLaktoz > AracTankLaktozMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Laktoz değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_111" && pVal.BeforeAction) // Araç Tank Katilan Su
                    {
                        try
                        {
                            double AracTankKatilanSu = edtAracTankSu.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankSu.Value);
                            if (AracTankKatilanSu != 0)
                            {
                                double AracTankKatilanSuMin = SutKabulUyarlama.Select(x => x.AracTankKatilanSuMin).FirstOrDefault();
                                double AracTankKatilanSuMax = SutKabulUyarlama.Select(x => x.AracTankKatilanSuMax).FirstOrDefault();

                                if (AracTankKatilanSuMin != 0 || AracTankKatilanSuMax != 0)
                                {
                                    if (AracTankKatilanSu < AracTankKatilanSuMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Katılan Su değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTankKatilanSu > AracTankKatilanSuMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Katılan Su değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_117" && pVal.BeforeAction) // Araç Tank Somatik Hücre
                    {
                        try
                        {
                            double AracTankSomatikHucre = edtAracTankSomatikHucre.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankSomatikHucre.Value);
                            if (AracTankSomatikHucre != 0)
                            {
                                double AracTankSomatikHucreMin = SutKabulUyarlama.Select(x => x.AracTankSomatikHucreMin).FirstOrDefault();
                                double AracTankSomatikHucreMax = SutKabulUyarlama.Select(x => x.AracTankSomatikHucreMax).FirstOrDefault();

                                if (AracTankSomatikHucreMin != 0 || AracTankSomatikHucreMax != 0)
                                {
                                    if (AracTankSomatikHucre < AracTankSomatikHucreMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Somatik Hücre değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTankSomatikHucre > AracTankSomatikHucreMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Somatik Hücre değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_119" && pVal.BeforeAction) // Araç Tank Metilen Mavi
                    {
                        try
                        {
                            double AracTankMetilenMavi = edtAracTankMetilenMavisi.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankMetilenMavisi.Value);
                            if (AracTankMetilenMavi != 0)
                            {
                                double AracTankMetilenMaviMin = SutKabulUyarlama.Select(x => x.AracTankMetilenMavisiMin).FirstOrDefault();
                                double AracTankMetilenMaviMax = SutKabulUyarlama.Select(x => x.AracTankMetilenMavisiMax).FirstOrDefault();

                                if (AracTankMetilenMaviMin != 0 || AracTankMetilenMaviMax != 0)
                                {
                                    if (AracTankMetilenMavi < AracTankMetilenMaviMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Metilen Mavisi değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTankMetilenMavi > AracTankMetilenMaviMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Metilen Mavisi değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_107" && pVal.BeforeAction) // Araç Tank Donma Noktasi
                    {
                        try
                        {
                            double AracTankDonmaNoktasi = edtAracTankDonNok.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankDonNok.Value);
                            if (AracTankDonmaNoktasi != 0)
                            {
                                double AracTankDonmaNoktasiMin = SutKabulUyarlama.Select(x => x.AracTankDonmaNoktasiMin).FirstOrDefault();
                                double AracTankDonmaNoktasiMax = SutKabulUyarlama.Select(x => x.AracTankDonmaNoktasiMax).FirstOrDefault();

                                if (AracTankDonmaNoktasiMin != 0 || AracTankDonmaNoktasiMax != 0)
                                {
                                    if (AracTankDonmaNoktasi < AracTankDonmaNoktasiMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Donma Noktası değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTankDonmaNoktasi > AracTankDonmaNoktasiMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank Donma Noktası değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_115" && pVal.BeforeAction) // Araç Tank TKM
                    {
                        try
                        {
                            double AracTankTKM = edtAracTankTkm.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracTankTkm.Value);
                            if (AracTankTKM != 0)
                            {
                                double AracTankTKMMin = SutKabulUyarlama.Select(x => x.AracTankTKMMin).FirstOrDefault();
                                double AracTankTKMMax = SutKabulUyarlama.Select(x => x.AracTankTKMMax).FirstOrDefault();

                                if (AracTankTKMMin != 0 || AracTankTKMMax != 0)
                                {
                                    if (AracTankTKM < AracTankTKMMin)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank TKM değeri minimum değerden düşük girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }
                                    else if (AracTankTKM > AracTankTKMMax)
                                    {
                                        btnAracTankKaydet.Item.Enabled = false;
                                        Handler.SAPApplication.MessageBox("Araç Tank TKM değeri maksimum değerden yüksek girilemez.");
                                        BubbleEvent = false;
                                        return false;
                                    }

                                    btnAracTankKaydet.Item.Enabled = true;
                                }
                            }
                        }
                        catch (Exception)
                        {
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
                    if (pVal.ItemUID == "Item_34" && pVal.BeforeAction)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                        oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                        SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                        oCFL = frmSutKabul.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                        SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                        SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                        SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                        oCFL.SetConditions(oEmptyConts);
                        oCons = oCFL.GetConditions();

                        oCon = oCons.Add();
                        oCon.Alias = "CardType";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "S";

                        oCon.Relationship = BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "QryGroup1";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";

                        oCon.Relationship = BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "validFor";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";

                        oCFL.SetConditions(oCons);
                    }
                    else if (pVal.ItemUID == "Item_34" && !pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                            string Val = "";
                            Val = oDataTable.GetValue("CardCode", 0).ToString();
                            try
                            {
                                edtAnalizTedarikci.Value = Val;
                            }
                            catch (Exception)
                            {
                            }

                            try
                            {
                                Val = oDataTable.GetValue("CardName", 0).ToString();
                                edtAnalizTedarikciAdi.Value = Val;
                            }
                            catch (Exception)
                            {
                            }

                            try
                            {
                                Val = oDataTable.GetValue("U_Nitelik", 0).ToString();
                                oComboNitelik.Select(Val, BoSearchKey.psk_ByValue);
                            }
                            catch (Exception)
                            {
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_63" && !pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                            string Val = "";
                            Val = oDataTable.GetValue("WhsCode", 0).ToString();
                            try
                            {
                                edtAnalizDepo.Value = Val;
                            }
                            catch (Exception)
                            {
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_159" && pVal.BeforeAction)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                        oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                        SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                        oCFL = frmSutKabul.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                        SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                        SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                        SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                        oCFL.SetConditions(oEmptyConts);
                        oCons = oCFL.GetConditions();

                        oCon = oCons.Add();
                        oCon.Alias = "CardType";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "S";

                        oCon.Relationship = BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "QryGroup1";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";

                        oCon.Relationship = BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "validFor";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";

                        oCFL.SetConditions(oCons);
                    }
                    else if (pVal.ItemUID == "Item_159" && !pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                            string Val = "";
                            Val = oDataTable.GetValue("CardCode", 0).ToString();
                            try
                            {
                                oEditSatici.Value = Val;
                            }
                            catch (Exception)
                            {
                            }

                            try
                            {
                                Val = oDataTable.GetValue("CardName", 0).ToString();
                                oEditSaticiAdi.Value = Val;

                                string msg = "";
                                if (frmSutKabul.Mode == BoFormMode.fm_ADD_MODE && oEditDocEntry.Value == "")
                                {
                                    msg = Environment.NewLine;
                                    msg += "İşlem zamanı:" + DateTime.Now + "--" + " tedarikçi seçimi yapıldı.";
                                    AIFConn.SutKabul.WriteToFile(msg);
                                }
                                else
                                {
                                    msg = Environment.NewLine;
                                    msg += "İşlem zamanı:" + DateTime.Now + "--" + oEditDocEntry.Value + " numaralı süt kabul belgesi için " + " tedarikçi seçimi yeniden yapıldı.";
                                    AIFConn.SutKabul.WriteToFile(msg);
                                }
                            }
                            catch (Exception)
                            {
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_121" && pVal.BeforeAction)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                        oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                        SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                        oCFL = frmSutKabul.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                        SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                        SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                        SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                        oCFL.SetConditions(oEmptyConts);
                        oCons = oCFL.GetConditions();

                        oCon = oCons.Add();
                        oCon.Alias = "CardType";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "S";

                        oCon.Relationship = BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "QryGroup1";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";

                        oCon.Relationship = BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "validFor";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";

                        oCFL.SetConditions(oCons);
                    }
                    else if (pVal.ItemUID == "Item_121" && !pVal.BeforeAction)
                    {
                        try
                        {
                            SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                            string Val = "";
                            Val = oDataTable.GetValue("CardCode", 0).ToString();
                            try
                            {
                                edtAracTedarikci.Value = Val;
                            }
                            catch (Exception)
                            {
                            }

                            try
                            {
                                Val = oDataTable.GetValue("CardName", 0).ToString();
                                edtAracTedarikciAdi.Value = Val;
                            }
                            catch (Exception)
                            {
                            }
                        }
                        catch (Exception)
                        {
                        }
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
        private string comboBelgeDurumuTiklanmaDegeri = "";

        private void AgirlikliPuanHesapla()
        {
            try
            {
                var SutTemizligiVal = edtSutTemizligi.Value == "" ? 0 : parseNumber.parservalues<double>(edtSutTemizligi.Value);
                var TankTemizligiVal = edtTankTemizligi.Value == "" ? 0 : parseNumber.parservalues<double>(edtTankTemizligi.Value);
                var EkipmanTemizligiVal = edtEkipmanTemizligi.Value == "" ? 0 : parseNumber.parservalues<double>(edtEkipmanTemizligi.Value);
                var AracKasasiTemizligiVal = edtAracKasasiTemizligi.Value == "" ? 0 : parseNumber.parservalues<double>(edtAracKasasiTemizligi.Value);

                var sonuc = (SutTemizligiVal * 40) / 100;
                sonuc += (TankTemizligiVal * 20) / 100;
                sonuc += (EkipmanTemizligiVal * 20) / 100;
                sonuc += (AracKasasiTemizligiVal * 20) / 100;

                frmSutKabul.DataSources.UserDataSources.Item("UD_65").Value = NoktalamaIsaretiDegistir(sonuc.ToString());
            }
            catch (Exception)
            {
            }
        }

        private bool AnalizBilgileriAlanlariniKontrolEt()
        {
            if (oComboAnalizSutTuru.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri süt türünü kontrol ediniz.");
                return false;
            }

            //if (oComboNitelik.Value.Trim() == "")
            //{
            //    Handler.SAPApplication.MessageBox("Analiz bilgileri nitelik alanını kontrol ediniz.");
            //    return false;
            //}

            if (edtAnalizTedarikci.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri tedarikçi alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizDepo.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri depo alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizMiktar.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri miktar alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizMiktar.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri miktar alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizSicaklik.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri sıcaklık alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizSicaklik.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri sıcaklık alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizBriks.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri briks alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizBriks.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri briks alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizYag.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri yağ alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizMiktar.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri yağ alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizYogunluk.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri yoğunluk alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizYogunluk.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri yoğunluk alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizpH.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri PH alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizpH.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri PH alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizSH.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri SH alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizSH.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri SH alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizKaynatSh.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri kaynamış SH alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizKaynatSh.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri kaynamış SH alanını kontrol ediniz.");
                return false;
            }

            if (oComboAnalizAlkol.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri alkol alanını kontrol ediniz.");
                return false;
            }

            //if (oComboAnalizTemizkikKoku.Value.Trim() == "")
            //{
            //    Handler.SAPApplication.MessageBox("Analiz bilgileri temizlik koku alanını kontrol ediniz.");
            //    return false;
            //}

            if (oComboAnalizSoda.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri soda alanını kontrol ediniz.");
                return false;
            }

            if (oComboAnalizAntibiyotik.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri antibiyotik alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizYkm.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri yağsız kuru madde alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizYkm.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri yağsız kuru madde alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizProtein.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri protein alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizProtein.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri protein alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizLaktoz.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri laktoz alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizLaktoz.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri laktoz alanını kontrol ediniz.");
                return false;
            }

            //if (edtAnalizKatilanSu.Value.Trim() == "")
            //{
            //    Handler.SAPApplication.MessageBox("Analiz bilgileri katılan su alanını kontrol ediniz.");
            //    return false;
            //}
            //else if (Convert.ToDouble(edtAnalizKatilanSu.Value) == 0)
            //{
            //    Handler.SAPApplication.MessageBox("Analiz bilgileri katılan su alanını kontrol ediniz.");
            //    return false;
            //}

            if (edtAnalizSomatikHucre.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri somatik hücre alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizSomatikHucre.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri somatik hücre alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizMetilenMavisi.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri metilen mavisi alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizMetilenMavisi.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri metilen mavisi alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizDonmaNoktasi.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri donma noktası alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizDonmaNoktasi.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri donma noktası alanını kontrol ediniz.");
                return false;
            }

            if (edtAnalizTKM.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri TKM alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAnalizTKM.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri TKM alanını kontrol ediniz.");
                return false;
            }

            if (edtSutTemizligi.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri süt temizliği alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtSutTemizligi.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri süt temizliği alanını kontrol ediniz.");
                return false;
            }

            if (edtTankTemizligi.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri tank temizliği alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtTankTemizligi.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri tank temizliği alanını kontrol ediniz.");
                return false;
            }

            if (edtEkipmanTemizligi.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri ekipman temizliği alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtEkipmanTemizligi.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri ekipman temizliği alanını kontrol ediniz.");
                return false;
            }

            if (edtAracKasasiTemizligi.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri araç kasası temizliği alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracKasasiTemizligi.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Analiz bilgileri araç kasası temizliği alanını kontrol ediniz.");
                return false;
            }

            return true;
        }

        private bool TankVerileriAlanlariniKontroEt()
        {
            if (edtAracTankTankNumarasi.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank verileri tank numarası kontrol ediniz.");
                return false;
            }

            if (edtAracTedarikci.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank verileri tedarikçi alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankLabMiktar.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank verileri LAB miktar alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankLabMiktar.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank verileri LAB miktar alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankDerece.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank verileri sıcaklık alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankDerece.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank verileri sıcaklık alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankBrix.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank verileri briks alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankBrix.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank verileri briks alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankYag.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank yağ alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankYag.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank yağ alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankDans.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank yoğunluk alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankDans.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank yoğunluk alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankpH.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank PH alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankpH.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank PH alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankSH.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank SH alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankSH.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank SH alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankKaynamisSH.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank kaynamış SH alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankKaynamisSH.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank kaynamış SH alanını kontrol ediniz.");
                return false;
            }

            if (oComboAracTankAlkol.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank alkol alanını kontrol ediniz.");
                return false;
            }

            if (oComboAracTankTemizlikDurum.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank temizlik durum alanını kontrol ediniz.");
                return false;
            }

            if (oComboAracTankKoku.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank koku alanını kontrol ediniz.");
                return false;
            }

            if (oComboAracTankSoda.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank koku alanını kontrol ediniz.");
                return false;
            }

            if (oComboAracTankAntibiyotik.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank antibiyotik alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankYagsiz.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank yağsız kuru madde alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankYagsiz.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank yağsız kuru madde alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankProtein.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank protein alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankProtein.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank protein alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankLaktoz.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank laktoz alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankLaktoz.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank laktoz alanını kontrol ediniz.");
                return false;
            }

            //if (edtAracTankSu.Value.Trim() == "")
            //{
            //    Handler.SAPApplication.MessageBox("Araç tank katılan su alanını kontrol ediniz.");
            //    return false;
            //}
            //else if (Convert.ToDouble(edtAracTankSu.Value) == 0)
            //{
            //    Handler.SAPApplication.MessageBox("Araç tank katılan su alanını kontrol ediniz.");
            //    return false;
            //}

            if (edtAracTankSomatikHucre.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank somatik hücre alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankSomatikHucre.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank somatik hücre alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankMetilenMavisi.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank metilen mavisi alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankMetilenMavisi.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank metilen mavisi alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankDonNok.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank donma noktası alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankDonNok.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank donma noktası alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankDonNok.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank donma noktası alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankDonNok.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank donma noktası alanını kontrol ediniz.");
                return false;
            }

            if (edtAracTankTkm.Value.Trim() == "")
            {
                Handler.SAPApplication.MessageBox("Araç tank TKM alanını kontrol ediniz.");
                return false;
            }
            else if (Convert.ToDouble(edtAracTankTkm.Value) == 0)
            {
                Handler.SAPApplication.MessageBox("Araç tank TKM alanını kontrol ediniz.");
                return false;
            }

            return true;
        }

        private string urunKodu = "";
        private string partiNo = "";

        #region SatinAlmaMalGirisiOlustur ADDON içerisinden yapıldığından buradan kaldırılmıştır.
        //public bool SatinAlmaMalGirisiOlustur(string TekilCogulGiris) //1 tekil giriş, 2 Çoğul Giriş anlamına gelir.
        //{
        //    //SAPbobsCOM.Documents oDocuments = (SAPbobsCOM.Documents)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.oPurchaseDeliveryNotes);

        //    string xml = oMatrixAnalizBilgileri.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
        //    var rows = (from x in XDocument.Parse(xml).Descendants("Row")
        //                select new AnalizBilgileri
        //                {
        //                    Tedarikci = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value,
        //                    TedarikciAdi = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_22" select new XElement(y.Element("Value"))).First().Value,
        //                    Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value),
        //                    Sicaklik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value),
        //                    Briks = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_4" select new XElement(y.Element("Value"))).First().Value),
        //                    Yag = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value),
        //                    Yogunluk = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value),
        //                    pH = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value),
        //                    SH = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value),
        //                    SutTuru = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_9" select new XElement(y.Element("Value"))).First().Value,
        //                    //TemizlikKoku = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_10" select new XElement(y.Element("Value"))).First().Value,
        //                    Soda = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_11" select new XElement(y.Element("Value"))).First().Value,
        //                    Antibiyotik = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_12" select new XElement(y.Element("Value"))).First().Value,
        //                    Ykm = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_13" select new XElement(y.Element("Value"))).First().Value),
        //                    Alkol = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_14" select new XElement(y.Element("Value"))).First().Value,
        //                    KaynatSh = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_15" select new XElement(y.Element("Value"))).First().Value),
        //                    Depo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_16" select new XElement(y.Element("Value"))).First().Value,
        //                    Protein = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_17" select new XElement(y.Element("Value"))).First().Value),
        //                    Laktoz = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_18" select new XElement(y.Element("Value"))).First().Value),
        //                    KatilanSu = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_19" select new XElement(y.Element("Value"))).First().Value),
        //                    SomatikHucre = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_20" select new XElement(y.Element("Value"))).First().Value),
        //                    MetilenMavisi = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_21" select new XElement(y.Element("Value"))).First().Value),
        //                    row = Convert.ToString(x.ElementsBeforeSelf().Count() + 1),
        //                    SutTemizlik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_28" select new XElement(y.Element("Value"))).First().Value),
        //                    TankTemizlik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_29" select new XElement(y.Element("Value"))).First().Value),
        //                    EkipmanTemizlik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_30" select new XElement(y.Element("Value"))).First().Value),
        //                    AracKasasiTemizlik = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_31" select new XElement(y.Element("Value"))).First().Value),
        //                    AgirlikliPuan = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_32" select new XElement(y.Element("Value"))).First().Value),
        //                }).ToList();
        //    int i = 0;

        //    OtatServiceSoapClient otatServiceSoapClient = new OtatServiceSoapClient();
        //    PurchaseDeliveryNotes purchaseDeliveryNotes = new PurchaseDeliveryNotes();

        //    //Süt türü 1 ve 2 yani yağlı ve yağsız için toplam 1 kalem üretileceği için gruplamada tek satır gelmesi için süt türü 2 ler 1 olarak güncellendi.
        //    rows.Where(x => x.SutTuru == "2").ToList().ForEach(x => x.SutTuru = "1");

        //    var result = rows.GroupBy(x => x.SutTuru).Select(y => new AnalizBilgileri
        //    {
        //        SutTuru = y.First().SutTuru,
        //        Miktar = y.Sum(z => z.Miktar)
        //    });

        //    var resultwithGroupBy = rows.GroupBy(x => new { x.Tedarikci });

        //    purchaseDeliveryNotes.BPLId = 1;

        //    purchaseDeliveryNotes.DocDate = DateTime.Now.ToString("yyyyMMdd");
        //    //purchaseDeliveryNotes.DocDate = oEditBelgeTarihi.Value;
        //    purchaseDeliveryNotes.userId = ConstVariables.oCompanyObject.UserSignature;
        //    purchaseDeliveryNotes.sutKabulId = Convert.ToInt32(oEditDocEntry.Value);
        //    PurchaseDeliveryNotesDetail purchaseDeliveryNotesDetail = new PurchaseDeliveryNotesDetail();
        //    List<PurchaseDeliveryNotesDetail> purchaseDeliveryNotesDetails = new List<PurchaseDeliveryNotesDetail>();
        //    PurchaseDeliveryNotesDetailParti purchaseDeliveryNotesDetailParti = new PurchaseDeliveryNotesDetailParti();
        //    List<PurchaseDeliveryNotesDetailParti> purchaseDeliveryNotesDetailPartis = new List<PurchaseDeliveryNotesDetailParti>();

        //    if (TekilCogulGiris == "1")
        //    {
        //        purchaseDeliveryNotes.CardCode = oEditSatici.Value;
        //        foreach (var itemAif in result)
        //        {
        //            if (itemAif.SutTuru == "1" || itemAif.SutTuru == "2")
        //            {
        //                purchaseDeliveryNotesDetail.ItemCode = "HAM-00001";
        //            }

        //            urunKodu = purchaseDeliveryNotesDetail.ItemCode;

        //            purchaseDeliveryNotesDetail.WareHouse = "200";
        //            purchaseDeliveryNotesDetail.Quantity = itemAif.Miktar;

        //            purchaseDeliveryNotesDetailParti.BatchNumber = DateTime.Now.Year + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0') + "-" + oEditDocEntry.Value;
        //            purchaseDeliveryNotesDetailParti.Quantity = itemAif.Miktar;

        //            partiNo = purchaseDeliveryNotesDetailParti.BatchNumber;
        //            purchaseDeliveryNotesDetailPartis.Add(purchaseDeliveryNotesDetailParti);

        //            purchaseDeliveryNotesDetail.PartiDetails = purchaseDeliveryNotesDetailPartis.ToArray();

        //            purchaseDeliveryNotesDetails.Add(purchaseDeliveryNotesDetail);
        //        }

        //        purchaseDeliveryNotes.Lines = purchaseDeliveryNotesDetails.ToArray();

        //        if (purchaseDeliveryNotes.Lines.Count() > 0)
        //        {
        //            var resp = otatServiceSoapClient.AddOrUpdatePurchaseDeliveryNotes(purchaseDeliveryNotes, ConstVariables.oCompanyObject.CompanyDB);

        //            if (resp.Value == 0)
        //            {
        //                oEditBelgeNo.Value = resp.DocEntry.ToString();
        //                //oComboBelgeDurumu.Select("C");
        //                oButtonAddOrUpdate.Item.Click();
        //            }
        //            else
        //            {
        //                Handler.SAPApplication.MessageBox("Satınalma belgesi oluşturulurken hata oluştu. " + resp.Description);
        //                return false;
        //            }
        //        }
        //    }
        //    else
        //    {
        //        foreach (var item in resultwithGroupBy)
        //        {
        //            purchaseDeliveryNotes.CardCode = item.Key.Tedarikci.ToString();

        //            var sumQty = rows.Where(x => x.Tedarikci == item.Key.Tedarikci.ToString()).Select(z => z.Miktar).Sum();

        //            purchaseDeliveryNotesDetail.ItemCode = "HAM-00001";

        //            purchaseDeliveryNotesDetail.WareHouse = "200";
        //            purchaseDeliveryNotesDetail.Quantity = sumQty;

        //            purchaseDeliveryNotesDetailParti.BatchNumber = DateTime.Now.Year + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0') + "-" + oEditDocEntry.Value;
        //            purchaseDeliveryNotesDetailParti.Quantity = sumQty;

        //            partiNo = purchaseDeliveryNotesDetailParti.BatchNumber;
        //            purchaseDeliveryNotesDetailPartis.Add(purchaseDeliveryNotesDetailParti);

        //            purchaseDeliveryNotesDetail.PartiDetails = purchaseDeliveryNotesDetailPartis.ToArray();

        //            purchaseDeliveryNotesDetails.Add(purchaseDeliveryNotesDetail);

        //            purchaseDeliveryNotes.Lines = purchaseDeliveryNotesDetails.ToArray();

        //            if (purchaseDeliveryNotes.Lines.Count() > 0)
        //            {
        //                var resp = otatServiceSoapClient.AddOrUpdatePurchaseDeliveryNotes(purchaseDeliveryNotes, ConstVariables.oCompanyObject.CompanyDB);

        //                purchaseDeliveryNotesDetail = new PurchaseDeliveryNotesDetail();
        //                purchaseDeliveryNotesDetails = new List<PurchaseDeliveryNotesDetail>();
        //                purchaseDeliveryNotesDetailParti = new PurchaseDeliveryNotesDetailParti();
        //                purchaseDeliveryNotesDetailPartis = new List<PurchaseDeliveryNotesDetailParti>();

        //                if (resp.Value == 0)
        //                {
        //                    var siraNo = rows.Where(x => x.Tedarikci == item.Key.Tedarikci.ToString()).ToList();
        //                    foreach (var itemAif in siraNo)
        //                    {
        //                        ((SAPbouiCOM.EditText)oMatrixAnalizBilgileri.Columns.Item("Col_23").Cells.Item(Convert.ToInt32(itemAif.row)).Specific).Value = resp.DocEntry.ToString();
        //                    }
        //                    //oEditBelgeNo.Value = resp.DocEntry.ToString();
        //                    //oComboBelgeDurumu.Select("C");
        //                    //oButtonAddOrUpdate.Item.Click();
        //                }
        //                else
        //                {
        //                    Handler.SAPApplication.MessageBox("Satınalma belgeleri oluşturulurken hata oluştu. " + resp.Description);
        //                    return false;
        //                }
        //            }
        //        }

        //        frmSutKabul.Mode = BoFormMode.fm_UPDATE_MODE;
        //        oButtonAddOrUpdate.Item.Click();

        //        if (frmSutKabul.Mode == BoFormMode.fm_UPDATE_MODE)
        //        {
        //            //oComboBelgeDurumu.Select("C");
        //        }
        //    }

        //    return true;

        //    //foreach (var item in result)
        //    //{
        //    //    oDocuments.BPL_IDAssignedToInvoice = 1;
        //    //    oDocuments.CardCode = oEditSatici.Value;
        //    //    oDocuments.DocDate = DateTime.Now;
        //    //    oDocuments.DocDueDate = DateTime.Now;
        //    //    oDocuments.TaxDate = DateTime.Now;

        //    //    if (item.SutTuru == "1")
        //    //    {
        //    //        oDocuments.Lines.ItemCode = "400003";
        //    //    }
        //    //    else if (item.SutTuru == "2")
        //    //    {
        //    //        oDocuments.Lines.ItemCode = "1100";
        //    //    }

        //    //    oDocuments.Lines.Quantity = item.Miktar;

        //    //    if (i != 0)
        //    //    {
        //    //        oDocuments.Lines.BatchNumbers.Add();
        //    //    }
        //    //    oDocuments.Lines.BatchNumbers.SetCurrentLine(i);
        //    //    oDocuments.Lines.BatchNumbers.BatchNumber = DateTime.Now.Year + DateTime.Now.Month.ToString().PadLeft(2, '0') + DateTime.Now.Day.ToString().PadLeft(2, '0') + "-" + oEditDocEntry.Value;
        //    //    oDocuments.Lines.BatchNumbers.Quantity = item.Miktar;

        //    //    i++;
        //    //    oDocuments.Lines.Add();
        //    //}

        //    //var aa = oDocuments.Add();

        //    //var resp = ConstVariables.oCompanyObject.GetLastErrorDescription();

        //    //if (aa == 0)
        //    //{
        //    //    oEditBelgeNo.Value = ConstVariables.oCompanyObject.GetNewObjectKey().ToString();
        //    //    oComboBelgeDurumu.Select("C");
        //    //    oButtonAddOrUpdate.Item.Click();
        //    //}
        //} 
        #endregion

        private void AnalizTemizle(bool comboChange = false)
        {
            try
            {
                frmSutKabul.Freeze(true);

                frmSutKabul.DataSources.UserDataSources.Item("UD_1").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_2").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_3").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_4").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_5").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_6").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_7").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_8").Value = "";

                if (!comboChange)
                {
                    frmSutKabul.DataSources.UserDataSources.Item("UD_9").Value = "0";
                }

                frmSutKabul.DataSources.UserDataSources.Item("UD_10").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_11").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_12").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_13").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_14").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_15").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_16").Value = "200";
                frmSutKabul.DataSources.UserDataSources.Item("UD_17").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_18").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_19").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_20").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_21").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_44").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_46").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_47").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_56").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_57").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_58").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_60").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_61").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_62").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_63").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_64").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_65").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_66").Value = "";

                if (!comboChange)
                {
                    frmSutKabul.DataSources.UserDataSources.Item("UD_9").Value = "1";
                }

                //Ogün beyin isteği üzerinde otomatik gelmeler kaldırıldı. 17/09/2020 17:35

                //frmSutKabul.DataSources.UserDataSources.Item("UD_10").Value = "1";
                //frmSutKabul.DataSources.UserDataSources.Item("UD_11").Value = "2";
                //frmSutKabul.DataSources.UserDataSources.Item("UD_12").Value = "2";
                //frmSutKabul.DataSources.UserDataSources.Item("UD_47").Value = "2";

                //aa
                btnKaydet.Item.Enabled = true;
            }
            catch (Exception)
            {
            }
            finally
            {
                frmSutKabul.Freeze(false);
            }
        }

        private void AracTankTemizle()
        {
            try
            {
                frmSutKabul.Freeze(true);

                frmSutKabul.DataSources.UserDataSources.Item("UD_22").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_23").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_24").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_25").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_26").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_27").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_28").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_29").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_30").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_31").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_32").Value = "";
                //frmSutKabul.DataSources.UserDataSources.Item("UD_33").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_34").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_35").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_36").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_37").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_38").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_39").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_40").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_41").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_42").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_43").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_45").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_53").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_54").Value = "";
                frmSutKabul.DataSources.UserDataSources.Item("UD_59").Value = "";

                //Ogün beyin isteği üzerinde otomatik gelmeler kaldırıldı. 17/09/2020 17:35

                //frmSutKabul.DataSources.UserDataSources.Item("UD_43").Value = "1";
                //frmSutKabul.DataSources.UserDataSources.Item("UD_39").Value = "1";
                //frmSutKabul.DataSources.UserDataSources.Item("UD_31").Value = "2";
                //frmSutKabul.DataSources.UserDataSources.Item("UD_32").Value = "2";
            }
            catch (Exception)
            {
            }
            finally
            {
                frmSutKabul.Freeze(false);
            }
        }

        public void statuKapaliIslemleri()
        {
            edtAnalizSicaklik.Item.Click(); //Her ihtimale karşı statülerle oynadığımız için her zaman açık olan bir tanesine tıkladım.
            oEditBelgeTarihi.Item.Enabled = false;
            oComboBelgeDurumu.Item.Enabled = false;
            edtSutKabulTarihi.Item.Enabled = false;
            edtSutKabulSaati.Item.Enabled = false;
            edtAracPlakasi.Item.Enabled = false;
            edtSurucu.Item.Enabled = false;
            edtFisNo.Item.Enabled = false;
            oChkTekilGiris.Item.Enabled = false;
            oChkCogulGiris.Item.Enabled = false;
            oEditSatici.Item.Enabled = false;
            oComboAnalizSutTuru.Item.Enabled = false;
            edtAnalizTedarikci.Item.Enabled = false;
            edtAnalizDepo.Item.Enabled = false;
            edtAnalizMiktar.Item.Enabled = false;
            edtAracTankTankNumarasi.Item.Enabled = false;
            edtAracTedarikci.Item.Enabled = false;
            edtAracTankLabMiktar.Item.Enabled = false;
            btnKantarAnaveri.Item.Enabled = false;
            oBtnSutKabulIsleminiTamamla.Item.Enabled = false;
            btnKantarVerisiniManuelGir.Item.Enabled = false;
        }

        private void statusAcikOnaylandi()
        {
            edtAnalizSicaklik.Item.Click(); //Her ihtimale karşı statülerle oynadığımız için her zaman açık olan bir tanesine tıkladım.
            oEditBelgeTarihi.Item.Enabled = true;
            oComboBelgeDurumu.Item.Enabled = true;
            edtSutKabulTarihi.Item.Enabled = true;
            edtSutKabulSaati.Item.Enabled = true;
            edtAracPlakasi.Item.Enabled = true;
            edtSurucu.Item.Enabled = true;
            edtFisNo.Item.Enabled = true;
            oChkTekilGiris.Item.Enabled = true;
            oChkCogulGiris.Item.Enabled = true;
            oEditSatici.Item.Enabled = true;
            oComboAnalizSutTuru.Item.Enabled = true;
            edtAnalizTedarikci.Item.Enabled = true;
            edtAnalizDepo.Item.Enabled = true;
            edtAnalizMiktar.Item.Enabled = true;
            edtAracTankTankNumarasi.Item.Enabled = true;
            edtAracTedarikci.Item.Enabled = true;
            edtAracTankLabMiktar.Item.Enabled = true;
            btnKantarAnaveri.Item.Enabled = true;
            oBtnSutKabulIsleminiTamamla.Item.Enabled = true;
            btnKantarVerisiniManuelGir.Item.Enabled = true;

        }

        public void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.MenuUID == "1282" && !pVal.BeforeAction)
                {
                    ekleme = false;
                    KantarVerisiOk = false;
                    oComboBelgeDurumu.Select("O");
                    oEditBelgeTarihi.String = "A";
                    oEditDepo.Value = "200";
                    AnalizTemizle();
                    AracTankTemizle();
                }
                else if (pVal.MenuUID == "1281" && !pVal.BeforeAction)
                {
                    try
                    {
                        frmSutKabul.Mode = BoFormMode.fm_FIND_MODE;
                        oEditDocEntry.Item.Enabled = true;
                    }
                    catch (Exception)
                    {
                    }
                }
                else if (matrixuid == "Item_77")
                {
                    if (pVal.MenuUID == "AIFRGHTCLK_DeleteRow" && pVal.BeforeAction)
                    {
                        int row = oMatrixAnalizBilgileri.GetNextSelectedRow();
                        if (row != -1)
                        {
                            oMatrixAnalizBilgileri.DeleteRow(row);

                            string xml = oMatrixAnalizBilgileri.SerializeAsXML(BoMatrixXmlSelect.mxs_All);
                            var rows = (from x in XDocument.Parse(xml).Descendants("Row")
                                        select new AnalizBilgileri
                                        {
                                            Miktar = parseNumber.parservalues<double>((from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value),
                                        }).ToList();

                            double total = rows.Sum(x => x.Miktar);

                            oedtIrsaliyeNetSutMiktari.Value = total.ToString();
                        }
                    }
                }
                else if (matrixuid == "Item_124")
                {
                    int row = oMatrixAracTankVerileri.GetNextSelectedRow();
                    if (row != -1)
                    {
                        oMatrixAracTankVerileri.DeleteRow(row);
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private string matrixuid = "";

        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\SutKabulLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
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
                    Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_AddRow");
                    return;
                }

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_DeleteRow");
                    Handler.SAPApplication.Menus.RemoveEx("AIFRGHTCLK_AddRow");
                    return;
                }

                matrixuid = eventInfo.ItemUID;
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

        public void SutTuruneGoreListDoldurma(string SutTuru)
        {
            ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //ConstVariables.oRecordset.DoQuery("Select \"DocEntry\" from \"@AIF_SUTKABULUYR\" where \"U_SutTuru\" = '" + oComboAnalizSutTuru.Value.Trim() + "'");
            ConstVariables.oRecordset.DoQuery("Select \"DocEntry\" from \"@AIF_SUTKABULUYR\" where \"U_SutTuru\" = '" + SutTuru + "'");

            var docentry = ConstVariables.oRecordset.Fields.Item(0).Value.ToString();

            ConstVariables.oRecordset.DoQuery("Select * from \"@AIF_SUTKABULUYR1\" where \"DocEntry\" = '" + docentry + "'");

            string xmll = ConstVariables.oRecordset.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData);
            XDocument xDoc = XDocument.Parse(xmll);
            XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
            var rows = (from t in xDoc.Descendants(ns + "Row")
                        select new Models.SutKabulUyarlama
                        {
                            SicaklikMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SicaklikMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            SicaklikMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SicaklikMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            BriksMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_BriksMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            BriksMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_BriksMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            YagMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YagMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            YagMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YagMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            YogunlukMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YogunlukMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            YogunlukMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YogunlukMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            phMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_pHMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            phMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_pHMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            shMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_sHMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            shMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_sHMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            KaynatsHMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_KaynatsHMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            KaynatsHMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_KaynatsHMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AlkolMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_AlkolMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AlkolMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_AlkolMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            TemizlikKokuMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_TemizlikKokuMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            TemizlikKokuMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_TemizlikKokuMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            SodaMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SodaMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            SodaMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SodaMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AntibiyotikMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_AntibiyotikMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AntibiyotikMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_AntibiyotikMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            YkmMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YkmMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            YkmMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YkmMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            ProteinMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_ProteinMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            ProteinMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_ProteinMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            LaktozMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_LaktozMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            LaktozMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_LaktozMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            KatilanSuMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_KatilanSuMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            KatilanSuMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_KatilanSuMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            SomatikHucreMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SomatikHucreMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            SomatikHucreMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SomatikHucreMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            MetilenMaviMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_MetilenMaviMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            MetilenMaviMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_MetilenMaviMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            SutTuru = oComboAnalizSutTuru.Value,
                            DonmaNoktasıMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_DonmaNokMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            DonmaNoktasıMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_DonmaNokMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            TKMMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_TKMMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            TKMMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_TKMMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            SutTemizligiMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SutTmzlkMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            SutTemizligiMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SutTmzlkMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            TankTemizligiMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_TankTmzlkMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            TankTemizligiMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_TankTmzlkMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            EkipmanTemizligiMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_EkipTmlzkMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            EkipmanTemizligiMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_EkipTmlzkMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracKasaTemizligiMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_AracKasaMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracKasaTemizligiMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_AracKasaMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankSicaklikMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SicaklikAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankSicaklikMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SicaklikAracMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankBriksMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_BriksAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankBriksMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_BriksAracMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankYagMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YagAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankYagMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YagAracMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankYogunlukMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YogunlukAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankYogunlukMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YogunlukAracMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankphMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_pHAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankphMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_pHAracMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankshMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_sHAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankshMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_sHAracMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankKaynamisShMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_KaynatsHAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankKaynamisShMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_KaynatsHAracMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankYagsizKuruMaddeMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YkmAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankYagsizKuruMaddeMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_YkmAracMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankProteinMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_ProteinAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankProteinMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_ProteinAracMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankLaktozMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_LaktozAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankLaktozMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_LaktozAracMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankKatilanSuMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_KatilanSuAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankKatilanSuMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_KatilanSuAracMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankSomatikHucreMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SomatikHucreAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankSomatikHucreMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_SomatikHucreAracMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankMetilenMavisiMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_MetilenMaviAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankMetilenMavisiMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_MetilenMaviAracMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankDonmaNoktasiMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_DonmaNokAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankDonmaNoktasiMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_DonmaNokAracMax" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankTKMMin = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_TKMAracMin" select new XElement(y.Element(ns + "Value"))).First().Value),
                            AracTankTKMMax = parseNumber.parservalues<double>((from y in t.Element(ns + "Fields").Elements(ns + "Field") where y.Element(ns + "Alias").Value == "U_TKMAracMax" select new XElement(y.Element(ns + "Value"))).First().Value)

                        }).ToList();

            SutKabulUyarlama = rows;
        }
    }
}