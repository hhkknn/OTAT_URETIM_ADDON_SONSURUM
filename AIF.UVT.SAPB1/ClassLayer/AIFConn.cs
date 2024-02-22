using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class AIFConn
    {
        public static string FormsViewDefault { get { return "AIF.UVT.SAPB1.FormsView."; } }
        public static string OrnekFrm = FormsViewDefault + "OrnekFrm.xml";
        public const string OrnekUID = "Ornek";
        public static Ornek Ornek { get { return Singleton<Ornek>.Instance; } }

        public static string YonetimEkraniFrmXML = FormsViewDefault + "YonetimEkrani.xml";
        public const string YonetimEkraniUID = "Yntm";
        public static YonetimEkrani Yntm { get { return Singleton<YonetimEkrani>.Instance; } }

        public static string SiparisTavsiyeFrmXML = FormsViewDefault + "SiparisTavsiyeParametre.xml";
        public const string SiparisTavsiyeUID = "TAVPARAM";
        public static SiparisTavsiyeParametre TAVPARAM { get { return Singleton<SiparisTavsiyeParametre>.Instance; } }

        public static string SutKabulFrmXML = FormsViewDefault + "SutKabul.xml";
        public const string SutKabulUID = "SutKabul";
        public static SutKabul SutKabul { get { return Singleton<SutKabul>.Instance; } }


        public static string MikrobiyolojikFrmXML = FormsViewDefault + "MikrobiyolojikAnalizEkrani.xml";
        public const string MikrobiyolojikUID = "Mkrbyljk";
        public static MikrobiyolojikAnalizEkrani Mkrbyljk { get { return Singleton<MikrobiyolojikAnalizEkrani>.Instance; } }


        public static string AnalizGirisFrmXML = FormsViewDefault + "AnalizGiris.xml";
        public const string AnalizGirisUID = "AnalizGiris";
        public static AnalizGiris AnalizGiris { get { return Singleton<AnalizGiris>.Instance; } }

        public static string SutKabulIyilestirmeXML = FormsViewDefault + "SutKabulIyilestirme.xml";
        public const string SutKabulIyilestirmeUID = "SutKabulIyl";
        public static SutKabulIyilestirme SutKabulIyl { get { return Singleton<SutKabulIyilestirme>.Instance; } }


        public static string AnalizParametreGirisXML = FormsViewDefault + "AnalizParametreGiris.xml";
        public const string AnalizParametreGirisUID = "AnalizParam";
        public static AnalizParametre AnalizParam { get { return Singleton<AnalizParametre>.Instance; } }

        public static string AktiviteParametreGirisXML = FormsViewDefault + "AktiviteParametreGiris.xml";
        public const string AktiviteParametreGirisUID = "AktiviteParam";
        public static AktiviteParametre AktiviteParam { get { return Singleton<AktiviteParametre>.Instance; } }

        public static string IndirimSihirbaziXML = FormsViewDefault + "IndirimSihirbazi.xml";
        public const string IndirimSihirbaziUID = "IndirimShr";
        public static IndirimSihirbazi IndirimShr { get { return Singleton<IndirimSihirbazi>.Instance; } }

        public static string ReportsXML = FormsViewDefault + "Reports.xml";
        public const string ReportsXMLUID = "Reports";
        public static Reports Reports { get { return Singleton<Reports>.Instance; } }

        public static string KaynakPlanlamaXML = FormsViewDefault + "KaynakPlanlama.xml";
        public const string KaynakPlanlamaXMLUID = "KaynakPlan";
        public static KaynakPlanlama KaynakPlan { get { return Singleton<KaynakPlanlama>.Instance; } }

        public static string UretimSiparisNoSecimXML = FormsViewDefault + "UretimSiparisNoSecim.xml";
        public const string UretimSiparisNoSecimUID = "UrSipNo";
        public static UretimSiparisNoSecim UrSipNo { get { return Singleton<UretimSiparisNoSecim>.Instance; } }

        public static string OperasyonPlaniXML = FormsViewDefault + "OperasyonPlani.xml";
        public const string OperasyonPlaniUID = "OpPlani";
        public static OperasyonPlani OpPlani { get { return Singleton<OperasyonPlani>.Instance; } }


        public static string HaftalikPlanXML = FormsViewDefault + "HaftalikPlanlama.xml";
        public const string HaftalikPlanUID = "HaftaPlan";
        public static HaftalikPlan HaftaPlan { get { return Singleton<HaftalikPlan>.Instance; } }

        public static string TreeViewXML = FormsViewDefault + "TreeViewTest.xml";
        public const string TreeViewUID = "TreeView";
        public static TreeView TreeView { get { return Singleton<TreeView>.Instance; } }
        public static SatisSiparisi Sys139 { get { return Singleton<SatisSiparisi>.Instance; } }
        public const string SatisSiparisi_FormUID = "Sys139";

        public static OzelFiyatlarOlcut Sys669 { get { return Singleton<OzelFiyatlarOlcut>.Instance; } }
        public const string OzelFiyatlar_FormUID = "Sys669";

        public static MuhatapOzelFiyat Sys668 { get { return Singleton<MuhatapOzelFiyat>.Instance; } }
        public const string MuhatapicinOzelFiyatlar_FormUID = "Sys669";


        public static string IndirimSablonXML = FormsViewDefault + "IndirimSablonlari.xml";
        public const string IndirimSablonUID = "IndirimTmp";
        public static IndirimSablonlari IndirimTmp { get { return Singleton<IndirimSablonlari>.Instance; } }

        public static ChooseFromList Sys9999 { get { return Singleton<ChooseFromList>.Instance; } }
        public const string ChooseFromList_FormUID = "Sys9999";


        public static string SutKabulUyarlamaXML = FormsViewDefault + "SutKabulUyarlama.xml";
        public const string SutKabulUyarlamaUID = "SutKabUyrl";
        public static SutKabulUyarlama SutKabUyrl { get { return Singleton<SutKabulUyarlama>.Instance; } }

        public static string GunlukPersonelPlanXML = FormsViewDefault + "GunlukPersonelPlanlama.xml";
        public const string GunlukPersonelPlanUID = "GunPersPl";
        public static GunlukPersonelPlanlama GunPersPl { get { return Singleton<GunlukPersonelPlanlama>.Instance; } }


        public static string GunlukPersonelPlan2XML = FormsViewDefault + "GunlukPersonelPlanlama2.xml";
        public const string GunlukPersonelPlan2UID = "PersPlanG";
        public static GunlukPersonelPlanlama2 PersPlanG { get { return Singleton<GunlukPersonelPlanlama2>.Instance; } }

        public static string GunlukPersonelPlan3XML = FormsViewDefault + "GunlukPersonelPlanlama3.xml";
        public const string GunlukPersonelPlan3UID = "PrsPlan";
        public static GunlukPersonelPlanlama3 PrsPlan { get { return Singleton<GunlukPersonelPlanlama3>.Instance; } }

        public static string BolumlerXML = FormsViewDefault + "Bolumler.xml";
        public const string BolumlerUID = "Bolum";
        public static Bolumler Bolum { get { return Singleton<Bolumler>.Instance; } }

        public static UretimSiparisi Sys65211 { get { return Singleton<UretimSiparisi>.Instance; } }
        public const string UretimSiparisi_FormUID = "Sys65211";

        public static string UretimSiparisPartiUretmeXML = FormsViewDefault + "UretimSiparisPartiUretme.xml";
        public const string UretimSiparisPartiUretmeUID = "UrtSipPart";
        public static UretimSiparisPartiUretme UrtSipPart { get { return Singleton<UretimSiparisPartiUretme>.Instance; } }

        public static string SutTankXML = FormsViewDefault + "SutTankAnalizGiris.xml";
        public const string SutTankUID = "SutTank";
        public static SutTankAnalizGiris SutTank { get { return Singleton<SutTankAnalizGiris>.Instance; } }

        public static string AyranAnalizTakibiXML = FormsViewDefault + "AyranAnalizTakibi.xml";
        public const string AyranAnalizTakibiUID = "AyranAnlz";
        public static AyranAnalizTakibi AyranAnlz { get { return Singleton<AyranAnalizTakibi>.Instance; } }

        public static string TelemeAnalizTakibiXML = FormsViewDefault + "TelemeAnalizTakip.xml";
        public const string TelemeAnalizTakibiUID = "TlmAnalz";
        public static TelemeAnaliz TlmAnalz { get { return Singleton<TelemeAnaliz>.Instance; } }

        public static string SutDepoSecimXML = FormsViewDefault + "SutDepoSecim.xml";
        public const string SutDepoSecimUID = "SutDepo";
        public static SutDepoSecim SutDepo { get { return Singleton<SutDepoSecim>.Instance; } }

        public static string CIPTemizlik1XML = FormsViewDefault + "CIPTemizlik1.xml";
        public const string CIPTemizlik1UID = "CIPTem1";
        public static CIPTemizlik1 CIPTem1 { get { return Singleton<CIPTemizlik1>.Instance; } }

        public static string CIPTemizlik2XML = FormsViewDefault + "CIPTemizlik2.xml";
        public const string CIPTemizlik2UID = "CIPTem2";
        public static CIPTemizlik2 CIPTem2 { get { return Singleton<CIPTemizlik2>.Instance; } }

        public static string AletEkipmanTemizlikXML = FormsViewDefault + "AletEkipmanTemizlik.xml";
        public const string AletEkipmanTemizlikUID = "AETTem";
        public static AletEkipmanTemizlik AETTem { get { return Singleton<AletEkipmanTemizlik>.Instance; } }


        public static string SiparisOnaylamaXML = FormsViewDefault + "SiparisOnaylama.xml";
        public const string SiparisOnaylamaUID = "SipOnay";
        public static SiparisOnaylama SipOnay { get { return Singleton<SiparisOnaylama>.Instance; } }


        public static string GunlukSutRaporuXML = FormsViewDefault + "GunlukSutRaporu.xml";
        public const string GunlukSutRaporuUID = "GnlkSut";
        public static GunlukSutRaporu GnlkSut { get { return Singleton<GunlukSutRaporu>.Instance; } }

        public static string SutPlanlamaXML = FormsViewDefault + "SutPlanlama.xml";
        public const string SutPlanlamaUID = "SutPlan";
        public static SutPlanlama SutPlan { get { return Singleton<SutPlanlama>.Instance; } }

        public static string UretimSiparisiIptalXML = FormsViewDefault + "UretimSiparisiIptal.xml";
        public const string UretimSiparisiIptalUID = "UrtmIpt";
        public static UretimSiparisiIptal UrtmIpt { get { return Singleton<UretimSiparisiIptal>.Instance; } }

        public static string SutPhCezaXML = FormsViewDefault + "SutPhCeza.xml";
        public const string SutPhCezaUID = "PhCeza";
        public static SutPhCeza PhCeza { get { return Singleton<SutPhCeza>.Instance; } }

        public static string SutAntCezaXML = FormsViewDefault + "SutAntibiyotikCeza.xml";
        public const string SutAntCezaUID = "AntCeza";
        public static SutAntibiyotikCeza AntCeza { get { return Singleton<SutAntibiyotikCeza>.Instance; } }

        public static string SutTemizlikCezaXML = FormsViewDefault + "SutTemizlikCeza.xml";
        public const string SutTemizlikCezaUID = "TmzlkCeza";
        public static SutTemizlikCeza TmzlkCeza { get { return Singleton<SutTemizlikCeza>.Instance; } }

        public static string SutBrixCezaXML = FormsViewDefault + "SutBrixCeza.xml";
        public const string SutBrixCezaUID = "BrixCeza";
        public static SutBrixCeza BrixCeza { get { return Singleton<SutBrixCeza>.Instance; } }

        public static string SutBrixPrimXML = FormsViewDefault + "SutBrixPrim.xml";
        public const string SutBrixPrimUID = "BrixPrim";
        public static SutBrixPrim BrixPrim { get { return Singleton<SutBrixPrim>.Instance; } }

        public static string SutYagCezaXML = FormsViewDefault + "SutYagCeza.xml";
        public const string SutYagCezaUID = "YagCeza";
        public static SutYagCeza YagCeza { get { return Singleton<SutYagCeza>.Instance; } }

        public static string SutYagPrimXML = FormsViewDefault + "SutYagPrim.xml";
        public const string SutYagPrimUID = "YagPrim";
        public static SutYagPrim YagPrim { get { return Singleton<SutYagPrim>.Instance; } }

        public static string SutSodaCezaXML = FormsViewDefault + "SutSoda.xml";
        public const string SutSodaCezaUID = "SodaCz";
        public static SutSodaCeza SodaCz { get { return Singleton<SutSodaCeza>.Instance; } }

        public static string SutIyilestirmeXML = FormsViewDefault + "SutIyilestirme.xml";
        public const string SutIyilestirmeUID = "SutIylstr";
        public static SutIyilestirme SutIylstr { get { return Singleton<SutIyilestirme>.Instance; } }


        public static string UretimSiparisiOlusturmaXML = FormsViewDefault + "UretimSiparisiOlusturma.xml";
        public const string UretimSiparisiOlusturmaUID = "sprsOlstr";
        public static UretimSiparisiOlusturma sprsOlstr { get { return Singleton<UretimSiparisiOlusturma>.Instance; } }

        public static string CommarchVarsayilanlariXML = FormsViewDefault + "CommarchVarsayilanlari.xml";
        public const string CommarchVarsayilanlariUID = "CmmrchVrs";
        public static CommarchVarsayilanlari CmmrchVrs { get { return Singleton<CommarchVarsayilanlari>.Instance; } }


        public static string AnalizveriParametreleriXML = FormsViewDefault + "AnalizveriParametreleri.xml";
        public const string AnalizveriParametreleriUID = "AnlzVrPrm";
        public static AnalizveriParametreleri AnlzVrPrm { get { return Singleton<AnalizveriParametreleri>.Instance; } }

        public static string IndirimGirisEkraniFrmXML = FormsViewDefault + "IndirimGirisEkrani.xml";
        public const string IndirimGirisEkraniUID = "IndrmGrs";
        public static IndirimGiris IndrmGrs { get { return Singleton<IndirimGiris>.Instance; } }

        public static string IndirimUrunEkleFrmXML = FormsViewDefault + "IndirimUrunEkle.xml";
        public const string IndirimUrunEkleUID = "UrunEkle";
        public static IndirimUrunEkle UrunEkle { get { return Singleton<IndirimUrunEkle>.Instance; } }

        public static string SatinalmaIskontoGirisEkraniFrmXML = FormsViewDefault + "SatinalmaIskontoGiris.xml";
        public const string SatinalmaIskontoGirisEkraniUID = "StnAlmIsk";
        public static SatinalmaIskontoGiris StnAlmIsk { get { return Singleton<SatinalmaIskontoGiris>.Instance; } }

        public static string SatinalmaIskontoluUrunEkleFrmXML = FormsViewDefault + "SatinalmaIskontoluUrunEkle.xml";
        public const string SatinalmaIskontoluUrunEkleUID = "StnAlmUrn";
        public static SatinalmaIskontoUrunEkle StnAlmUrn { get { return Singleton<SatinalmaIskontoUrunEkle>.Instance; } }
        public static SatinalmaSiparisi Sys142 { get { return Singleton<SatinalmaSiparisi>.Instance; } }
        public const string SatinalmaSiparisi_FormUID = "Sys142";

        public static string SirketBilgieriFrmXML = FormsViewDefault + "SirketBilgileri.xml";
        public const string SirketBilgileriUID = "SrktBlg";
        public static SirketBilgileri SrktBlg { get { return Singleton<SirketBilgileri>.Instance; } }

        public static SAPDuranVarlik Sys1473000075 { get { return Singleton<SAPDuranVarlik>.Instance; } }
        public const string SAPDuranVarlik_FormUID = "Sys1473000075";

        public static string frmDolapTayinFrmXML = FormsViewDefault + "DolapTayin.xml";
        public const string DolapTayinUID = "DolapTayin";
        public static DolapTayin DolapTayin { get { return Singleton<DolapTayin>.Instance; } }

        public static string PartiliUretimRaporuXML = FormsViewDefault + "PartiliUretimRaporu.xml";
        public const string PartiliUretimRaporuUID = "PartiUretRap";
        public static PartiliUretimRaporu PartiUretRap { get { return Singleton<PartiliUretimRaporu>.Instance; } }

        public static string PartiliUretimRaporuSecimXML = FormsViewDefault + "PartiliUretimRaporuSecim.xml";
        public const string PartiliUretimRaporuSecimUID = "UretimSecm";
        public static PartiliUretimRaporuSecim UretimSecm { get { return Singleton<PartiliUretimRaporuSecim>.Instance; } }

        public static string AnalizGirisSecimXML = FormsViewDefault + "AnalizGirisSecim.xml";
        public const string AnalizGirisSecimUID = "AnlzGrsSec";
        public static AnalizGirisSecim AnlzGrsSec { get { return Singleton<AnalizGirisSecim>.Instance; } }

        public static string GirdiKontrolFormuXML = FormsViewDefault + "GirdiKontrolFormu.xml";
        public const string GirdiKontrolFormuUID = "GirdiKntrl";
        public static GirdiKontrolFormu GirdiKntrl { get { return Singleton<GirdiKontrolFormu>.Instance; } }

        public static SAPSatinalmaSiparisliMalGirisi Sys143 { get { return Singleton<SAPSatinalmaSiparisliMalGirisi>.Instance; } }
        public const string SAPSatinalmaSiparisliMalGirisi_FormUID = "Sys143";
        public static SAPSiparisTavsiye Sys65217 { get { return Singleton<SAPSiparisTavsiye>.Instance; } }
        public const string SAPSiparisTavsiye_FormUID = "Sys65217";


        public static string UygunsuzUrunlerXML = FormsViewDefault + "UygunsuzUrunler.xml";
        public const string UygunsuzUrunlerUID = "UygnszUrun";
        public static UygunsuzUrunler UygnszUrun { get { return Singleton<UygunsuzUrunler>.Instance; } }

        public static string MusteriSikayetleriXML = FormsViewDefault + "MusteriSikayetleri.xml";
        public const string MusteriSikayetleriUID = "MustSkayet";
        public static MusteriSikayetleri MustSkayet { get { return Singleton<MusteriSikayetleri>.Instance; } }

        public static SAPIade Sys180 { get { return Singleton<SAPIade>.Instance; } }
        public const string SAPIade_FormUID = "Sys180";

        public static string UrunIadeSecimXML = FormsViewDefault + "UrunIadeSecim.xml";
        public const string UrunIadeSecimUID = "UrnIadeScm";
        public static UrunIadeSecim UrnIadeScm { get { return Singleton<UrunIadeSecim>.Instance; } }

        public static string AnalizGirisBolumlerXML = FormsViewDefault + "AnalizGirisBolumler.xml";
        public const string AnalizGirisBolumlerUID = "AnlzBolum";
        public static AnalizGirisBolumler AnlzBolum { get { return Singleton<AnalizGirisBolumler>.Instance; } }

        public static string AnalizGirisNumuneXML = FormsViewDefault + "AnalizGirisNumune.xml";
        public const string AnalizGirisNumuneUID = "AnlzNumune";
        public static AnalizGirisNumune AnlzNumune { get { return Singleton<AnalizGirisNumune>.Instance; } }

        public static string IadeGrubuXML = FormsViewDefault + "IadeGrubu.xml";
        public const string IadeGrubuUID = "IadeGrup";
        public static IadeGrubu IadeGrup { get { return Singleton<IadeGrubu>.Instance; } }

        public static string IadeGrupNedeniXML = FormsViewDefault + "IadeGrupNedeni.xml";
        public const string IadeGrupNedeniUID = "IadeGrpNdn";
        public static IadeGrupNedeni IadeGrpNdn { get { return Singleton<IadeGrupNedeni>.Instance; } }

        public static string MusteriSikayetGrubuXML = FormsViewDefault + "MusteriSikayetGrubu.xml";
        public const string MusteriSikayetGrubuUID = "MustSkGrb";
        public static MusteriSikayetGrubu MustSkGrb { get { return Singleton<MusteriSikayetGrubu>.Instance; } }

        public static string MusteriSikayetNedeniXML = FormsViewDefault + "MusteriSikayetNedeni.xml";
        public const string MusteriSikayetNedeniUID = "MustSkNdn";
        public static MusteriSikayetNedeni MustSkNdn { get { return Singleton<MusteriSikayetNedeni>.Instance; } }

        public static string GenelParametrelerXML = FormsViewDefault + "GenelParametreler.xml";
        public const string GenelParametrelerUID = "GenelParam";
        public static GenelParametreler GenelParam { get { return Singleton<GenelParametreler>.Instance; } }

        public static string UretimSiparisCogaltXML = FormsViewDefault + "UretimSiparisCogalt.xml";
        public const string UretimSiparisCogaltUID = "UrtSipCog";
        public static UretimSiparisCogalt UrtSipCog { get { return Singleton<UretimSiparisCogalt>.Instance; } }

        public static string SutPlananGerceklesenKarsilastirmaRaporuXML = FormsViewDefault + "SutPlanlananGerceklesenRaporu.xml";
        public const string SutPlananGerceklesenKarsilastirmaRaporuUID = "SutPlnGrck";
        public static SutPlananGerceklesenRaporu SutPlnGrck { get { return Singleton<SutPlananGerceklesenRaporu>.Instance; } }

        public static string KaliteListeDetaylariXML = FormsViewDefault + "KaliteDetayListesi.xml";
        public const string KaliteListeDetaylariuUID = "KltDetay";
        public static KaliteDetayListesi KltDetay { get { return Singleton<KaliteDetayListesi>.Instance; } }

        public static string MIPMiktarDuzenlemeEkraniXML = FormsViewDefault + "MIPMiktarDuzenlemeEkrani.xml";
        public const string MIPMiktarDuzenlemeEkraniUUID = "MIPTV";
        public static MIPMiktarDuzenlemeEkrani MIPTV { get { return Singleton<MIPMiktarDuzenlemeEkrani>.Instance; } }
    }
}
