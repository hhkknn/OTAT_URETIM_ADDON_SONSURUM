using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MailConfirmation.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index(string docEntry, string onayTipi)
        {
            //docEntry = "585";
            //onayTipi = "1";
            //guid = "f0c4d306-4dea-4c69-b118-5abf3ee287f2";
            if (docEntry == null || docEntry == "")
            {
                ViewBag.Info = "Benzersiz numara gönderimi olmadan işlem yapılamaz.";
                return View("Info");
            }

            ViewBag.onayTipi = onayTipi;
            //ViewData["GuidID"] = guid.ToString();
            int _docEntry = Convert.ToInt32(docEntry);
            var aif = new OTATEntities1();
            var liste = aif.OPORs.Where(x => x.DocEntry == _docEntry).ToList();

            if (liste.Count > 0)
            {
                if (liste[0].U_OnayRed != null && liste[0].U_OnayRed != "B")
                {
                    string zaman = Convert.ToDateTime(liste[0].U_YonIslmTar).ToString("dd/MM/yyyy") + " " + liste[0].U_YonIslmSaat;
                    ViewBag.Info = "Kaydınız " + zaman + " tarihinde " + liste[0].U_OnaylayanYon + " tarafından değerlendirilmiştir.";
                    return View("Info");
                }
            }
            else
            {
                ViewBag.Info = "Satınalma belgesi bulunamadı.";
                return View("Info");
            }
            ViewBag.SatinalmaVerileri = liste;
            var satinalmaverileriDetay = aif.POR1s.Where(x => x.DocEntry == _docEntry).FirstOrDefault();
            ViewBag.SatisVerileri = aif.ORDRs.Where(x => x.DocEntry == satinalmaverileriDetay.BaseEntry).ToList();
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpPost]
        public ActionResult Confirm(string guid, string onaylayan, string cevap, string aciklamalar)
        {

            //var ctx = new OTATTEST18Entities();
            //var datas = ctx.C_AIF_GONDMUTMUH.Where(x => x.U_GuidId == guid).ToList();

            //if (datas[0].U_Onay != null && datas[0].U_Onay != "")
            //{
            //    ViewBag.Info = "Kaydınız " + datas[0].U_OnayTarihi + " tarihinde " + datas[0].U_Onaylayan + " tarafından değerlendirilmiştir.";
            //}
            //else
            //{
            //    datas[0].U_Onaylayan = onaylayan;
            //    datas[0].U_Onay = cevap == "1" ? "Y" : "N";
            //    datas[0].U_Aciklama = aciklamalar;
            //    datas[0].U_OnayTarihi = DateTime.Now;

            //    ctx.SaveChanges();

            //    ViewBag.Info = "Kaydınız başarı ile alınmıştır.";
            //}

            return View("Info");
        }

        [HttpPost]
        public ActionResult ConfirmSatinalma(string docEntry, string onaylayan, string cevap, string aciklamalar)
        {
            //ViewBag.Info = "Kaydınız başarı ile alınmıştır.";

            //return View();


            var ctx = new OTATEntities1();
            int _docEntry = Convert.ToInt32(docEntry);
            var datas = ctx.OPORs.Where(x => x.DocEntry == _docEntry).ToList();

            if (datas != null)
            {
                datas[0].U_OnaylayanYon = onaylayan;
                datas[0].U_OnayRed = cevap == "1" ? "O" : "R";
                datas[0].U_YonIslmTar = DateTime.Now;
                datas[0].U_YonIslmSaat = DateTime.Now.Hour.ToString().PadLeft(2, '0') + ":" + DateTime.Now.Minute.ToString().PadLeft(2, '0');
                datas[0].Comments = datas[0].Comments + " " + aciklamalar;

                ctx.SaveChanges();

                ViewBag.Info = "Kaydınız başarı ile alınmıştır.";
            }
            //if (datas[0].U_Onay != null && datas[0].U_Onay != "")
            //{
            //    ViewBag.Info = "Kaydınız " + datas[0].U_OnayTarihi + " tarihinde " + datas[0].U_Onaylayan + " tarafından değerlendirilmiştir.";
            //}
            //else
            //{
            //    datas[0].U_Onaylayan = onaylayan;
            //    datas[0].U_Onay = cevap == "1" ? "Y" : "N";
            //    datas[0].U_Aciklama = aciklamalar;
            //    datas[0].U_OnayTarihi = DateTime.Now;

            //    ctx.SaveChanges();

            //    ViewBag.Info = "Kaydınız başarı ile alınmıştır.";
            //}

            return View("Info");
        }
    }
}