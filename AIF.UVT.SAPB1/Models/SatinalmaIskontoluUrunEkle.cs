using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIF.UVT.SAPB1.Models
{ 
    public class SatinalmaIskontolu_UrunEkle
    {
        public string Sec { get; set; }

        public string UrunKodu { get; set; }

        public string UrunTanimi { get; set; }

        public string UrunGrupKod { get; set; }

        public string UrunGrupAdi { get; set; }

        public string MiktarOlcegi { get; set; }

        public string KalemGrubu2Kodu { get; set; }

        public string KalemGrubu2Adi { get; set; }

        public string KalemGrubu3Kodu { get; set; }

        public string KalemGrubu3Adi { get; set; }

        public string KalemGrubu4Kodu { get; set; }

        public string KalemGrubu4Adi { get; set; }

        public string birinciIskonto { get; set; }

        public string ikinciIskonto { get; set; }

        public string ucuncuIskonto { get; set; }

        public string dorduncuIskonto { get; set; }

        public string besinciIskonto { get; set; }

        public string toplamIskonto { get; set; }

        public string birimFiyat { get; set; }

        public string fiyat { get; set; }

        public string dahaOnceEkli { get; set; }

        public int Sira { get; set; }
    }
}
