using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIF.UVT.SAPB1.Models
{
    public class AnalizPartiler
    {
        public string Sec { get; set; }
        public string PartiNumarasi { get; set; }
        public string KabulTarihi { get; set; }
        public string GecerlilikSonu { get; set; }
        public string AnalizAdi { get; set; }
        public int SatirNo { get; set; }
    }

    public class AnalizSecilmisPartiler
    {
        public string PartiNumarasi { get; set; }

    }
}
