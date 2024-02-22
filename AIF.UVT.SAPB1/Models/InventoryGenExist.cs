﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIF.UVT.SAPB1.Models
{
    public class InventoryGenExist
    {
        public int UretimSiparisi { get; set; }
        public int SatirNumarasi { get; set; }
        public double Miktar { get; set; }
        public List<Parti> Parti { get; set; }
        public string RotaKodu { get; set; }
        public string PartiNo { get; set; }
        public string DepoKodu { get; set; }
    }

    public class Parti
    {
        public string PartiNo { get; set; }
        public double PartikMiktar { get; set; }
    }
}
