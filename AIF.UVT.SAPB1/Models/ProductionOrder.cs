using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIF.UVT.SAPB1.Models
{
    public class ProductionOrder
    {
        public string ItemCode { get; set; }
        public double Miktar { get; set; }
        public string Depo { get; set; }
    }
}
