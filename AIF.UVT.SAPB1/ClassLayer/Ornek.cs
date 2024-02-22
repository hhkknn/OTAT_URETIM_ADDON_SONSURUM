using AIF.ObjectsDLL;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class Ornek
    { 
        public void LoadForms()
        { 
        }

        public void InitForms()
        {
            throw new NotImplementedException();
        }
        public bool SAP_FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            throw new NotImplementedException();
        }

        public bool SAP_ItemEvent(string FormUID, ref ItemEvent pVal, ref bool BubbleEvent)
        {
            throw new NotImplementedException();
        }
    }
}
