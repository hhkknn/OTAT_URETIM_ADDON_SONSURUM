using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class SirketBilgileri : IUserForm
    {
        [ItemAtt(AIFConn.SirketBilgileriUID)]
        public SAPbouiCOM.Form frmSirketBilgileri;
        [ItemAtt("Item_24")]
        public SAPbouiCOM.EditText EdtDocEntry;

        [ItemAtt("1")]
        public SAPbouiCOM.Button btnAddOrUpdate;
        [ItemAtt("Item_25")]
        public SAPbouiCOM.EditText EdtSirketKodu;
        [ItemAtt("Item_2")]
        public SAPbouiCOM.EditText EdtSirketAdi;
        [ItemAtt("Item_19")]
        public SAPbouiCOM.ComboBox CmbServerTipi;

        [ItemAtt("Item_6")]
        public SAPbouiCOM.EditText EdtLisansServer;

        [ItemAtt("Item_8")]
        public SAPbouiCOM.EditText EdtServer;

        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.SirketBilgieriFrmXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.SirketBilgieriFrmXML));
            Functions.CreateUserOrSystemFormComponent<SirketBilgileri>(AIFConn.SrktBlg);

            InitForms();
        }

        public void InitForms()
        {
            int gen = frmSirketBilgileri.Width;
            int yuk = frmSirketBilgileri.Height;
            ConstVariables.oRecordset = (Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
            ConstVariables.oRecordset.DoQuery("Select TOP 1 \"DocEntry\" from \"@AIF_UVT_CONSTRNG\" as T0 where T0.\"U_CompanyDBCode\" = '" + ConstVariables.oCompanyObject.CompanyDB + "'");

            if (ConstVariables.oRecordset.RecordCount > 0)
            {
                frmSirketBilgileri.Mode = BoFormMode.fm_FIND_MODE;
                EdtDocEntry.Value = ConstVariables.oRecordset.Fields.Item(0).Value.ToString();
                btnAddOrUpdate.Item.Click();
            }
            else
            {
                EdtSirketKodu.Value = ConstVariables.oCompanyObject.CompanyDB;
                EdtSirketAdi.Value = ConstVariables.oCompanyObject.CompanyName;
                EdtLisansServer.Value = ConstVariables.oCompanyObject.LicenseServer;
                EdtServer.Value = ConstVariables.oCompanyObject.Server;
                //var dbtype = ConstVariables.oCompanyObject.DbServerType;
                //CmbServerTipi.ValidValues.Add(ConstVariables.oCompanyObject.DbServerType.ToString(),ConstVariables.oCompanyObject.DbServerType.ToString());
                //CmbServerTipi.Select(dbtype, BoSearchKey.psk_ByDescription);
            }

            CmbServerTipi.ExpandType = BoExpandType.et_DescriptionOnly;
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
    }
}
