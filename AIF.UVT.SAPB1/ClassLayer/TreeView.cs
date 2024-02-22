using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class TreeView
    {
        [ItemAtt(AIFConn.TreeViewUID)]
        public SAPbouiCOM.Form frmTreeView;

        [ItemAtt("Item_1")]
        public SAPbouiCOM.Matrix oMatrix;
        //[ItemAtt("1")]
        //public SAPbouiCOM.Button btnAddOrUpdate;
        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.TreeViewXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.TreeViewXML));
            Functions.CreateUserOrSystemFormComponent<TreeView>(AIFConn.TreeView);

            InitForms();
        }

        public class Datas
        {
            public string Seviye { get; set; }
            public string Grup { get; set; }
            public string Tur { get; set; }
            public string Kod { get; set; }
            public string Ad { get; set; }
        }
        List<Datas> datas = new List<Datas>();
        public void InitForms()
        {
            try
            {
                SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRSSeviye2 = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRSOItem = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


                oRS.DoQuery("Select * from \"@AIF_SEVIYE1\"");
                oRSSeviye2.DoQuery("Select * from \"@AIF_SEVIYE2\"");
                while (!oRS.EoF)
                {
                    datas.Add(new Datas { Seviye = oRS.Fields.Item("U_Code").Value.ToString(), Grup = oRS.Fields.Item("U_Name").Value.ToString() });

                    while (!oRSSeviye2.EoF)
                    {
                        datas.Add(new Datas { Seviye = oRSSeviye2.Fields.Item("U_Code").Value.ToString(), Grup = "", Tur = oRSSeviye2.Fields.Item("U_Name").Value.ToString() });

                        if (oRSSeviye2.Fields.Item("U_Name").Value.ToString() == "YARIMAMÜL" || oRSSeviye2.Fields.Item("U_Name").Value.ToString() == "MAMÜL")
                        {
                            string sql = "Select \"ItemCode\",\"ItemName\" from \"OITM\" where \"U_SEVIYE1\" = '" + oRS.Fields.Item("DocEntry").Value.ToString() + "' and \"U_SEVIYE2\" = '" + oRSSeviye2.Fields.Item("DocEntry").Value.ToString() + "' ";
                            oRSOItem.DoQuery("Select \"ItemCode\",\"ItemName\" from \"OITM\" where \"U_SEVIYE1\" = '" + oRS.Fields.Item("DocEntry").Value.ToString() + "' and \"U_SEVIYE2\" = '" + oRSSeviye2.Fields.Item("DocEntry").Value.ToString() + "' ");

                            while (!oRSOItem.EoF)
                            {
                                datas.Add(new Datas { Seviye = "3", Grup = "", Tur = "", Kod = oRSOItem.Fields.Item("ItemCode").Value.ToString(), Ad = oRSOItem.Fields.Item("ItemName").Value.ToString() });
                                oRSOItem.MoveNext();
                            }
                        }
                        oRSSeviye2.MoveNext();
                    }

                    oRSSeviye2.MoveFirst();

                    oRS.MoveNext();
                }


                foreach (var item in datas)
                {
                    oMatrix.AddRow();
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_7").Cells.Item(oMatrix.RowCount).Specific).Value = item.Seviye;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(oMatrix.RowCount).Specific).Value = item.Grup;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_1").Cells.Item(oMatrix.RowCount).Specific).Value = item.Tur;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(oMatrix.RowCount).Specific).Value = item.Kod;
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_3").Cells.Item(oMatrix.RowCount).Specific).Value = item.Ad;
                }


                oMatrix.AutoResizeColumns();



            }
            catch (Exception ex)
            {
            }
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

        string val = "";
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
        public void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

        }
    }
}
