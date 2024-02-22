using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class UretimSiparisNoSecim
    {
        [ItemAtt(AIFConn.UretimSiparisNoSecimUID)]
        public SAPbouiCOM.Form frmUretimSiparisNoSecim;

        [ItemAtt("Item_1")]
        public SAPbouiCOM.EditText EdtDocEntry;
        //[ItemAtt("1")]
        //public SAPbouiCOM.Button btnAddOrUpdate;
        [ItemAtt("Item_6")]
        public SAPbouiCOM.EditText EdtStartDate;
        [ItemAtt("Item_7")]
        public SAPbouiCOM.EditText EdtEndDate;
        SAPbouiCOM.Form baseForm = null;
        int baserow = 0;
        public void LoadForms(SAPbouiCOM.Form _baseForm, string start, string end, int _baserow)
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.UretimSiparisNoSecimXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.UretimSiparisNoSecimXML));
            Functions.CreateUserOrSystemFormComponent<UretimSiparisNoSecim>(AIFConn.UrSipNo);
            baseForm = _baseForm;
            EdtStartDate.Value = start.Trim();
            EdtEndDate.Value = end.Trim();
            baserow = _baserow;
            InitForms();
        }
        public void InitForms()
        {
            try
            {
                int colnum = 3;
                Saatler = new List<Tuple<int, TimeSpan, TimeSpan>>();
                for (int i = 0; i < 24; i++)
                {
                    TimeSpan tsStart = new TimeSpan(i, 00, 00);

                    TimeSpan tsEnd = new TimeSpan(i, 30, 00);


                    Saatler.Add(Tuple.Create(colnum, tsStart, tsEnd));

                    colnum++;

                    tsStart = new TimeSpan(i + 1, 00, 00);

                    if (i == 23)
                    {
                        tsStart = new TimeSpan(0, 0, 0);
                    }
                    Saatler.Add(Tuple.Create(colnum, tsEnd, tsStart));

                    colnum++;
                }

            }
            catch (Exception ex)
            {
            }
        }
        List<Tuple<int, TimeSpan, TimeSpan>> Saatler = new List<Tuple<int, TimeSpan, TimeSpan>>();
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
                    if (pVal.ItemUID == "Item_2" && !pVal.BeforeAction)
                    {
                        List<int> row = new List<int>();
                        var start = EdtStartDate.Value.Split(':');
                        var end = EdtEndDate.Value.Split(':');


                        TimeSpan tsStart = new TimeSpan(Convert.ToInt32(start[0]), Convert.ToInt32(start[1]), 0);
                        TimeSpan tsEnd = new TimeSpan(Convert.ToInt32(end[0]), Convert.ToInt32(end[1]), 0);


                        var list = Saatler.Where(x => x.Item2 <= tsStart).ToList();

                        var list2 = Saatler.Where(x => x.Item3 > tsEnd).ToList();


                        int baslangickolonu = list[list.Count - 1].Item1;
                        int bitiskolonu = list2[0].Item1;

                        SAPbouiCOM.Matrix oBaseMatrix = (SAPbouiCOM.Matrix)baseForm.Items.Item("Item_0").Specific;
                        int color = ColorTranslator.ToOle(System.Drawing.Color.Yellow);

                        for (int i = baslangickolonu; i <= bitiskolonu; i++)
                        {
                            ((SAPbouiCOM.EditText)oBaseMatrix.Columns.Item(i).Cells.Item(baserow).Specific).Value = EdtDocEntry.Value;
                            oBaseMatrix.CommonSetting.SetCellBackColor(baserow, i, color);
                            //((SAPbouiCOM.EditText)oBaseMatrix.Columns.Item(i).Cells.Item(baserow).Specific).BackColor = color;
                        }

                        try
                        {
                            frmUretimSiparisNoSecim.Close();
                        }
                        catch (Exception)
                        {
                        }
                    }
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
                    if (pVal.ItemUID == "Item_1" && pVal.BeforeAction)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = default(SAPbouiCOM.IChooseFromListEvent);
                        oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;


                        SAPbouiCOM.ChooseFromList oCFL = default(SAPbouiCOM.ChooseFromList);
                        oCFL = frmUretimSiparisNoSecim.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

                        SAPbouiCOM.Conditions oCons = default(SAPbouiCOM.Conditions);
                        SAPbouiCOM.Condition oCon = default(SAPbouiCOM.Condition);
                        SAPbouiCOM.Conditions oEmptyConts = new SAPbouiCOM.Conditions();

                        oCFL.SetConditions(oEmptyConts);
                        oCons = oCFL.GetConditions();

                        oCon = oCons.Add();
                        oCon.Alias = "Status";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "P";


                        oCFL.SetConditions(oCons);
                    }
                    else if (pVal.ItemUID == "Item_1" && !pVal.BeforeAction)
                    {

                        try
                        {
                            SAPbouiCOM.DataTable oDataTable = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;
                            string DocEntry = "";
                            DocEntry = oDataTable.GetValue("DocEntry", 0).ToString();
                            EdtDocEntry.Value = DocEntry;
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
