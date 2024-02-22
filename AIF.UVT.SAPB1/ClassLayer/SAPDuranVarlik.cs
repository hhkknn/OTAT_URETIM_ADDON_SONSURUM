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
    public class SAPDuranVarlik
    {
        [ItemAtt(AIFConn.SAPDuranVarlik_FormUID)]
        public SAPbouiCOM.Form frmSAPDuranVarlik;

        //[ItemAtt("Item_10")]
        //public SAPbouiCOM.EditText EdtDocEntry;

        //[ItemAtt("Item_0")] 
        static string formuid = "";
        private SAPbouiCOM.Folder oFolder = null;
        public void LoadForms()
        {
            if (Program.mKod == "010OTATURVT")
            {
                try
                {
                    Functions.CreateUserOrSystemFormComponent<SAPDuranVarlik>(AIFConn.Sys1473000075, true, formuid);

                    System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();
                    System.IO.Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("AIF.UVT.SAPB1.FormsView.SAPDuranVarlik.xml");

                    System.IO.StreamReader streamreader = new System.IO.StreamReader(stream, true);
                    xmldoc.LoadXml(string.Format(streamreader.ReadToEnd(), formuid));
                    Handler.SAPApplication.LoadBatchActions(xmldoc.InnerXml);

                    streamreader.Close();
                }
                catch (Exception ex)
                {
                    //Handler.SAPApplication.MessageBox("hata:" + ex.Message);

                }

                InitForms();
            }
        }
        public void InitForms()
        {
            try
            {

                frmSAPDuranVarlik.Items.Item("Item_0").Top = frmSAPDuranVarlik.Items.Item("2").Top;
                frmSAPDuranVarlik.Items.Item("Item_0").Height = frmSAPDuranVarlik.Items.Item("2").Height;
                frmSAPDuranVarlik.Items.Item("Item_0").Left = frmSAPDuranVarlik.Items.Item("2").Left + 85;
                frmSAPDuranVarlik.Items.Item("Item_0").LinkTo = "2";
                //oFolder = (SAPbouiCOM.Folder)frmSAPDuranVarlik.Items.Item("Item_0").Specific;
                //oFolder.Pane = 99;
                //oFolder.Item.Left = frmSAPDuranVarlik.Items.Item("1470002136").Left;
                //oFolder.GroupWith("1470002136");
                //oFolder.Item.AffectsFormMode = false;
                //oFolder.Item.Visible = false;
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Hata oluştu." + ex.Message);
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
                    if (Program.mKod == "010OTATURVT")
                    {
                        if (!BusinessObjectInfo.BeforeAction)
                        {
                            try
                            {
                                //oFolder.Item.Visible = true;

                                //string itemCode = ((SAPbouiCOM.EditText)frmKalemAnaverileri.Items.Item("5").Specific).Value;

                                //kayitlariListele(itemCode);

                                //frmKalemAnaverileri.Mode = BoFormMode.fm_OK_MODE;

                                //silinmisler = new List<string>();
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }
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
                    if (Program.mKod == "010OTATURVT")
                    {
                        //if (pVal.ItemUID == "Item_0" && !pVal.BeforeAction)
                        //{
                        //    try
                        //    {
                        //        oFolder.Pane = 99;
                        //        oFolder.Item.Left = frmSAPDuranVarlik.Items.Item("1470002136").Left;

                        //        frmSAPDuranVarlik.PaneLevel = 99;
                        //    }
                        //    catch (Exception)
                        //    {
                        //    }
                        //}
                        if (pVal.ItemUID == "Item_0" && !pVal.BeforeAction)
                        {
                            try
                            {
                                string kalemKodu = ((SAPbouiCOM.EditText)frmSAPDuranVarlik.Items.Item("5").Specific).Value.ToString();
                                AIFConn.DolapTayin.LoadForms(kalemKodu);
                            }
                            catch (Exception)
                            {

                            }
                        }
                    }
                    break;
                case BoEventTypes.et_DOUBLE_CLICK:
                    //if (pVal.ColUID == "15" && !pVal.BeforeAction)
                    //{
                    //    //AIFConn.IndirimShr.LoadForms(frmSatisSiparisi, "38", pVal.Row);
                    //}
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
                    if (pVal.BeforeAction)
                    {
                        frmSAPDuranVarlik = Handler.SAPApplication.Forms.Item(pVal.FormUID);
                        //frmInit = false; 
                        formuid = pVal.FormUID;
                        AIFConn.Sys1473000075.LoadForms();
                        //SAPDuranVarlik.baseForm = frmSAPDuranVarlik;
                    }
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

                    if (Program.mKod == "010OTATURVT")
                    {
                        //try
                        //{
                        //    if (frmSAPDuranVarlik.Mode != BoFormMode.fm_FIND_MODE)
                        //    {
                        //        //if (oMatrixKonfigurator != null)
                        //        //{
                        //        //oMatrixKonfigurator.AutoResizeColumns();
                        //        oFolder.Item.Left = frmSAPDuranVarlik.Items.Item("1470002136").Left;
                        //        oFolder.GroupWith("1470002136");
                        //        //}
                        //    }
                        //}
                        //catch (Exception)
                        //{
                        //} 
                    }
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
