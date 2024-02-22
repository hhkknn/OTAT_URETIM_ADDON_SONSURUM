using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class MIPMiktarDuzenlemeEkrani
    {
        [ItemAtt(AIFConn.MIPMiktarDuzenlemeEkraniUUID)]
        public SAPbouiCOM.Form frmMIPMiktarDuzenlemeEkrani;

        [ItemAtt("Item_1")]
        public SAPbouiCOM.EditText EdtBasTarih;
        [ItemAtt("Item_3")]
        public SAPbouiCOM.EditText EdtBitTarih;
        [ItemAtt("Item_4")]
        public SAPbouiCOM.Button BtnListele;
        [ItemAtt("Item_5")]
        public SAPbouiCOM.Grid oGrid;
        [ItemAtt("Item_6")]
        public SAPbouiCOM.Button BtnStokNakliTalebi;

        [ItemAtt("Item_7")]
        public SAPbouiCOM.Matrix oMatrix;

        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.MIPMiktarDuzenlemeEkraniXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.MIPMiktarDuzenlemeEkraniXML));
            Functions.CreateUserOrSystemFormComponent<MIPMiktarDuzenlemeEkrani>(AIFConn.MIPTV);

            InitForms();
        }
        public void InitForms()
        {
            try
            {
                frmMIPMiktarDuzenlemeEkrani.EnableMenu("1283", false);
                frmMIPMiktarDuzenlemeEkrani.EnableMenu("1284", false);
                frmMIPMiktarDuzenlemeEkrani.EnableMenu("1286", false);

                oDT = frmMIPMiktarDuzenlemeEkrani.DataSources.DataTables.Add("DATA");
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
        SAPbouiCOM.DataTable oDT = null;
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
                    if (pVal.ItemUID == "Item_4" && !pVal.BeforeAction)
                    {
                        #region MyRegion
                        string BaslangicTarihi = EdtBasTarih.Value.ToString();
                        string BitisTarihi = EdtBitTarih.Value.ToString(); ;
                        //string Sorgu = " Select T1.ItemCode as 'Ürün Kodu',T1.Dscription as 'Ürün Tanımı',T1.Quantity as 'Miktar' ,cast(T2.U_StkNkOran as float) as 'Oran',T1.OpenQty + ((T1.OpenQty * cast(T2.U_StkNkOran as float)) / 100) as 'Revize Edilecek Miktar',(T1.OpenQty + ((T1.OpenQty * cast(T2.U_StkNkOran as float)) / 100)) - T1.OpenQty as 'Fark',T0.DocEntry as 'Stok Nakli Talep No',T1.\"LineNum\",T1.\"FromWhsCod\",T1.\"WhsCode\" from OWTQ T0 inner join WTQ1 T1 on T0.DocEntry = T1.DocEntry left join OITM T2 on T1.ItemCode = T2.ItemCode where T0.DocDate between '" + BaslangicTarihi + "' and '" + BitisTarihi + "' and cast(ISNULL(T2.U_StkNkOran,'0') as float) > 0 and T1.OpenQty > 0   and ISNULL(T1.U_RevizeMi,'N')='N' ";

                        //oDT.Clear();
                        //oDT.ExecuteQuery(Sorgu);

                        ////oMatrix.Columns.Item("Col_0").DataBind.Bind("DATA", "Sec");
                        //oMatrix.Columns.Item("Col_1").DataBind.Bind("DATA", "Ürün Kodu");
                        //oMatrix.Columns.Item("Col_2").DataBind.Bind("DATA", "Ürün Tanımı");
                        //oMatrix.Columns.Item("Col_3").DataBind.Bind("DATA", "Miktar");
                        //oMatrix.Columns.Item("Col_4").DataBind.Bind("DATA", "Oran");
                        //oMatrix.Columns.Item("Col_5").DataBind.Bind("DATA", "Revize Edilecek Miktar");
                        //oMatrix.Columns.Item("Col_6").DataBind.Bind("DATA", "Fark");
                        //oMatrix.Columns.Item("Col_7").DataBind.Bind("DATA", "Stok Nakli Talep No");
                        //oMatrix.Columns.Item("Col_8").DataBind.Bind("DATA", "LineNum");
                        //oMatrix.Columns.Item("Col_9").DataBind.Bind("DATA", "FromWhsCod");
                        //oMatrix.Columns.Item("Col_10").DataBind.Bind("DATA", "WhsCode");


                        //oMatrix.LoadFromDataSource();
                        //oMatrix.AutoResizeColumns();
                        //oMatrix.AutoResizeColumns();

                        Listele(BaslangicTarihi, BitisTarihi);
                        #region Grid
                        //oGrid.DataTable.ExecuteQuery(Sorgu);

                        //EditTextColumn UrunKodu = ((SAPbouiCOM.EditTextColumn)(oGrid.Columns.Item("Ürün Kodu")));

                        //UrunKodu.LinkedObjectType = "4";

                        //EditTextColumn StokNakliTalepNo = ((SAPbouiCOM.EditTextColumn)(oGrid.Columns.Item("Stok Nakli Talep No")));

                        //StokNakliTalepNo.LinkedObjectType = "1250000001";

                        //EditTextColumn oran = ((SAPbouiCOM.EditTextColumn)(oGrid.Columns.Item("oran")));
                        //oran.Visible = false;
                        //EditTextColumn LineNum = ((SAPbouiCOM.EditTextColumn)(oGrid.Columns.Item("LineNum")));
                        //LineNum.Width = 0;

                        //EditTextColumn FromWhsCod = ((SAPbouiCOM.EditTextColumn)(oGrid.Columns.Item("FromWhsCod")));
                        //FromWhsCod.Width = 0;

                        //EditTextColumn WhsCode = ((SAPbouiCOM.EditTextColumn)(oGrid.Columns.Item("WhsCode")));
                        //WhsCode.Width = 0;

                        //oGrid.AutoResizeColumns(); 
                        #endregion
                        #endregion

                        //Listele();
                    }
                    else if (pVal.ItemUID == "Item_6" && pVal.BeforeAction)
                    {

                        string xmlmatrix = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                        var rows = (from x in XDocument.Parse(xmlmatrix).Descendants("Row")
                                    where (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value == "Y"
                                    select new
                                    {
                                        UrunKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value.ToString(),
                                        UrunTanim = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value.ToString(),
                                        Miktar = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value.ToString(),
                                        RevizeMiktar = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value.ToString(),
                                        Fark = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value.ToString(),
                                        StokNakliTalepNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value.ToString(),
                                        LineNum = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value.ToString(),


                                    }).ToList();

                        if (rows.Count == 0)//(oGrid.Rows.SelectedRows.Count == 0)
                        {
                            Handler.SAPApplication.MessageBox("Stok nakli talebi miktar revizesi yapabilmeniz için satır seçmelisiniz");
                            BubbleEvent = false;
                            break;
                        }
                    }
                    else if (pVal.ItemUID == "Item_6" && !pVal.BeforeAction)
                    {
                        string BaslangicTarihi = EdtBasTarih.Value.ToString();
                        string BitisTarihi = EdtBitTarih.Value.ToString();
                        try
                        {


                            string xmlmatrix = oMatrix.SerializeAsXML(BoMatrixXmlSelect.mxs_All);

                            var rows = (from x in XDocument.Parse(xmlmatrix).Descendants("Row")
                                        where (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_0" select new XElement(y.Element("Value"))).First().Value == "Y"
                                        select new
                                        {
                                            UrunKodu = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_1" select new XElement(y.Element("Value"))).First().Value.ToString(),
                                            UrunTanim = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_2" select new XElement(y.Element("Value"))).First().Value.ToString(),
                                            Miktar = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_3" select new XElement(y.Element("Value"))).First().Value.ToString(),
                                            RevizeMiktar = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_5" select new XElement(y.Element("Value"))).First().Value.ToString(),
                                            Fark = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_6" select new XElement(y.Element("Value"))).First().Value.ToString(),
                                            StokNakliTalepNo = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_7" select new XElement(y.Element("Value"))).First().Value.ToString(),
                                            LineNum = (from y in x.Element("Columns").Elements("Column") where y.Element("ID").Value == "Col_8" select new XElement(y.Element("Value"))).First().Value.ToString(),


                                        }).ToList();

                            string Successmessage = "";// "Başarılı " + System.Environment.NewLine;

                            string Errormessage = "";// " Başarısız " + System.Environment.NewLine;

                            int basarili = 0;

                            foreach (var item in rows)
                            {
                                SAPbobsCOM.StockTransfer oInventoryTransferRequest = (SAPbobsCOM.StockTransfer)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest);

                                oInventoryTransferRequest.GetByKey(Convert.ToInt32(item.StokNakliTalepNo));
                                oInventoryTransferRequest.Lines.SetCurrentLine(Convert.ToInt32(item.LineNum));


                                oInventoryTransferRequest.Lines.Quantity = Convert.ToDouble(item.RevizeMiktar);
                                oInventoryTransferRequest.Lines.UserFields.Fields.Item("U_RevizeMi").Value = "Y";


                                int retval = oInventoryTransferRequest.Update();
                                if (retval == 0)
                                {
                                    basarili++;
                                    // Successmessage += item.StokNakliTalepNo + " numaralı belgenin " + (Convert.ToInt32(item.LineNum) + 1) + " satır miktarı revize edilmiştir.\r Miktar : " + item.Miktar + "\r Revize Miktar :" + item.RevizeMiktar;
                                    //Handler.SAPApplication.MessageBox(StokNakliTalepNo + " numaralı belgenin " + (Convert.ToInt32(LineNum) + 1) + " satır miktarı revize edilmiştir.\r Miktar : " + Miktar + "\r Revize Miktar :" + RevizeMiktar);
                                }
                                else
                                {
                                    Errormessage += item.StokNakliTalepNo + "-" + item.UrunKodu + " numaralı Stok nakli talep numarası güncellenirken hata oluştu. " + ConstVariables.oCompanyObject.GetLastErrorDescription() + System.Environment.NewLine;
                                    //Handler.SAPApplication.MessageBox("Stok nakli talep numarası güncellenirken hata oluştu. \r" + ConstVariables.oCompanyObject.GetLastErrorDescription());
                                }
                            }
                            string message = basarili + " adet belge başarılı bir şekilde revize edilmiştir. " + System.Environment.NewLine + Errormessage;
                            Handler.SAPApplication.MessageBox(message);
                        }
                        catch (Exception ex)
                        {
                            Handler.SAPApplication.MessageBox(ex.ToString());
                        }
                        finally
                        {
                           // Listele(BaslangicTarihi, BitisTarihi);
                        }


                        #region Gridli işlemdi.
                        //for (int i = 0; i < oGrid.Rows.SelectedRows.Count; i++)
                        //{
                        //    try
                        //    {
                        //        int row = oGrid.Rows.SelectedRows.Item(i, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
                        //        string UrunKodu = oGrid.DataTable.GetValue("Ürün Kodu", row).ToString();
                        //        string UrunTanimi = oGrid.DataTable.GetValue("Ürün Tanımı", row).ToString();
                        //        string Miktar = oGrid.DataTable.GetValue("Miktar", row).ToString();
                        //        string RevizeMiktar = oGrid.DataTable.GetValue("Revize Edilecek Miktar", row).ToString();
                        //        string Fark = oGrid.DataTable.GetValue("Fark", row).ToString();
                        //        string StokNakliTalepNo = oGrid.DataTable.GetValue("Stok Nakli Talep No", row).ToString();
                        //        string LineNum = oGrid.DataTable.GetValue("LineNum", row).ToString();

                        //        SAPbobsCOM.StockTransfer oInventoryTransferRequest = (SAPbobsCOM.StockTransfer)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest);

                        //        oInventoryTransferRequest.GetByKey(Convert.ToInt32(StokNakliTalepNo));
                        //        oInventoryTransferRequest.Lines.SetCurrentLine(Convert.ToInt32(LineNum));


                        //        oInventoryTransferRequest.Lines.Quantity = Convert.ToDouble(RevizeMiktar);
                        //        oInventoryTransferRequest.Lines.UserFields.Fields.Item("U_RevizeMi").Value = "Y";


                        //        int retval = oInventoryTransferRequest.Update();
                        //        if (retval == 0)
                        //        {

                        //            Successmessage += StokNakliTalepNo + " numaralı belgenin " + (Convert.ToInt32(LineNum) + 1) + " satır miktarı revize edilmiştir.\r Miktar : " + Miktar + "\r Revize Miktar :" + RevizeMiktar;
                        //            //Handler.SAPApplication.MessageBox(StokNakliTalepNo + " numaralı belgenin " + (Convert.ToInt32(LineNum) + 1) + " satır miktarı revize edilmiştir.\r Miktar : " + Miktar + "\r Revize Miktar :" + RevizeMiktar);
                        //        }
                        //        else
                        //        {
                        //            Errormessage += StokNakliTalepNo + "Stok nakli talep numarası güncellenirken hata oluştu. \r" + ConstVariables.oCompanyObject.GetLastErrorDescription();
                        //            //Handler.SAPApplication.MessageBox("Stok nakli talep numarası güncellenirken hata oluştu. \r" + ConstVariables.oCompanyObject.GetLastErrorDescription());
                        //        }
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        Errormessage += "Stok nakli talep numarası güncellenirken hata oluştu. \r" + ex.ToString();
                        //        //Handler.SAPApplication.MessageBox("Stok nakli talep numarası güncellenirken hata oluştu. \r" + ex.ToString());
                        //    }

                        //} 
                        #endregion
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
        public void Listele(string BaslangicTarih, string BitisTarih)
        {
            #region MyRegion
            //string BaslangicTarihi = EdtBasTarih.Value.ToString();
            //string BitisTarihi = EdtBitTarih.Value.ToString(); ;
            //string Sorgu = " Select T1.ItemCode as 'Ürün Kodu',T1.Dscription as 'Ürün Tanımı',T1.Quantity as 'Miktar' ,cast(T2.U_StkNkOran as float) as 'oran',T1.OpenQty + ((T1.OpenQty * cast(T2.U_StkNkOran as float)) / 100) as 'Revize Edilecek Miktar',(T1.OpenQty + ((T1.OpenQty * cast(T2.U_StkNkOran as float)) / 100)) - T1.OpenQty as 'Fark',T0.DocEntry as 'Stok Nakli Talep No',T1.\"LineNum\" from OWTQ T0 inner join WTQ1 T1 on T0.DocEntry = T1.DocEntry left join OITM T2 on T1.ItemCode = T2.ItemCode where T0.DocDate between '" + BaslangicTarihi + "' and '" + BitisTarihi + "' and cast(ISNULL(T2.U_StkNkOran,'0') as float) > 0 and T1.OpenQty > 0   and ISNULL(T1.U_RevizeMi,'N')='N' ";
            //oGrid.DataTable.ExecuteQuery(Sorgu);

            //EditTextColumn UrunKodu = ((SAPbouiCOM.EditTextColumn)(oGrid.Columns.Item("Ürün Kodu")));

            //UrunKodu.LinkedObjectType = "4";

            //EditTextColumn StokNakliTalepNo = ((SAPbouiCOM.EditTextColumn)(oGrid.Columns.Item("Stok Nakli Talep No")));

            //StokNakliTalepNo.LinkedObjectType = "1250000001";

            //EditTextColumn oran = ((SAPbouiCOM.EditTextColumn)(oGrid.Columns.Item("oran")));
            //oran.Visible = false;
            //EditTextColumn LineNum = ((SAPbouiCOM.EditTextColumn)(oGrid.Columns.Item("LineNum")));
            //LineNum.Width = 0;

            //oGrid.AutoResizeColumns(); 
            #endregion

            try
            {
                //    string BaslangicTarihi = EdtBasTarih.Value.ToString();
                //    string BitisTarihi = EdtBitTarih.Value.ToString(); ;
                string Sorgu = " Select T1.ItemCode as 'Ürün Kodu',T1.Dscription as 'Ürün Tanımı',T1.Quantity as 'Miktar' ,cast(T2.U_StkNkOran as float) as 'Oran',T1.OpenQty + ((T1.OpenQty * cast(T2.U_StkNkOran as float)) / 100) as 'Revize Edilecek Miktar',(T1.OpenQty + ((T1.OpenQty * cast(T2.U_StkNkOran as float)) / 100)) - T1.OpenQty as 'Fark',T0.DocEntry as 'Stok Nakli Talep No',T1.\"LineNum\",T1.\"FromWhsCod\",T1.\"WhsCode\" from OWTQ T0 inner join WTQ1 T1 on T0.DocEntry = T1.DocEntry left join OITM T2 on T1.ItemCode = T2.ItemCode where T0.DocDate between '" + BaslangicTarih + "' and '" + BitisTarih + "' and cast(ISNULL(T2.U_StkNkOran,'0') as float) > 0 and T1.OpenQty > 0   and ISNULL(T1.U_RevizeMi,'N')='N' ";

                oDT.Clear();
                oDT.ExecuteQuery(Sorgu);

                //oMatrix.Columns.Item("Col_0").DataBind.Bind("DATA", "Sec");
                oMatrix.Columns.Item("Col_1").DataBind.Bind("DATA", "Ürün Kodu");
                oMatrix.Columns.Item("Col_2").DataBind.Bind("DATA", "Ürün Tanımı");
                oMatrix.Columns.Item("Col_3").DataBind.Bind("DATA", "Miktar");
                oMatrix.Columns.Item("Col_4").DataBind.Bind("DATA", "Oran");
                oMatrix.Columns.Item("Col_5").DataBind.Bind("DATA", "Revize Edilecek Miktar");
                oMatrix.Columns.Item("Col_6").DataBind.Bind("DATA", "Fark");
                oMatrix.Columns.Item("Col_7").DataBind.Bind("DATA", "Stok Nakli Talep No");
                oMatrix.Columns.Item("Col_8").DataBind.Bind("DATA", "LineNum");
                oMatrix.Columns.Item("Col_9").DataBind.Bind("DATA", "FromWhsCod");
                oMatrix.Columns.Item("Col_10").DataBind.Bind("DATA", "WhsCode");


                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Handler.SAPApplication.MessageBox("Sorgulama yapılırken hata oluştu. " + ex.ToString());
            }

        }
    }

}
