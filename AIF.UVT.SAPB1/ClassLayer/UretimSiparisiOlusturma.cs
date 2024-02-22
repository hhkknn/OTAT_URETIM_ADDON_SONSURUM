using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.UVT.SAPB1.HelperClass;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AIF.UVT.SAPB1.ClassLayer
{
    public class UretimSiparisiOlusturma : IUserForm, IMenuEvents, IRightEvents
    {
        [ItemAtt(AIFConn.UretimSiparisiOlusturmaUID)]
        public SAPbouiCOM.Form frmUretimSiparisiOlusturma;

        //[ItemAtt("Item_3")]
        //public SAPbouiCOM.Folder oFolderSonUrunKimyasal;

        [ItemAtt("Item_0")]
        public SAPbouiCOM.Matrix oMatrix;

        [ItemAtt("Item_4")]
        public SAPbouiCOM.EditText oEditSatirSayisi;

        //[ItemAtt("Item_9")]
        //public SAPbouiCOM.Matrix oMatrixMikrobiyolojik;
        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.UretimSiparisiOlusturmaXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.UretimSiparisiOlusturmaXML));
            Functions.CreateUserOrSystemFormComponent<UretimSiparisiOlusturma>(AIFConn.sprsOlstr);

            InitForms();
        }

        public void InitForms()
        {
            //oMatrix.AddRow();
            oMatrix.AutoResizeColumns();
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

        public WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
        {

            string relId = "";
            try
            {
                relId = workbookPart.Workbook.Descendants<Sheet>().First(s => sheetName.Equals(s.Name)).Id;
            }
            catch (Exception)
            {
                return null;
            }
            return (WorksheetPart)workbookPart.GetPartById(relId);
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
                    if (pVal.ItemUID == "Item_1" && !pVal.BeforeAction)
                    {
                        SAPbobsCOM.ProductionOrders oWOR = (SAPbobsCOM.ProductionOrders)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);

                        //oWOR.GetByKey(Convert.ToInt32(EdtUretimSiparisNo.Value));

                        //oWOR.PlannedQuantity = parseNumber.parservalues<double>(EdtUretimSiparisMiktari.Value);

                        try
                        {
                            for (int i = 1; i <= oMatrix.RowCount; i++)
                            {
                                oWOR = (SAPbobsCOM.ProductionOrders)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);

                                if (((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("#").Cells.Item(i).Specific).Checked == false)
                                {
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_5").Cells.Item(i).Specific).Value = "Seçim işaretli olmadığından işlem yapılmadı.";
                                    continue;
                                }

                                if (((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value == "")
                                {
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_5").Cells.Item(i).Specific).Value = "Kalem Kodu eksik olduğundan işlem yapılmadı.";
                                    continue;
                                }

                                oWOR.ItemNo = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value;
                                oWOR.PlannedQuantity = parseNumber.parservalues<double>(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(i).Specific).Value);
                                oWOR.Warehouse = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_3").Cells.Item(i).Specific).Value.PadLeft(3, '0');

                                string baslangicTarihi = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_4").Cells.Item(i).Specific).Value;

                                if (baslangicTarihi != "")
                                {
                                    DateTime dt = new DateTime(Convert.ToInt32(baslangicTarihi.Substring(0, 4)), Convert.ToInt32(baslangicTarihi.Substring(4, 2)), Convert.ToInt32(baslangicTarihi.Substring(6, 2)));

                                    oWOR.StartDate = dt;
                                }

                                //oWOR.UserFields.Fields.Item("U_PartiBelgeNo").Value = docentry;

                                ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                ConstVariables.oRecordset.DoQuery("Select ISNULL(\"U_ISTASYON\",'') from \"OITT\" where \"Code\" = '" + oWOR.ItemNo + "'");

                                string istasyon = ConstVariables.oRecordset.Fields.Item(0).Value.ToString();

                                if (istasyon != "")
                                {
                                    oWOR.UserFields.Fields.Item("U_ISTASYON").Value = istasyon;
                                }

                                var aaa = oWOR.Add();
                                var bbb = ConstVariables.oCompanyObject.GetLastErrorDescription();

                                if (aaa < 0)
                                {
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_5").Cells.Item(i).Specific).Value = bbb;
                                    continue;
                                }

                                //    oWOR.GetByKey(Convert.ToInt32(ConstVariables.oCompanyObject.GetNewObjectKey()));

                                //    SAPbobsCOM.Recordset orec = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                //    string docentry = "";
                                //    var oCompanyServiceData = ConstVariables.oCompanyObject.GetCompanyService();
                                //    var oGeneralServiceData = oCompanyServiceData.GetGeneralService("AIF_URT_PART");
                                //    SAPbobsCOM.GeneralDataCollection oChildren = null;
                                //    SAPbobsCOM.GeneralData oChild = null;
                                //    var oGeneralDataData = ((SAPbobsCOM.GeneralData)(oGeneralServiceData.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                                //    ConstVariables.oRecordset = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
                                //    ConstVariables.oRecordset.DoQuery("Select ISNULL(\"U_PartiMiktari\",0),\"U_ISTASYON\" from \"OITT\" where \"Code\" = '" + oWOR.ItemNo + "'");

                                //    double UrunAgaciPartiMiktari = Convert.ToDouble(ConstVariables.oRecordset.Fields.Item(0).Value);
                                //    double OworPlananMiktar = oWOR.PlannedQuantity;
                                //    double yazilacakMiktar = 0;
                                //    double yazilanToplamMiktar = 0;
                                //    DateTime dttarih = oWOR.StartDate;

                                //    //oGeneralDataData.SetProperty("DocNum", Convert.ToInt32(docentry));
                                //    oGeneralDataData.SetProperty("U_UretimSipNo", oWOR.DocumentNumber.ToString());
                                //    oGeneralDataData.SetProperty("U_UretimSipTar", dttarih);
                                //    oGeneralDataData.SetProperty("U_UretimSipMik", oWOR.PlannedQuantity);
                                //    oGeneralDataData.SetProperty("U_UrunKodu", oWOR.ItemNo);
                                //    oGeneralDataData.SetProperty("U_UrunTanimi", oWOR.ProductDescription);

                                //    oChildren = oGeneralDataData.Child("AIF_URT_PART1");

                                //    int sira = 1;

                                //    if (UrunAgaciPartiMiktari > 0)
                                //    {
                                //    tekrarla:
                                //        if (OworPlananMiktar < UrunAgaciPartiMiktari)
                                //        {
                                //            yazilacakMiktar = OworPlananMiktar;
                                //            OworPlananMiktar = OworPlananMiktar - yazilacakMiktar;
                                //        }
                                //        else
                                //        {
                                //            yazilacakMiktar = UrunAgaciPartiMiktari;
                                //            OworPlananMiktar = Math.Round(OworPlananMiktar, 2) - yazilacakMiktar;
                                //        }

                                //        oChild = oChildren.Add();

                                //        oChild.SetProperty("U_PartiNo", DateTime.Now.ToString("yyyyMMdd") + "-" + oWOR.DocumentNumber.ToString() + "-" + sira.ToString());
                                //        oChild.SetProperty("U_Miktar", yazilacakMiktar);
                                //        oChild.SetProperty("U_PartiKatSayi", Convert.ToDouble(oWOR.PlannedQuantity / yazilacakMiktar));

                                //        if (OworPlananMiktar > 0)
                                //        {
                                //            yazilanToplamMiktar = yazilanToplamMiktar + yazilacakMiktar;
                                //            sira++;
                                //            goto tekrarla;

                                //        }
                                //    }
                                //    else
                                //    {
                                //        oChild = oChildren.Add();

                                //        oChild.SetProperty("U_PartiNo", DateTime.Now.ToString("yyyyMMdd") + "-" + oWOR.DocumentNumber.ToString() + "-" + sira.ToString());
                                //        oChild.SetProperty("U_Miktar", OworPlananMiktar);
                                //        oChild.SetProperty("U_PartiKatSayi", Convert.ToDouble("1"));
                                //    }

                                //    SAPbobsCOM.GeneralDataParams result = oGeneralServiceData.Add(oGeneralDataData);

                                //    oGeneralDataData = ((SAPbobsCOM.GeneralData)(oGeneralServiceData.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                                //    oChildren = null;
                                //    oChild = null;

                                //    orec.DoQuery("Select MAX(\"DocEntry\") from \"@AIF_URT_PART\" where \"U_UretimSipNo\" = '" + oWOR.DocumentNumber + "'");

                                //    docentry = orec.Fields.Item(0).Value.ToString();

                                //    if (docentry != "")
                                //    {
                                //        oWOR.ProductionOrderStatus = BoProductionOrderStatusEnum.boposReleased;
                                //        oWOR.UserFields.Fields.Item("U_PartiBelgeNo").Value = docentry;
                                //        if (istasyon != "")
                                //        {
                                //            oWOR.UserFields.Fields.Item("U_ISTASYON").Value = istasyon;
                                //        }
                                //        oWOR.Update();
                                //    }

                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_5").Cells.Item(i).Specific).Value = "Başarılı";
                            }

                            Handler.SAPApplication.MessageBox("İşlem tamamnlandı.");
                            oMatrix.AutoResizeColumns();
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Item_3" && !pVal.BeforeAction)
                    {
                        //string strPath = "";

                        //OpenFileDialog openDlg = new OpenFileDialog();

                        //openDlg.InitialDirectory = @"C:\";

                        //openDlg.Filter = "Excel Dosyası (*.xls)|*.xls| Excel Dosyası (*.xlsx)|*.xlsx| Excel Dosyası (*.xlsm)|*.xlsm";

                        //openDlg.FilterIndex = 0;

                        //openDlg.RestoreDirectory = true;

                        //openDlg.Title = "Excel seçimi";

                        //DialogResult ret = STAShowDialog(openDlg);

                        //if (ret == System.Windows.Forms.DialogResult.OK)
                        //{
                        //    strPath = openDlg.FileName;
                        //}
                        //else
                        //{
                        //    Handler.SAPApplication.MessageBox("Dosya seçimi yapılmadı.");
                        //    return false;
                        //}

                        oMatrix.Clear();
                        oMatrix.Columns.Item("Col_0").Editable = false;
                        oMatrix.Columns.Item("Col_1").Editable = false;
                        oMatrix.Columns.Item("Col_3").Editable = false;
                        oMatrix.Columns.Item("Col_4").Editable = false;

                        //string filePath = showOpenFileDialog();
                        //DosyaSec fd = new DosyaSec();

                        //string FilePath = fd.GetFileName();


                        var fssservice = new AttachmentCreate();
                        string dosya = fssservice.showOpenFileDialog();

                        if (dosya == "")
                        {
                            return false;
                        }
                        else
                        {
                            var val = dosya.Split('|');

                            FilePath = val[0];

                            //oEditDosyaYolu.Value = val[0];
                            //oEditDosyaAdi.Value = val[1];
                        }
                        //var fi = new FileInfo(oEditDosyaYolu.Value);

                        SAPbouiCOM.ProgressBar oProgressBar = null;
                        try
                        {
                            frmUretimSiparisiOlusturma.Freeze(true);
                            string fileName = FilePath;

                            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            {
                                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                                {
                                    WorkbookPart workbookPart = doc.WorkbookPart;
                                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                                    SharedStringTable sst = sstpart.SharedStringTable;

                                    WorksheetPart worksheetPart = GetWorksheetPart(workbookPart, "SAPB1");


                                    if (worksheetPart == null)
                                    {
                                        Handler.SAPApplication.MessageBox("Geçersiz excel formatı girişi yapılmıştır. Lütfen kontrol ediniz.");
                                        return false;
                                    }
                                    Worksheet sheet = worksheetPart.Worksheet;

                                    var cells = sheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>();
                                    var rows = sheet.Descendants<Row>();

                                    // One way: go through each cell in the sheet
                                    //foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in cells)
                                    //{
                                    //    if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                                    //    {
                                    //        int ssid = int.Parse(cell.CellValue.Text);
                                    //        string str = sst.ChildElements[ssid].InnerText;
                                    //        Console.WriteLine("Shared string {0}: {1}", ssid, str);
                                    //    }
                                    //    else if (cell.CellValue != null)
                                    //    {
                                    //        Console.WriteLine("Cell contents: {0}", cell.CellValue.Text);
                                    //    }
                                    //}

                                    int Progress = 0;
                                    oProgressBar = Handler.SAPApplication.StatusBar.CreateProgressBar("Excel'den aktarım sağlanıyor. Kayıt Sayısı:" + rows.Count(), rows.Count(), true);
                                    oProgressBar.Text = "Excel'den aktarım sağlanıyor. Kayıt Sayısı:" + Convert.ToInt32(rows.Count() - 1).ToString();

                                    int i = 0;
                                    int column = 1;
                                    int index = 1;
                                    foreach (Row row in rows)
                                    {
                                        if (i >= 1)
                                        {
                                            oMatrix.AddRow();
                                        }
                                        column = 1;
                                        foreach (DocumentFormat.OpenXml.Spreadsheet.Cell c in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                                        {
                                            if (i > 0)
                                            {
                                                if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                                                {
                                                    int ssid = int.Parse(c.CellValue.Text);
                                                    string str = sst.ChildElements[ssid].InnerText;

                                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item(column).Cells.Item(index).Specific).Value = str;
                                                    column++;
                                                }
                                                else if (c.CellValue != null)
                                                {
                                                    if (c.CellValue.Text == "44210")
                                                    {
                                                        double d = double.Parse(c.InnerText);
                                                        DateTime conv = DateTime.FromOADate(d).Date;

                                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item(column).Cells.Item(index).Specific).Value = conv.ToString("yyyyMMdd");
                                                    }
                                                    else
                                                    {
                                                        if (oMatrix.Columns.Item(column).TitleObject.Caption == "Depo Kodu")
                                                        {
                                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item(column).Cells.Item(index).Specific).Value = c.CellValue.Text.PadLeft(3, '0');
                                                        }
                                                        else
                                                        {
                                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item(column).Cells.Item(index).Specific).Value = c.CellValue.Text;
                                                        }
                                                    }
                                                    column++;
                                                }
                                            }
                                        }
                                        i++;

                                        if (i > 1)
                                        {
                                            index++;

                                            Progress += 1;
                                            oProgressBar.Value = Progress;
                                        }
                                    }
                                }
                            }
                            oMatrix.AutoResizeColumns();
                        }
                        catch (Exception Ex)
                        {
                        }
                        finally
                        {
                            if (oProgressBar != null)
                            {
                                try
                                {
                                    oProgressBar.Stop();
                                }
                                catch (Exception)
                                {
                                }
                                finally
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar);
                                }
                            }
                            frmUretimSiparisiOlusturma.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "Item_5" && !pVal.BeforeAction)
                    {
                        int satirSayisi = oEditSatirSayisi.Value == "" ? 0 : Convert.ToInt32(oEditSatirSayisi.Value);

                        if (satirSayisi > 0)
                        {
                            oMatrix.Clear();
                            oMatrix.AddRow(satirSayisi);

                            oMatrix.Columns.Item("Col_0").Editable = true;
                            oMatrix.Columns.Item("Col_1").Editable = true;
                            oMatrix.Columns.Item("Col_3").Editable = true;
                            oMatrix.Columns.Item("Col_4").Editable = true;
                            oMatrix.AutoResizeColumns();
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

        public string FileName;
        public string FilePath;

        public string showOpenFileDialog()
        {
            System.Threading.Thread ShowFolderBrowserThread = default(System.Threading.Thread);
            try
            {
                ShowFolderBrowserThread = new System.Threading.Thread(ShowFolderBrowser);
                if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Unstarted)
                {
                    ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    ShowFolderBrowserThread.Start();
                }
                else if (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Stopped)
                {
                    ShowFolderBrowserThread.Start();
                    ShowFolderBrowserThread.Join();
                }
                while (ShowFolderBrowserThread.ThreadState == System.Threading.ThreadState.Running)
                {
                    System.Windows.Forms.Application.DoEvents();
                }
                if (!string.IsNullOrEmpty(FileName))
                {
                    return FileName;
                }
            }
            catch (Exception ex)
            {
                //SBO_Application.MessageBox("FileFile" & ex.Message)
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }

            return "";
        }

        public void ShowFolderBrowser()
        {
            System.Diagnostics.Process[] MyProcs = null;
            FileName = "";
            System.Windows.Forms.OpenFileDialog OpenFile = new System.Windows.Forms.OpenFileDialog();

            try
            {
                OpenFile.Multiselect = false;
                OpenFile.Filter = "Excel Dosyası (*.xls)|*.xls|Excel Dosyası (*.xlsx)|*.xlsx|Excel Dosyası (*.xlsm)|*.xlsm";

                int filterindex = 0;
                try
                {
                    filterindex = 0;
                }
                catch (Exception ex)
                {
                }

                OpenFile.FilterIndex = filterindex;

                OpenFile.RestoreDirectory = true;
                MyProcs = System.Diagnostics.Process.GetProcessesByName("SAP Business One").ToArray();
                //MyProcs = System.Diagnostics.Process.GetProcessesByName("AIF.UVT.SAPB1.exe").ToArray();

                if (MyProcs.Length == 1)
                {
                    for (int i = 0; i <= MyProcs.Length - 1; i++)
                    {
                        WindowWrapper MyWindow = new WindowWrapper(MyProcs[i].MainWindowHandle);
                        System.Windows.Forms.DialogResult ret = OpenFile.ShowDialog(MyWindow);

                        if (ret == System.Windows.Forms.DialogResult.OK)
                        {
                            FilePath = System.IO.Path.Combine("{0}", OpenFile.FileName);

                            FileName = OpenFile.SafeFileName;

                            //File.Move(FilePath, string.Format("D:\\{0}", FileName));

                            OpenFile.Dispose();
                        }
                        else
                        {
                            System.Windows.Forms.Application.ExitThread();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //SBO_Application.StatusBar.SetText(ex.Message)
                System.Windows.Forms.MessageBox.Show(ex.ToString());
                FileName = "";
            }
            finally
            {
                OpenFile.Dispose();
            }
        }

        public class DialogState

        {
            public DialogResult result;

            public FileDialog dialog;

            public void ThreadProcShowDialog()

            {
                result = dialog.ShowDialog();
            }
        }

        private static DialogResult STAShowDialog(FileDialog dialog)

        {
            DialogState state = new DialogState();

            state.dialog = dialog;

            System.Threading.Thread t = new System.Threading.Thread(state.ThreadProcShowDialog);

            t.SetApartmentState(System.Threading.ApartmentState.STA);

            t.Start();

            t.Join();

            return state.result;
        }

        public void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }
    }
}