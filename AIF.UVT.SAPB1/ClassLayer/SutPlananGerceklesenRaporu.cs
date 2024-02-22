using AIF.ObjectsDLL;
using AIF.ObjectsDLL.Abstarct;
using AIF.ObjectsDLL.Events;
using AIF.ObjectsDLL.Lib;
using AIF.ObjectsDLL.Utils;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Handler = AIF.ObjectsDLL.Events.Handler;


namespace AIF.UVT.SAPB1.ClassLayer
{
    public class SutPlananGerceklesenRaporu
    {
        [ItemAtt(AIFConn.SutPlananGerceklesenKarsilastirmaRaporuUID)]
        public SAPbouiCOM.Form frmSutPlanlananGerceklesenRaporu;

        //[ItemAtt("Item_5")]
        //public SAPbouiCOM.EditText EdtDocEntry;
        //[ItemAtt("1")]
        //public SAPbouiCOM.Button btnAddOrUpdate;
        [ItemAtt("Item_7")]
        public SAPbouiCOM.Grid oGrid;
        [ItemAtt("Item_1")]
        public SAPbouiCOM.ComboBox oComboYil;
        [ItemAtt("Item_6")]
        public SAPbouiCOM.ComboBox oComboHafta;
        public void LoadForms()
        {
            ConstVariables.oFnc.LoadSAPXML(AIFConn.SutPlananGerceklesenKarsilastirmaRaporuXML, Assembly.GetExecutingAssembly().GetManifestResourceStream(AIFConn.SutPlananGerceklesenKarsilastirmaRaporuXML));
            Functions.CreateUserOrSystemFormComponent<SutPlananGerceklesenRaporu>(AIFConn.SutPlnGrck);

            InitForms();
        }
        public void InitForms()
        {
            try
            {
                frmSutPlanlananGerceklesenRaporu.EnableMenu("1283", false);
                frmSutPlanlananGerceklesenRaporu.EnableMenu("1284", false);
                frmSutPlanlananGerceklesenRaporu.EnableMenu("1286", false);

                int yil = 2022;
                for (int i = 1; i <= 10; i++)
                {
                    oComboYil.ValidValues.Add(yil.ToString(), yil.ToString());
                    yil++;
                }
                oComboYil.Select(DateTime.Now.Year.ToString(), BoSearchKey.psk_ByValue);
                string d = "";
                for (int i = 1; i <= 52; i++)
                {
                    d = "(" + FirstDateOfWeekISO8601(Convert.ToInt32(oComboYil.Value.Trim()), i).ToString("dd.MM.yyyy") + " - " + LastDateOfWeekISO8601(Convert.ToInt32(oComboYil.Value.Trim()), i).ToString("dd.MM.yyyy") + ")";
                    oComboHafta.ValidValues.Add(i.ToString() + ".Hafta", d);
                    yil++;
                }
            }
            catch (Exception ex)
            {
            }
        }
        public static DateTime FirstDateOfWeekISO8601(int year, int weekOfYear)
        {
            DateTime jan1 = new DateTime(year, 1, 1);
            int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;

            // Use first Thursday in January to get first week of the year as
            // it will never be in Week 52/53
            DateTime firstThursday = jan1.AddDays(daysOffset);
            var cal = CultureInfo.CurrentCulture.Calendar;
            int firstWeek = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            var weekNum = weekOfYear;
            // As we're adding days to a date in Week 1,
            // we need to subtract 1 in order to get the right date for week #1
            if (firstWeek == 1)
            {
                weekNum -= 1;
            }

            // Using the first Thursday as starting week ensures that we are starting in the right year
            // then we add number of weeks multiplied with days
            var result = firstThursday.AddDays(weekNum * 7);

            // Subtract 3 days from Thursday to get Monday, which is the first weekday in ISO8601
            return result.AddDays(-3);
        }
        public static DateTime LastDateOfWeekISO8601(int year, int weekOfYear)
        {
            DateTime jan1 = new DateTime(year, 1, 1);
            int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;

            // Use first Thursday in January to get first week of the year as
            // it will never be in Week 52/53
            DateTime firstThursday = jan1.AddDays(daysOffset);
            var cal = CultureInfo.CurrentCulture.Calendar;
            int firstWeek = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            var weekNum = weekOfYear;
            // As we're adding days to a date in Week 1,
            // we need to subtract 1 in order to get the right date for week #1
            if (firstWeek == 1)
            {
                weekNum -= 1;
            }

            // Using the first Thursday as starting week ensures that we are starting in the right year
            // then we add number of weeks multiplied with days
            var result = firstThursday.AddDays(weekNum * 7);

            // Subtract 3 days from Thursday to get Monday, which is the first weekday in ISO8601
            return result.AddDays(+3);
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
                    if (pVal.ItemUID == "Item_1" && !pVal.BeforeAction)
                    {
                        // Cem
                        //for (int i = oComboHafta.ValidValues.Count - 1; i >= 0; i--)
                        //{
                        //    oComboHafta.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        //}


                        //string Yil = ((SAPbouiCOM.ComboBox)frmSutPlanlananGerceklesenRaporu.Items.Item("Item_1").Specific).Selected.Value.ToString();
                        ////oComboYil.Select(Yil, BoSearchKey.psk_ByValue);
                        //string d = "";
                        //for (int i = 1; i <= 52; i++)
                        //{
                        //    d = "(" + FirstDateOfWeekISO8601(Convert.ToInt32(Yil), i).ToString("dd.MM.yyyy") + " - " + LastDateOfWeekISO8601(Convert.ToInt32(Yil), i).ToString("dd.MM.yyyy") + ")";
                        //    oComboHafta.ValidValues.Add(i.ToString() + ".Hafta", d);

                        //}
                    }
                    break;
                case BoEventTypes.et_CLICK:
                    if (pVal.ItemUID == "Item_10" && !pVal.BeforeAction)
                    {


                        try
                        {
                            var haftasayisi = oComboHafta.Value.ToString().Replace(".Hafta", "");

                            var ilkTarih = FirstDateOfWeekISO8601(Convert.ToInt32(oComboYil.Value.Trim()), Convert.ToInt32(haftasayisi));
                            var sonTarih = LastDateOfWeekISO8601(Convert.ToInt32(oComboYil.Value.Trim()), Convert.ToInt32(haftasayisi));


                            raporCalistir(ilkTarih.ToString("yyyyMMdd"), sonTarih.ToString("yyyyMMdd"));
                        }
                        catch (Exception)
                        {

                        }

                    }
                    else if (pVal.ItemUID == "Item_9" && !pVal.BeforeAction)
                    {
                        try
                        {
                            var haftasayisi = oComboHafta.Value.ToString().Replace(".Hafta", "");

                            var ilkTarih = FirstDateOfWeekISO8601(Convert.ToInt32(oComboYil.Value.Trim()), Convert.ToInt32(haftasayisi));
                            var sonTarih = LastDateOfWeekISO8601(Convert.ToInt32(oComboYil.Value.Trim()), Convert.ToInt32(haftasayisi));


                            htmldenGoster(ilkTarih.ToString("yyyyMMdd"), sonTarih.ToString("yyyyMMdd"));
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
        List<string> silinecekler = new List<string>();
        public void MenuEvent(ref MenuEvent pVal, ref bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        public void RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }


        private void raporCalistir(string tarih1, string tarih2)
        {
            string sql = "exec PRC_SutPlanlananGerceklesenKarsilastirma '" + tarih1 + "','" + tarih2 + "'";
            //sql += "select Tbl4.CardName AS TEDARIKCI ,Cast(Cast(sum(tbl2.P1) as decimal(15,3)) as varchar(20)) as Col2,Cast(Cast(sum(Tbl3.G1) as decimal(15,3)) as varchar(20)) as Col3,Cast(Cast(sum(Tbl1.O1) as decimal(15,2)) as varchar(20)) as Col4 ,Cast(Cast(sum(tbl2.P2) as decimal(15,3)) as varchar(20)) as Col5,Cast(Cast(sum(Tbl3.G2) as decimal(15,3)) as varchar(20)) as Col6,Cast(Cast(sum(Tbl1.O2) as decimal(15,2)) as varchar(20)) as Col7 ,Cast(Cast(sum(tbl2.P3) as decimal(15,3)) as varchar(20)) as Col8,Cast(Cast(sum(Tbl3.G3) as decimal(15,3)) as varchar(20)) as Col9,Cast(Cast(sum(Tbl1.O3) as decimal(15,2)) as varchar(20)) as Col10 ,Cast(Cast(sum(tbl2.P4) as decimal(15,3)) as varchar(20)) as Col11,Cast(Cast(sum(Tbl3.G4) as decimal(15,3)) as varchar(20)) as Col12,Cast(Cast(sum(Tbl1.O4) as decimal(15,2)) as varchar(20)) as Col13 ,Cast(Cast(sum(tbl2.P5) as decimal(15,3)) as varchar(20)) as Col14,Cast(Cast(sum(Tbl3.G5) as decimal(15,3)) as varchar(20)) as Col15,Cast(Cast(sum(Tbl1.O5) as decimal(15,2)) as varchar(20)) as Col16 ,Cast(Cast(sum(tbl2.P6) as decimal(15,3)) as varchar(20)) as Col17,Cast(Cast(sum(Tbl3.G6) as decimal(15,3)) as varchar(20)) as Col18,Cast(Cast(sum(Tbl1.O6) as decimal(15,2)) as varchar(20)) as Col19 ,Cast(Cast(sum(tbl2.P7) as decimal(15,3)) as varchar(20)) as Col20,Cast(Cast(sum(Tbl3.G7) as decimal(15,3)) as varchar(20)) as Col21,Cast(Cast(sum(Tbl1.O7) as decimal(15,2)) as varchar(20)) as Col22 ,Cast(Cast(sum(tbl2.P1) as decimal(15,2))+ Cast(sum(tbl2.P2) as decimal(15,3))+ Cast(sum(tbl2.P3) as decimal(15,3))+ Cast(sum(tbl2.P4) as decimal(15,3))+ Cast(sum(tbl2.P5) as decimal(15,3))+ Cast(sum(tbl2.P6) as decimal(15,3))+ Cast(sum(tbl2.P7) as decimal(15,3)) as varchar(20)) as \"Col23\" ,Cast(Cast(sum(tbl3.G1) as decimal(15,3))+ Cast(sum(tbl3.G2) as decimal(15,3))+ Cast(sum(tbl3.G3) as decimal(15,3))+ Cast(sum(tbl3.G4) as decimal(15,3))+ Cast(sum(tbl3.G5) as decimal(15,3))+ Cast(sum(tbl3.G6) as decimal(15,3))+ Cast(sum(tbl3.G7) as decimal(15,3)) as varchar(20)) as \"Col24\" ,Cast(Cast((Cast(sum(tbl2.P1) as decimal(15,3))+ Cast(sum(tbl2.P2) as decimal(15,3))+ Cast(sum(tbl2.P3) as decimal(15,3))+ Cast(sum(tbl2.P4) as decimal(15,3))+ Cast(sum(tbl2.P5) as decimal(15,3))+ Cast(sum(tbl2.P6) as decimal(15,3))+ Cast(sum(tbl2.P7) as decimal(15,3))) / ISNULL((Cast(sum(tbl3.G1) as decimal(15,3))+ Cast(sum(tbl3.G2) as decimal(15,3))+ Cast(sum(tbl3.G3) as decimal(15,3))+ Cast(sum(tbl3.G4) as decimal(15,3))+ Cast(sum(tbl3.G5) as decimal(15,3))+ Cast(sum(tbl3.G6) as decimal(15,3))+ Cast(sum(tbl3.G7) as decimal(15,3))),1) as decimal(15,2)) as varchar(20)) as \"Col25\" from (select U_SaticiKodu,ISNULL([1],0) as \"O1\",ISNULL([2],0) as \"O2\",ISNULL([3],0) as \"O3\",ISNULL([4],0) as \"O4\",ISNULL([5],0) as \"O5\",ISNULL([6],0) as \"O6\",ISNULL([7],0) as \"O7\" from ( SELECT  A.U_SaticiKodu ,ISNULL(A.Miktar,0) /ISNULL(B.U_Miktar,1) AS \"Oran\" ,day(A.U_Tarih)-Day('" + tarih1 + "')+1 as \"Gün\" ,(A.U_Tarih) FROM ( SELECT  T0.U_Tarih,t1.U_SaticiKodu ,SUM(T1.U_Miktar) AS \"Miktar\" FROM \"@AIF_SUTPLANLAMA\" t0 INNER JOIN \"@AIF_SUTPLANLAMA1\" T1 ON t1.DocEntry = T0.DocEntry GROUP BY  T0.U_Tarih ,t1.U_SaticiKodu ) A INNER JOIN (SELECT  S1.U_BelgeTarihi ,S0.U_Tedarikci ,S0.U_Miktar FROM \"@AIF_SUTKABUL2\" S0 INNER JOIN \"@AIF_SUTKABUL\" S1 ON s1.DocEntry = S0.DocEntry ) B ON B.U_BelgeTarihi = A.U_Tarih AND A.U_SaticiKodu = B.U_Tedarikci where A.U_Tarih between '" + tarih1 + "' and '" + tarih2 + "' ) as ORANQ PIVOT( SUM(Oran) for \"Gün\" in ([1],[2],[3],[4],[5],[6],[7]) ) as PVT1 ) AS TBL1 inner join (select U_SaticiKodu,ISNULL([1],0) as \"P1\",ISNULL([2],0) as \"P2\",ISNULL([3],0) as \"P3\",ISNULL([4],0) as \"P4\",ISNULL([5],0) as \"P5\",ISNULL([6],0) as \"P6\",ISNULL([7],0) as \"P7\" from (SELECT  A.U_SaticiKodu ,ISNULL(SUM(A.Miktar),0)   AS \"Planlanan Miktar\" ,day(A.U_Tarih)-Day('" + tarih1 + "')+1 as \"Gün\" ,(A.U_Tarih) FROM (SELECT T0.U_Tarih, T1.U_SaticiKodu,T1.U_Miktar as \"Miktar\" FROM \"@AIF_SUTPLANLAMA\" T0 LEFT JOIN \"@AIF_SUTPLANLAMA1\" T1 ON T0.DocEntry = T1.DocEntry union all SELECT T0.U_Tarih,T2.U_SaticiKodu,T2.U_Miktar as \"Miktar\" FROM \"@AIF_SUTPLANLAMA\" T0 LEFT JOIN \"@AIF_SUTPLANLAMA2\" T2 ON T2.DocEntry = T0.DocEntry ) A where A.U_Tarih between '" + tarih1 + "' and '" + tarih2 + "' Group by A.U_SaticiKodu,(A.U_Tarih) ) as PlanlanQ PIVOT(SUM(\"Planlanan Miktar\") for \"Gün\" in ([1],[2],[3],[4],[5],[6],[7]) ) as PVT2 ) tbl2 on tbl2.U_SaticiKodu=Tbl1.U_SaticiKodu inner join (select U_Tedarikci,ISNULL([1],0) as \"G1\",ISNULL([2],0) as \"G2\",ISNULL([3],0) as \"G3\",ISNULL([4],0) as \"G4\",ISNULL([5],0) as \"G5\",ISNULL([6],0) as \"G6\",ISNULL([7],0) as \"G7\" from (SELECT  B.U_Tedarikci ,ISNULL(B.U_Miktar,0) AS \"Gelen Miktar\" ,day(B.U_BelgeTarihi)-Day('" + tarih1 + "')+1 as \"Gün\" ,(B.U_BelgeTarihi) FROM (SELECT  S1.U_BelgeTarihi ,S0.U_Tedarikci ,S0.U_Miktar FROM \"@AIF_SUTKABUL2\" S0 INNER JOIN \"@AIF_SUTKABUL\" S1 ON s1.DocEntry = S0.DocEntry ) B where B.U_BelgeTarihi between '" + tarih1 + "' and '" + tarih2 + "' ) as GercQ PIVOT( SUM(\"Gelen Miktar\") for \"Gün\" in ([1],[2],[3],[4],[5],[6],[7]) ) as PVT3 )tbl3 on tbl3.U_Tedarikci=tbl2.U_SaticiKodu inner join OCRD tbl4 on tbl4.CardCode=tbl1.U_SaticiKodu group by tbl1.U_SaticiKodu, Tbl4.CardName ";

            oGrid.DataTable.ExecuteQuery(sql);


            int i = 1;
            int x = 1;
            DateTime dt = new DateTime(Convert.ToInt32(tarih1.Substring(0, 4)), Convert.ToInt32(tarih1.Substring(4, 2)), Convert.ToInt32(tarih1.Substring(6, 2)));
            for (int c = 1; c < oGrid.Columns.Count; c++)
            {
                if (i == 1)
                {
                    if (x < 22)
                    {
                        oGrid.Columns.Item(c).TitleObject.Caption = dt.Day + "." + dt.Month.ToString().PadLeft(2, '0') + " PLAN";

                        oGrid.Columns.Item(c).RightJustified = true;
                    }
                    else
                    {
                        oGrid.Columns.Item(c).TitleObject.Caption = "PLAN TOPLAM";

                        oGrid.Columns.Item(c).RightJustified = true;
                    }
                    i++;
                    x++;
                }
                else if (i == 2)
                {
                    if (x < 22)
                    {
                        oGrid.Columns.Item(c).TitleObject.Caption = dt.Day + "." + dt.Month.ToString().PadLeft(2, '0') + " GER";

                        oGrid.Columns.Item(c).RightJustified = true;
                    }
                    else
                    {
                        oGrid.Columns.Item(c).TitleObject.Caption = "GER TOPLAM";

                        oGrid.Columns.Item(c).RightJustified = true;
                    }
                    i++;
                    x++;
                }
                else if (i == 3)
                {
                    if (x < 22)
                    {
                        oGrid.Columns.Item(c).TitleObject.Caption = dt.Day + "." + dt.Month.ToString().PadLeft(2, '0') + " G/P%";

                        oGrid.Columns.Item(c).RightJustified = true;
                    }
                    else
                    {
                        oGrid.Columns.Item(c).TitleObject.Caption = "G/P% TOPLAM";

                        oGrid.Columns.Item(c).RightJustified = true;
                    }
                    dt = dt.AddDays(1);
                    i = 1;
                    x++;
                }
            }



            oGrid.AutoResizeColumns();
        }

        private void htmldenGoster(string tarih1, string tarih2)
        {
            string sql = "exec PRC_SutPlanlananGerceklesenKarsilastirma '" + tarih1 + "','" + tarih2 + "'";

            //sql += "select Tbl4.CardName AS TEDARIKCI ,Cast(Cast(sum(tbl2.P1) as decimal(15,3)) as varchar(20)) as Col2,Cast(Cast(sum(Tbl3.G1) as decimal(15,3)) as varchar(20)) as Col3,Cast(Cast(sum(Tbl1.O1) as decimal(15,2)) as varchar(20)) as Col4 ,Cast(Cast(sum(tbl2.P2) as decimal(15,3)) as varchar(20)) as Col5,Cast(Cast(sum(Tbl3.G2) as decimal(15,3)) as varchar(20)) as Col6,Cast(Cast(sum(Tbl1.O2) as decimal(15,2)) as varchar(20)) as Col7 ,Cast(Cast(sum(tbl2.P3) as decimal(15,3)) as varchar(20)) as Col8,Cast(Cast(sum(Tbl3.G3) as decimal(15,3)) as varchar(20)) as Col9,Cast(Cast(sum(Tbl1.O3) as decimal(15,2)) as varchar(20)) as Col10 ,Cast(Cast(sum(tbl2.P4) as decimal(15,3)) as varchar(20)) as Col11,Cast(Cast(sum(Tbl3.G4) as decimal(15,3)) as varchar(20)) as Col12,Cast(Cast(sum(Tbl1.O4) as decimal(15,2)) as varchar(20)) as Col13 ,Cast(Cast(sum(tbl2.P5) as decimal(15,3)) as varchar(20)) as Col14,Cast(Cast(sum(Tbl3.G5) as decimal(15,3)) as varchar(20)) as Col15,Cast(Cast(sum(Tbl1.O5) as decimal(15,2)) as varchar(20)) as Col16 ,Cast(Cast(sum(tbl2.P6) as decimal(15,3)) as varchar(20)) as Col17,Cast(Cast(sum(Tbl3.G6) as decimal(15,3)) as varchar(20)) as Col18,Cast(Cast(sum(Tbl1.O6) as decimal(15,2)) as varchar(20)) as Col19 ,Cast(Cast(sum(tbl2.P7) as decimal(15,3)) as varchar(20)) as Col20,Cast(Cast(sum(Tbl3.G7) as decimal(15,3)) as varchar(20)) as Col21,Cast(Cast(sum(Tbl1.O7) as decimal(15,2)) as varchar(20)) as Col22 ,Cast(Cast(sum(tbl2.P1) as decimal(15,2))+ Cast(sum(tbl2.P2) as decimal(15,3))+ Cast(sum(tbl2.P3) as decimal(15,3))+ Cast(sum(tbl2.P4) as decimal(15,3))+ Cast(sum(tbl2.P5) as decimal(15,3))+ Cast(sum(tbl2.P6) as decimal(15,3))+ Cast(sum(tbl2.P7) as decimal(15,3)) as varchar(20)) as \"Col23\" ,Cast(Cast(sum(tbl3.G1) as decimal(15,3))+ Cast(sum(tbl3.G2) as decimal(15,3))+ Cast(sum(tbl3.G3) as decimal(15,3))+ Cast(sum(tbl3.G4) as decimal(15,3))+ Cast(sum(tbl3.G5) as decimal(15,3))+ Cast(sum(tbl3.G6) as decimal(15,3))+ Cast(sum(tbl3.G7) as decimal(15,3)) as varchar(20)) as \"Col24\" ,Cast(Cast((Cast(sum(tbl2.P1) as decimal(15,3))+ Cast(sum(tbl2.P2) as decimal(15,3))+ Cast(sum(tbl2.P3) as decimal(15,3))+ Cast(sum(tbl2.P4) as decimal(15,3))+ Cast(sum(tbl2.P5) as decimal(15,3))+ Cast(sum(tbl2.P6) as decimal(15,3))+ Cast(sum(tbl2.P7) as decimal(15,3))) / ISNULL((Cast(sum(tbl3.G1) as decimal(15,3))+ Cast(sum(tbl3.G2) as decimal(15,3))+ Cast(sum(tbl3.G3) as decimal(15,3))+ Cast(sum(tbl3.G4) as decimal(15,3))+ Cast(sum(tbl3.G5) as decimal(15,3))+ Cast(sum(tbl3.G6) as decimal(15,3))+ Cast(sum(tbl3.G7) as decimal(15,3))),1) as decimal(15,2)) as varchar(20)) as \"Col25\" from (select U_SaticiKodu,ISNULL([1],0) as \"O1\",ISNULL([2],0) as \"O2\",ISNULL([3],0) as \"O3\",ISNULL([4],0) as \"O4\",ISNULL([5],0) as \"O5\",ISNULL([6],0) as \"O6\",ISNULL([7],0) as \"O7\" from ( SELECT  A.U_SaticiKodu ,ISNULL(A.Miktar,0) /ISNULL(B.U_Miktar,1) AS \"Oran\" ,day(A.U_Tarih)-Day('" + tarih1 + "')+1 as \"Gün\" ,(A.U_Tarih) FROM ( SELECT  T0.U_Tarih,t1.U_SaticiKodu ,SUM(T1.U_Miktar) AS \"Miktar\" FROM \"@AIF_SUTPLANLAMA\" t0 INNER JOIN \"@AIF_SUTPLANLAMA1\" T1 ON t1.DocEntry = T0.DocEntry GROUP BY  T0.U_Tarih ,t1.U_SaticiKodu ) A INNER JOIN (SELECT  S1.U_BelgeTarihi ,S0.U_Tedarikci ,S0.U_Miktar FROM \"@AIF_SUTKABUL2\" S0 INNER JOIN \"@AIF_SUTKABUL\" S1 ON s1.DocEntry = S0.DocEntry ) B ON B.U_BelgeTarihi = A.U_Tarih AND A.U_SaticiKodu = B.U_Tedarikci where A.U_Tarih between '" + tarih1 + "' and '" + tarih2 + "' ) as ORANQ PIVOT( SUM(Oran) for \"Gün\" in ([1],[2],[3],[4],[5],[6],[7]) ) as PVT1 ) AS TBL1 inner join (select U_SaticiKodu,ISNULL([1],0) as \"P1\",ISNULL([2],0) as \"P2\",ISNULL([3],0) as \"P3\",ISNULL([4],0) as \"P4\",ISNULL([5],0) as \"P5\",ISNULL([6],0) as \"P6\",ISNULL([7],0) as \"P7\" from (SELECT  A.U_SaticiKodu ,ISNULL(SUM(A.Miktar),0)   AS \"Planlanan Miktar\" ,day(A.U_Tarih)-Day('" + tarih1 + "')+1 as \"Gün\" ,(A.U_Tarih) FROM (SELECT T0.U_Tarih, T1.U_SaticiKodu,T1.U_Miktar as \"Miktar\" FROM \"@AIF_SUTPLANLAMA\" T0 LEFT JOIN \"@AIF_SUTPLANLAMA1\" T1 ON T0.DocEntry = T1.DocEntry union all SELECT T0.U_Tarih,T2.U_SaticiKodu,T2.U_Miktar as \"Miktar\" FROM \"@AIF_SUTPLANLAMA\" T0 LEFT JOIN \"@AIF_SUTPLANLAMA2\" T2 ON T2.DocEntry = T0.DocEntry ) A where A.U_Tarih between '" + tarih1 + "' and '" + tarih2 + "' Group by A.U_SaticiKodu,(A.U_Tarih) ) as PlanlanQ PIVOT(SUM(\"Planlanan Miktar\") for \"Gün\" in ([1],[2],[3],[4],[5],[6],[7]) ) as PVT2 ) tbl2 on tbl2.U_SaticiKodu=Tbl1.U_SaticiKodu inner join (select U_Tedarikci,ISNULL([1],0) as \"G1\",ISNULL([2],0) as \"G2\",ISNULL([3],0) as \"G3\",ISNULL([4],0) as \"G4\",ISNULL([5],0) as \"G5\",ISNULL([6],0) as \"G6\",ISNULL([7],0) as \"G7\" from (SELECT  B.U_Tedarikci ,ISNULL(B.U_Miktar,0) AS \"Gelen Miktar\" ,day(B.U_BelgeTarihi)-Day('" + tarih1 + "')+1 as \"Gün\" ,(B.U_BelgeTarihi) FROM (SELECT  S1.U_BelgeTarihi ,S0.U_Tedarikci ,S0.U_Miktar FROM \"@AIF_SUTKABUL2\" S0 INNER JOIN \"@AIF_SUTKABUL\" S1 ON s1.DocEntry = S0.DocEntry ) B where B.U_BelgeTarihi between '" + tarih1 + "' and '" + tarih2 + "' ) as GercQ PIVOT( SUM(\"Gelen Miktar\") for \"Gün\" in ([1],[2],[3],[4],[5],[6],[7]) ) as PVT3 )tbl3 on tbl3.U_Tedarikci=tbl2.U_SaticiKodu inner join OCRD tbl4 on tbl4.CardCode=tbl1.U_SaticiKodu group by tbl1.U_SaticiKodu, Tbl4.CardName ";


            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oRS.DoQuery(sql);


            string text = System.IO.File.ReadAllText(System.Windows.Forms.Application.StartupPath + "\\PlananGerceklesenSutRaporu.html");


            string htmldegisiklik = "";
            string yesilKirmizi = "";
            while (!oRS.EoF)
            {
                htmldegisiklik += "<tr style=\"font-size:13px\" >";
                htmldegisiklik += "<td scope=\"row\">" + oRS.Fields.Item("TEDARIKCI").Value.ToString() + "</td>";
                htmldegisiklik += "<td style=\"background-color: Moccasin\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col2").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: PeachPuff\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col3").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";


                yesilKirmizi = Convert.ToDouble(oRS.Fields.Item("Col4").Value, System.Globalization.CultureInfo.InvariantCulture) > 100 ? "PaleGreen" : "OrangeRed";


                htmldegisiklik += "<td style=\"background-color: " + yesilKirmizi + "\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col4").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: Moccasin\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col5").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: PeachPuff\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col6").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";


                yesilKirmizi = Convert.ToDouble(oRS.Fields.Item("Col7").Value, System.Globalization.CultureInfo.InvariantCulture) > 100 ? "PaleGreen" : "OrangeRed";
                htmldegisiklik += "<td style=\"background-color: " + yesilKirmizi + "\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col7").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: Moccasin\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col8").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: PeachPuff\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col9").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";


                yesilKirmizi = Convert.ToDouble(oRS.Fields.Item("Col7").Value, System.Globalization.CultureInfo.InvariantCulture) > 100 ? "PaleGreen" : "OrangeRed";
                htmldegisiklik += "<td style=\"background-color: " + yesilKirmizi + "\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col10").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: Moccasin\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col11").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: PeachPuff\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col12").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";

                yesilKirmizi = Convert.ToDouble(oRS.Fields.Item("Col13").Value, System.Globalization.CultureInfo.InvariantCulture) > 100 ? "PaleGreen" : "OrangeRed";
                htmldegisiklik += "<td style=\"background-color: " + yesilKirmizi + "\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col13").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: Moccasin\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col14").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: PeachPuff\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col15").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";


                yesilKirmizi = Convert.ToDouble(oRS.Fields.Item("Col16").Value, System.Globalization.CultureInfo.InvariantCulture) > 100 ? "PaleGreen" : "OrangeRed";
                htmldegisiklik += "<td style=\"background-color: " + yesilKirmizi + "\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col16").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: Moccasin\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col17").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: PeachPuff\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col18").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";



                yesilKirmizi = Convert.ToDouble(oRS.Fields.Item("Col19").Value, System.Globalization.CultureInfo.InvariantCulture) > 100 ? "PaleGreen" : "OrangeRed";
                htmldegisiklik += "<td style=\"background-color: " + yesilKirmizi + "\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col19").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: Moccasin\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col20").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: PeachPuff\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col21").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";



                yesilKirmizi = Convert.ToDouble(oRS.Fields.Item("Col22").Value, System.Globalization.CultureInfo.InvariantCulture) > 100 ? "PaleGreen" : "OrangeRed";
                htmldegisiklik += "<td  style=\"background-color: " + yesilKirmizi + "\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col22").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: Moccasin\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col23").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td  style=\"background-color: PeachPuff\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col24").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "<td style=\"background-color: " + yesilKirmizi + "\" align=\"center\">" + Convert.ToDouble(oRS.Fields.Item("Col25").Value, System.Globalization.CultureInfo.InvariantCulture).ToString("N2").Replace(",00", "") + "</td>";
                htmldegisiklik += "</tr>";


                oRS.MoveNext();
            }


            DateTime dt = new DateTime(Convert.ToInt32(tarih1.Substring(0, 4)), Convert.ToInt32(tarih1.Substring(4, 2)), Convert.ToInt32(tarih1.Substring(6, 2)));

            text = text.Replace("{Tarih1}", dt.ToString("dd.MM.yyyy"));
            dt = dt.AddDays(1);
            text = text.Replace("{Tarih2}", dt.ToString("dd.MM.yyyy"));
            dt = dt.AddDays(1);
            text = text.Replace("{Tarih3}", dt.ToString("dd.MM.yyyy"));
            dt = dt.AddDays(1);
            text = text.Replace("{Tarih4}", dt.ToString("dd.MM.yyyy"));
            dt = dt.AddDays(1);
            text = text.Replace("{Tarih5}", dt.ToString("dd.MM.yyyy"));
            dt = dt.AddDays(1);
            text = text.Replace("{Tarih6}", dt.ToString("dd.MM.yyyy"));
            dt = dt.AddDays(1);
            text = text.Replace("{Tarih7}", dt.ToString("dd.MM.yyyy"));
            text = text.Replace("{Satirlar}", htmldegisiklik);

            string uuid = System.Guid.NewGuid().ToString();

            File.WriteAllText(System.Windows.Forms.Application.StartupPath + "\\" + uuid + ".html", text);


            Process p = new Process
            {
                StartInfo = { FileName = System.Windows.Forms.Application.StartupPath + "\\" + uuid + ".html" }
            };
            p.Start();
        }
    }
}
