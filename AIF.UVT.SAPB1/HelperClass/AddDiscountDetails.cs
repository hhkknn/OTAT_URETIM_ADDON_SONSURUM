using AIF.ObjectsDLL.Events;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace AIF.UVT.SAPB1.HelperClass
{
    public class AddDiscountDetails
    {
        public int addDiscountDetails(int templateCode, double mainprice)
        {
            string docentry = "";
            try
            {
                SAPbobsCOM.Recordset orec = (SAPbobsCOM.Recordset)ConstVariables.oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                orec.DoQuery("Select * from \"@AIF_TMP_SLS_DISC1\" where \"DocEntry\" = '" + templateCode + "'");

                XNamespace ns = "http://www.sap.com/SBO/SDK/DI";
                XDocument xDoc = XDocument.Parse(orec.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData));

                var rows = (from t in xDoc.Descendants(ns + "Row")
                            select new
                            {
                                DiscType = (from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_DiscType" select new XElement(k.Element(ns + "Value"))).First().Value,
                                Discount = (from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_DiscRate" select new XElement(k.Element(ns + "Value"))).First().Value
                            }).ToList();


                orec.DoQuery("Select * from \"@AIF_TMP_SLS_DISC\" where \"DocEntry\" = '" + templateCode + "'");

                ns = "http://www.sap.com/SBO/SDK/DI";
                xDoc = XDocument.Parse(orec.GetFixedXML(SAPbobsCOM.RecordsetXMLModeEnum.rxmData));

                var header = (from t in xDoc.Descendants(ns + "Row")
                              select new
                              {
                                  PriceBefDisc = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_PriceBefDisc" select new XElement(k.Element(ns + "Value"))).First().Value),
                                  TotalDisc = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_TotalDisc" select new XElement(k.Element(ns + "Value"))).First().Value),
                                  DiscRate = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_DiscRate" select new XElement(k.Element(ns + "Value"))).First().Value),
                                  PriceAfterDisc = parservalues<double>((from k in t.Descendants(ns + "Field") where k.Element(ns + "Alias").Value == "U_PriceAfterDisc" select new XElement(k.Element(ns + "Value"))).First().Value),
                              }).ToList();

                var oCompanyServiceData = ConstVariables.oCompanyObject.GetCompanyService();
                var oGeneralServiceData = oCompanyServiceData.GetGeneralService("AIF_SALES_DISC");
                SAPbobsCOM.GeneralDataCollection oChildren = null;
                SAPbobsCOM.GeneralData oChild = null;
                var oGeneralDataData = ((SAPbobsCOM.GeneralData)(oGeneralServiceData.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                orec.DoQuery("Select MAX(\"DocEntry\")+1 from \"@AIF_SALES_DISC\"");

                docentry = orec.Fields.Item(0).Value.ToString();

                if (docentry == "")
                {
                    docentry = "1";
                }
                oGeneralDataData.SetProperty("DocNum", Convert.ToInt32(docentry));
                oGeneralDataData.SetProperty("U_PriceBefDisc", parservalues<double>(mainprice.ToString()).ToString());

                oChildren = oGeneralDataData.Child("AIF_SALES_DISC1");
                int i = 0;
                double price = 0;
                double discountRateLine = 0;
                double sumtotaldiscount = 0;
                double sumpriceafterdisc = 0;
                double subtotal = 0;
                foreach (var item in rows)
                {
                    oChild = oChildren.Add();

                    discountRateLine = parservalues<double>(item.Discount.ToString());
                    if (i == 0)
                    {
                        oChild.SetProperty("U_SubTotal", mainprice);
                    }
                    else
                    {
                        oChild.SetProperty("U_SubTotal", price);
                    }
                    oChild.SetProperty("U_DiscType", item.DiscType);
                    oChild.SetProperty("U_DiscRate", item.Discount);

                    if (item.DiscType == "1")
                    {
                        price = mainprice;
                    }

                    price = (price / 100) * discountRateLine;
                    sumtotaldiscount += price;
                    oChild.SetProperty("U_DiscTotal", price);

                    if (i == 0)
                    {
                        price = mainprice - price;
                    }
                    else
                    {
                        price = subtotal - price;
                    }

                    subtotal = price;
                    sumpriceafterdisc = subtotal;


                    oChild.SetProperty("U_SubTotal2", price);
                    i++;
                }


                oGeneralDataData.SetProperty("U_TotalDisc", sumtotaldiscount.ToString());
                oGeneralDataData.SetProperty("U_DiscRate", header.Select(x => x.DiscRate).FirstOrDefault());
                oGeneralDataData.SetProperty("U_PriceAfterDisc", subtotal);

                SAPbobsCOM.GeneralDataParams result = oGeneralServiceData.Add(oGeneralDataData);

            }
            catch (Exception ex)
            {
                return -1;
            }
            return docentry == "" ? -1 : Convert.ToInt32(docentry);
        }


        private static CultureInfo _parsCult;
        public static T parservalues<T>(string val) where T : struct
        {

            object returnvals = 0;
            try
            {
                if (string.IsNullOrEmpty(val)) return (T)Convert.ChangeType(returnvals, typeof(T));

                val = Regex.Replace(val, @"[^0-9.,]+", "").Trim();

                CultureInfo parser = parsCult;
                if (new Regex("^(?!0|\\.00)[0-9]+(,\\d{3})*([.][0-9]{0,9})$").IsMatch(val))
                {
                    parser = System.Globalization.CultureInfo.InvariantCulture;
                }

                if (typeof(T) == typeof(decimal))
                {
                    returnvals = decimal.Parse(val, parser);
                }
                else if (typeof(T) == typeof(double))
                {
                    returnvals = double.Parse(val, parser);
                }

                else if (typeof(T) == typeof(float))
                {
                    returnvals = float.Parse(val, parser);
                }
                else if (typeof(T) == typeof(int))
                {
                    returnvals = (int)decimal.Parse(val, parser);
                }
            }
            catch (Exception ex)
            {
            }

            return (T)returnvals;

        }

        private static CultureInfo parsCult
        {
            get
            {
                if (_parsCult == null)
                {
                    CultureInfo ci = CultureInfo.InvariantCulture;
                    _parsCult = (CultureInfo)ci.Clone();

                    _parsCult.NumberFormat.CurrencyDecimalSeparator = ",";
                    _parsCult.NumberFormat.NumberDecimalSeparator = ",";
                    _parsCult.NumberFormat.PercentDecimalSeparator = ",";
                    _parsCult.NumberFormat.CurrencyGroupSeparator = ".";
                    _parsCult.NumberFormat.NumberGroupSeparator = ".";
                    _parsCult.NumberFormat.PercentGroupSeparator = ".";
                }

                return _parsCult;

            }
        }
    }
}
