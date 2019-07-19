using System;
using System.Collections.Generic;
using System.Globalization;
using SAPbobsCOM;
using Translator;
using Company = SAPbobsCOM.Company;

namespace ExciseDll
{
    public class Excise
    {
        private static string QueryHanaTransalte(string query, bool isHana)
        {
            if (isHana)
            {
                int numOfStatements;
                int numOfErrors;
                TranslatorTool translateTool = new TranslatorTool();
                query = translateTool.TranslateQuery(query, out numOfStatements, out numOfErrors);
                return query;
            }
            return query;
        }

        /// <summary>
        /// Returns Dictionary that contains about success of posting and log details in inner Dictionary
        /// </summary>
        /// <param name="company"></param>
        /// <param name="invoiceDocEntry"></param>
        /// <returns></returns>
        public static Dictionary<bool, Dictionary<string, List<string>>> CreateExciseEntryForInovice(Company company, int invoiceDocEntry)
        {
            bool isHana = company.DbServerType.ToString() == "dst_HANADB";
            Documents invoice = (Documents)company.GetBusinessObject(BoObjectTypes.oInvoices);
            invoice.GetByKey(invoiceDocEntry);
            Recordset recSetAct = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSetAct.DoQuery(QueryHanaTransalte($"Select * From [@RSM_EXCP]", isHana));
            string exciseAccount = recSetAct.Fields.Item("U_ExciseAcc").Value.ToString();
            string exciseAccountReturn = recSetAct.Fields.Item("U_ExciseAccReturn").Value.ToString();

            var x = new List<string>() { "აქციზის ანგარიში არ არის განსაზღვრული" };

            if (string.IsNullOrWhiteSpace(exciseAccount))
            {
                return new Dictionary<bool, Dictionary<string, List<string>>>
                {
                    { false, new Dictionary<string, List<string>> { { invoiceDocEntry.ToString(), new List<string> { "აქციზის ანგარიში არ არის განსაზღვრული" } } } }
                };
            }

            if (invoice.DocCurrency != "GEL")
            {
                return new Dictionary<bool, Dictionary<string, List<string>>>
                {
                    { true, new Dictionary<string, List<string>> { { invoiceDocEntry.ToString(), new List<string> { "Currency Must Be GEL" } } } }
                };
            }

            Dictionary<string, List<string>> res = new Dictionary<string, List<string>>();
            for (int i = 0; i < invoice.Lines.Count; i++)
            {
                invoice.Lines.SetCurrentLine(i);
                string glRevenueAccount = invoice.Lines.AccountCode;
                SAPbobsCOM.Items item = (SAPbobsCOM.Items)company.GetBusinessObject(BoObjectTypes.oItems);
                item.GetByKey(invoice.Lines.ItemCode);
                string exciseString = string.Empty;
                try
                {
                    exciseString = item.UserFields.Fields.Item("U_Excise").Value.ToString();
                }
                catch (Exception)
                {
                    if (res.ContainsKey(invoice.Lines.ItemCode))
                    {
                        res[invoice.Lines.ItemCode].Add("Excise UDF დასამატებელია");
                        return new Dictionary<bool, Dictionary<string, List<string>>>
                        {
                            { false, res }
                        };
                    }
                    res.Add(invoice.Lines.ItemCode, new List<string> { "Excise UDF დასამატებელია" });
                    return new Dictionary<bool, Dictionary<string, List<string>>>
                    {
                        { false, res }
                    };
                }

                double excise = double.Parse(exciseString, CultureInfo.InvariantCulture);

                if (string.IsNullOrWhiteSpace(exciseString) || excise == 0)
                {
                    if (res.ContainsKey(invoice.Lines.ItemCode))
                    {
                        res[invoice.Lines.ItemCode].Add("საქონელზე აქციზის განაკვეთი არ არის მითითებული");
                    }
                    else
                    {
                        res.Add(invoice.Lines.ItemCode, new List<string> { "საქონელზე აქციზის განაკვეთი არ არის მითითებული" });
                    }
                    continue;
                }


                if (invoice.Lines.Quantity == 0)
                {
                    if (res.ContainsKey(invoice.Lines.ItemCode))
                    {
                        res[invoice.Lines.ItemCode].Add("საქონლის რაოდენობა უდრის 0");
                    }
                    else
                    {
                        res.Add(invoice.Lines.ItemCode, new List<string> { "საქონლის რაოდენობა უდრის 0" });
                    }
                    continue;
                }

                var roundAccuracy = company.GetCompanyService().GetAdminInfo().TotalsAccuracy;
                double fullExcise = Math.Round(invoice.Lines.Quantity * excise, roundAccuracy);

                string resultJdt = AddJournalEntryCredit(company, exciseAccount, glRevenueAccount, fullExcise, invoice.Series,
                   "IN" + invoice.DocNum + " " + invoice.Lines.ItemCode, "", invoice.DocDate, invoice.BPL_IDAssignedToInvoice, invoice.DocCurrency);

                if (res.ContainsKey(invoice.Lines.ItemCode))
                {
                    res[invoice.Lines.ItemCode].Add(resultJdt);
                }
                else
                {
                    res.Add(invoice.Lines.ItemCode, new List<string> { resultJdt });
                }

            }
            return new Dictionary<bool, Dictionary<string, List<string>>> { { true, res } };

        }

        public static Dictionary<bool, Dictionary<string, List<string>>> CreateExciseEntryForCreditMemo(Company company, int invoiceDocEntry)
        {
            bool isHana = company.DbServerType.ToString() == "dst_HANADB";
            Documents invoice = (Documents)company.GetBusinessObject(BoObjectTypes.oCreditNotes);
            invoice.GetByKey(invoiceDocEntry);
            Recordset recSetAct = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSetAct.DoQuery(QueryHanaTransalte($"Select * From [@RSM_EXCP]", isHana));
            string exciseAccountReturn = recSetAct.Fields.Item("U_ExciseAccReturn").Value.ToString();

            var x = new List<string>() { "აქციზის ანგარიში არ არის განსაზღვრული" };

            if (string.IsNullOrWhiteSpace(exciseAccountReturn))
            {
                return new Dictionary<bool, Dictionary<string, List<string>>>
                {
                    { false, new Dictionary<string, List<string>> { { invoiceDocEntry.ToString(), new List<string> { "აქციზის ანგარიში არ არის განსაზღვრული" } } } }
                };
            }

            if (invoice.DocCurrency != "GEL")
            {
                return new Dictionary<bool, Dictionary<string, List<string>>>
                {
                    { true, new Dictionary<string, List<string>> { { invoiceDocEntry.ToString(), new List<string> { "Currency Must Be GEL" } } } }
                };
            }

            Dictionary<string, List<string>> res = new Dictionary<string, List<string>>();
            for (int i = 0; i < invoice.Lines.Count; i++)
            {
                invoice.Lines.SetCurrentLine(i);
                string glRevenueAccount = invoice.Lines.AccountCode;
                Items item = (Items)company.GetBusinessObject(BoObjectTypes.oItems);
                item.GetByKey(invoice.Lines.ItemCode);
                string exciseString = string.Empty;
                try
                {
                    exciseString = item.UserFields.Fields.Item("U_Excise").Value.ToString();
                }
                catch (Exception)
                {
                    if (res.ContainsKey(invoice.Lines.ItemCode))
                    {
                        res[invoice.Lines.ItemCode].Add("Excise UDF დასამატებელია");
                        return new Dictionary<bool, Dictionary<string, List<string>>>
                        {
                            { false, res }
                        };
                    }
                    res.Add(invoice.Lines.ItemCode, new List<string> { "Excise UDF დასამატებელია" });
                    return new Dictionary<bool, Dictionary<string, List<string>>>
                    {
                        { false, res }
                    };
                }

                double excise = double.Parse(exciseString, CultureInfo.InvariantCulture);

                if (string.IsNullOrWhiteSpace(exciseString) || excise == 0)
                {
                    if (res.ContainsKey(invoice.Lines.ItemCode))
                    {
                        res[invoice.Lines.ItemCode].Add("საქონელზე აქციზის განაკვეთი არ არის მითითებული");
                    }
                    else
                    {
                        res.Add(invoice.Lines.ItemCode, new List<string> { "საქონელზე აქციზის განაკვეთი არ არის მითითებული" });
                    }
                    continue;
                }


                if (invoice.Lines.Quantity == 0)
                {
                    if (res.ContainsKey(invoice.Lines.ItemCode))
                    {
                        res[invoice.Lines.ItemCode].Add("საქონლის რაოდენობა უდრის 0");
                    }
                    else
                    {
                        res.Add(invoice.Lines.ItemCode, new List<string> { "საქონლის რაოდენობა უდრის 0" });
                    }
                    continue;
                }

                var roundAccuracy = company.GetCompanyService().GetAdminInfo().TotalsAccuracy;
                double fullExcise = Math.Round(invoice.Lines.Quantity * excise, roundAccuracy);

                string resultJdt = AddJournalEntryCredit(company, exciseAccountReturn, glRevenueAccount, -fullExcise, invoice.Series,
                   "CR" + invoice.DocNum + " " + invoice.Lines.ItemCode, "", invoice.DocDate, invoice.BPL_IDAssignedToInvoice, invoice.DocCurrency);

                if (res.ContainsKey(invoice.Lines.ItemCode))
                {
                    res[invoice.Lines.ItemCode].Add(resultJdt);
                }
                else
                {
                    res.Add(invoice.Lines.ItemCode, new List<string> { resultJdt });
                }

            }
            return new Dictionary<bool, Dictionary<string, List<string>>> { { true, res } };

        }

        public static string AddJournalEntryCredit(SAPbobsCOM.Company _comp, string creditCode, string debitCode,
            double amount, int series, string comment, string code, DateTime DocDate, int BPLID = 235, string currency = "GEL")
        {
            SAPbobsCOM.JournalEntries vJE =
                (SAPbobsCOM.JournalEntries)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

            vJE.ReferenceDate = DocDate;
            vJE.DueDate = DocDate;
            vJE.TaxDate = DocDate;
            vJE.Memo = comment;

            vJE.Lines.BPLID = BPLID; //branch

            if (currency == "GEL")
            {
                vJE.Lines.Debit = amount;
                vJE.Lines.FCDebit = 0;
            }
            else
            {
                vJE.Lines.FCCurrency = currency;
                vJE.Lines.FCDebit = amount;
            }

            vJE.Lines.Credit = 0;
            vJE.Lines.FCCredit = 0;

            vJE.Lines.AccountCode = debitCode;
            vJE.Lines.Add();


            vJE.Lines.BPLID = BPLID;
            vJE.Lines.AccountCode = creditCode;
            vJE.Lines.Debit = 0;
            vJE.Lines.FCDebit = 0;
            if (currency == "GEL")
            {
                vJE.Lines.Credit = amount;
                vJE.Lines.FCCredit = 0;
            }
            else
            {
                vJE.Lines.FCCurrency = currency;
                vJE.Lines.FCCredit = amount;
            }

            vJE.Lines.Add();

            int i = vJE.Add();
            if (i == 0)
            {
                string transId = _comp.GetNewObjectKey();
                return transId;
            }
            else
            {
                throw new Exception(_comp.GetLastErrorDescription());
            }
        }

    }
}
