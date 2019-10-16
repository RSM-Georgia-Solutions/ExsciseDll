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
        /// <summary>
        /// Returns Dictionary that contains about success of posting and log details in inner Dictionary
        /// </summary>
        /// <param name="company"></param>
        /// <param name="invoiceDocEntry"></param>
        /// <param name="objectTypes"></param>
        /// <returns></returns>
        public static Dictionary<bool, Dictionary<string, List<string>>> CreateExciseEntry(Company company, int invoiceDocEntry, BoObjectTypes objectTypes)
        {
            string tableHeader = objectTypes == BoObjectTypes.oInvoices ? "OINV" : "OPCH";
            string tableRow = objectTypes == BoObjectTypes.oInvoices ? "INV1" : "PCH1";

            Recordset recSetAct = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSetAct.DoQuery($"Select * From [@RSM_EXCP]");
            string exciseAccount = recSetAct.Fields.Item("U_ExciseAcc").Value.ToString();
            string exciseAccountReturn = recSetAct.Fields.Item("U_ExciseAccReturn").Value.ToString();
            var roundAccuracy = company.GetCompanyService().GetAdminInfo().TotalsAccuracy;
            Recordset recSet2 = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);

            recSet2.DoQuery($@"SELECT {tableRow}.AcctCode, 
       U_EXCISE, 
       ROUND({tableRow}.Quantity * U_EXCISE,{roundAccuracy}) AS 'ExciseAmount', 
       {tableHeader}.Series, 
       {tableHeader}.DocNum, 
       {tableRow}.ItemCode, 
       {tableHeader}.DocDate, 
       {tableHeader}.BPLId, 
       {tableHeader}.DocCur
FROM {tableHeader}
     JOIN {tableRow} ON {tableHeader}.DocEntry = {tableRow}.DocEntry
     JOIN OITM ON OITM.ItemCode = {tableRow}.ItemCode
WHERE {tableHeader}.DocEntry = 1
      AND DocCur = 'GEL'
      AND OITM.U_EXCISE != 0
      AND {tableRow}.Quantity != 0
      AND {tableHeader}.DocType = 'I'
      AND OITM.ItemType = 'I'");

            Dictionary<string, List<string>> res = new Dictionary<string, List<string>>();
            while (!recSet2.EoF)
            {
                string glRevenueAccount = recSet2.Fields.Item("AcctCode").Value.ToString();
                double fullExcise = (double)recSet2.Fields.Item("ExciseAmount").Value;
                int resultJdt = 0;
                if (objectTypes == BoObjectTypes.oInvoices)
                {
                    resultJdt = AddJournalEntryCredit(company,
                        exciseAccount,
                        glRevenueAccount,
                        fullExcise,
                        (int)recSet2.Fields.Item("AcctCode").Value,
                        "IN" + recSet2.Fields.Item("DocNum").Value + " " + recSet2.Fields.Item("ItemCode").Value,
                        "",
                        DateTime.Parse(recSet2.Fields.Item("DocDate").Value.ToString(), CultureInfo.InvariantCulture),
                        (int)recSet2.Fields.Item("BPLId").Value);
                }
                else if (objectTypes == BoObjectTypes.oCreditNotes)
                {
                    resultJdt = AddJournalEntryCredit(company,
                        exciseAccountReturn,
                        glRevenueAccount,
                        -fullExcise,
                        (int)recSet2.Fields.Item("AcctCode").Value,
                        "CR" + recSet2.Fields.Item("DocNum").Value + " " + recSet2.Fields.Item("ItemCode").Value,
                        "",
                        DateTime.Parse(recSet2.Fields.Item("DocDate").Value.ToString(), CultureInfo.InvariantCulture),
                        (int)recSet2.Fields.Item("BPLId").Value);
                }
                recSet2.MoveNext();
            }

            return new Dictionary<bool, Dictionary<string, List<string>>> { { true, res } };

        }

        public static int AddJournalEntryCredit(Company _comp, string creditCode, string debitCode,
            double amount, int series, string comment, string code, DateTime DocDate, int BPLID = 235, string currency = "GEL")
        {
            JournalEntries vJE =
                (JournalEntries)_comp.GetBusinessObject(BoObjectTypes.oJournalEntries);

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

            return vJE.Add();

        }

    }
}
