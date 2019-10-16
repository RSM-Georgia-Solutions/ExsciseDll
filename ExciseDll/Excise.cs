using System;
using System.Globalization;
using SAPbobsCOM;
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
        public static int CreateExciseEntry(Company company, int invoiceDocEntry, BoObjectTypes objectTypes)
        {
            string tableHeader = objectTypes == BoObjectTypes.oInvoices ? "OINV" : "OPCH";
            string tableRow = objectTypes == BoObjectTypes.oInvoices ? "INV1" : "PCH1";
            var roundAccuracy = company.GetCompanyService().GetAdminInfo().TotalsAccuracy;

            Recordset recSet2 = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet2.DoQuery($"Select * From [@RSM_EXCP]");
            string exciseAccount = recSet2.Fields.Item("U_ExciseAcc").Value.ToString();
            string exciseAccountReturn = recSet2.Fields.Item("U_ExciseAccReturn").Value.ToString();

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

            int resultJdt = 0;


            while (!recSet2.EoF)
            {
                double fullExcise = (double)recSet2.Fields.Item("ExciseAmount").Value;

                var AcctCode = (string)recSet2.Fields.Item("AcctCode").Value;
                var Series = (int)recSet2.Fields.Item("Series").Value;
                var DocNum =  "IN" + recSet2.Fields.Item("DocNum").Value + " " + recSet2.Fields.Item("ItemCode").Value;
                var DocDate = DateTime.Parse(recSet2.Fields.Item("DocDate").Value.ToString(), CultureInfo.InvariantCulture);
                var BPLId = (int)recSet2.Fields.Item("BPLId").Value;

                if (objectTypes == BoObjectTypes.oInvoices)
                {
                    resultJdt = AddJournalEntryCredit(company,
                        exciseAccount,
                        AcctCode,
                        fullExcise, 
                        Series,
                        "IN" + recSet2.Fields.Item("DocNum").Value + " " + recSet2.Fields.Item("ItemCode").Value,
                        "",
                        DateTime.Parse(recSet2.Fields.Item("DocDate").Value.ToString(), CultureInfo.InvariantCulture),
                        (int)recSet2.Fields.Item("BPLId").Value);
                }
                else if (objectTypes == BoObjectTypes.oCreditNotes)
                {
                    resultJdt = AddJournalEntryCredit(company,
                        exciseAccountReturn,
                        AcctCode,
                        -fullExcise,
                        Series,
                        "CR" + recSet2.Fields.Item("DocNum").Value + " " + recSet2.Fields.Item("ItemCode").Value,
                        "",
                        DateTime.Parse(recSet2.Fields.Item("DocDate").Value.ToString(), CultureInfo.InvariantCulture),
                        (int)recSet2.Fields.Item("BPLId").Value);
                }
                recSet2.MoveNext();
            }

            return resultJdt;

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
