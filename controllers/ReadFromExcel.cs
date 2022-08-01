using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.IO;
using SpreadsheetReader.models;
using SpreadsheetReader;

namespace SpreadSheetReader.controllers
{
    public class ReadFromExcel
    {
        public string path { set; get; }
        public Table table = new Table("");

        public ReadFromExcel(string path, string tableName)
        {
            this.path = path;
            this.table.tableName = tableName;
        }
        public Table Read()
        {
            /**
            *Use of Microsoft Office Interop
            */

            Excel.Application xlApp = new Excel.Application();
            try
            {
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@path);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                /**
                * Use of a custom class
                * RowItem is a custom class reperesenting a single row in a db table
                */

                List<RowItem> rowList = new List<RowItem>();
                List<string> emptyCells = new List<string>();

                if (colCount != 11)
                {
                    string message = "The table format is incorrect...";
                    ErrorTable(message);
                    return table;
                }

                for (int i = 2; i <= rowCount; i++)
                {
                    RowItem row = new RowItem();

                    for (int j = 1; j <= colCount; j++)
                    {
                        string columnName = xlRange.Cells[1, j].Value2.ToString();
                        columnName = columnName.Replace(" ", "").ToLower();
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            switch (columnName)
                            {
                                case "referencenumber":
                                    row.vrcRefNum = xlRange.Cells[i, j].Value2.ToString();
                                    break;
                                case "taxtype":
                                    row.vrcTaxType = xlRange.Cells[i, j].Value2.ToString();
                                    break;

                                case "period":
                                    row.vrcPeriod = xlRange.Cells[i, j].Value2.ToString();
                                    break;

                                case "taxyear":
                                    row.vrcTaxYear = xlRange.Cells[i, j].Value2.ToString();
                                    break;
                                case "riskdescription":
                                    row.vrcRiskDescription = xlRange.Cells[i, j].Value2.ToString();
                                    break;

                                case "riskid":
                                    row.vrcRiskID = xlRange.Cells[i, j].Value2.ToString();
                                    break;

                                case "casetype":
                                    row.vrcCaseType = xlRange.Cells[i, j].Value2.ToString();
                                    break;

                                case "suppresscommunicationid":
                                    row.vrcSuppressCommunicationInd = xlRange.Cells[i, j].Value2.ToString();
                                    break;

                                case "casenumber":
                                    row.vrcCaseNum = xlRange.Cells[i, j].Value2.ToString();
                                    break;

                                case "requestoperation":
                                    row.vrcRequestOperation = xlRange.Cells[i, j].Value2.ToString();
                                    break;
                                case "comments":
                                    row.vrcComments = xlRange.Cells[i, j].Value2.ToString();
                                    break;

                                default:
                                    break;
                            }
                        }
                        else
                            emptyCells.Add((xlWorksheet.Columns[j].Address).Replace("$", "") + i);
                    }
                    rowList.Add(row);
                }
                if (emptyCells.Count == 0)
                {
                    table.tableRows = rowList;
                    return table;
                }

                else
                {
                    /*
                    * Format the the message for the empty cell
                    */

                    string cells = "";
                    for (int i = 0; i < emptyCells.Count; i++)
                    {
                        if (i == 0)
                            cells += emptyCells[i];
                        else
                            cells += ", " + emptyCells[i];
                    }
                    string message = "Please double check the rows & columns! It seems that CELLS: " + cells + " do not have values...";

                    ErrorTable(message);
                    return table;
                }
            }
            catch (Exception e)
            {
                string message = "Error occured while reading the file... " + e.ToString();

                ErrorTable(message);
                return table;
            }
        }

        public void ErrorTable(string m)
        {
            this.table.tableName = "errorTable";

            List<RowItem> rowList = new List<RowItem>();
            RowItem row1 = new RowItem();

            Column columnName1 = new Column("consoleLog");
            Column columnName2 = new Column("errorMessage");
            row1.columnList.Add(columnName1);
            row1.columnList.Add(columnName2);


            RowItem row2 = new RowItem();
            Time time = new Time();
            String timeStamp = time.GetTimestamp();

            Column console = new Column(timeStamp);
            Column decription = new Column(m);
            row2.columnList.Add(console);
            row2.columnList.Add(decription);

            rowList.Add(row1);
            rowList.Add(row2);

            table.tableRows = rowList;
        }
    }
}