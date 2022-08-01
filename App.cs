using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.CompilerServices;
using System.IO;
using SpreadsheetReader.networking;
using SpreadsheetReader.SQL;
using SpreadSheetReader.controllers;
using SpreadsheetReader.models;

namespace SpreadsheetReader
{
    public class App
    { 
        public string connectionString { set; get; }
        public string tableName { set; get;  }

        public App(string path)
        {
            Setup();
            Run(path);
        }

        /* 
        * Setup Method
        * Sets the connection string & name of the table the app will be writing to
        */
        public void Setup()
        {

            string dir = new GetDirectory().dir;

            string connectionStringPath = @dir + @"\" + "SpreadsheetReader" + @"\" +"NetworkConfiguration.txt";

            string[] connectionArray = System.IO.File.ReadAllLines(connectionStringPath);
            string connectionStringFormatted = "";
            for (int i = 0; i < connectionArray.Length; i++)
                connectionStringFormatted += connectionArray[i];


            string tableNamePath = @dir + @"\" + "SpreadsheetReader" + @"\" + "DataTableConfiguration.txt";
            string pathNameTable = System.IO.File.ReadAllText(tableNamePath);
            string formattedPathNameTable = pathNameTable.Replace(" ", "");

            this.tableName = formattedPathNameTable;
            this.connectionString = connectionStringFormatted; 
        }

        /* 
        * Run Method
        * This will execute the sql query after creating a table object with rows
        */

        public void Run(string path)
        {
            ReadFromExcel reader = new ReadFromExcel(path, this.tableName);

            Table table;
            table = reader.Read();

            SQLFormatter sqlBuilder = new SQLFormatter(table);
            string sql = ""; 

            if (table.tableName == "errorTable")
                sql = sqlBuilder.InserQueryDynamic();
            else
                sql = sqlBuilder.InsertQueryBuilder();
            
            WriteToDB writer = new WriteToDB(sql, this.connectionString);

            writer.Write();
        }
    }
}
