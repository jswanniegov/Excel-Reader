using SpreadsheetReader.models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetReader.SQL
{
    
    public class SQLFormatter
    {
        public Table table { set; get; }

        public SQLFormatter(Table table)
        {
            this.table = table;
        }

        /*
        * Make use of the insert query builder function for non-error tables
        */
        public string InsertQueryBuilder()
        {
            string sql = "INSERT INTO " + this.table.tableName + "(vrcRefNum, vrcTaxType, vrcPeriod, vrcTaxYear, vrcRiskDescription, vrcRiskID, vrcCaseType, vrcSuppressCommunicationId, vrcCaseNum, vrcRequestOperation, vrcComments) VALUES";
            for (int i = 0; i < this.table.tableRows.Count; i++)
            {
                if (i == this.table.tableRows.Count - 1)
                    sql += this.table.tableRows[i].SQL() + ";";
                else
                    sql += this.table.tableRows[i].SQL() + ", ";
            }

            return sql;
        }

        public string InserQueryDynamic()
        {
            string tableName = this.table.tableName;
            string columnNames = "(";
            string values = ""; 

            for(int i = 0; i < this.table.tableRows.Count; i++)
            {
                for( int j = 0; j < this.table.tableRows[i].columnList.Count; j++)
                {
                    if(i == 0)
                    {
                        if (j == 0)
                            columnNames += this.table.tableRows[i].columnList[j].description + ", ";
                        else if (j == (this.table.tableRows[i].columnList.Count - 1))
                            columnNames += ((this.table.tableRows[i].columnList[j].description) + ")");
                        else
                            columnNames += ((this.table.tableRows[i].columnList[j].description) + ", ");
                    }
                    else
                    {
                        if (j == 0)
                            values += (" VALUES ('" +this.table.tableRows[i].columnList[j].description +"', ");
                        else if (j == (this.table.tableRows[i].columnList.Count - 1))
                            values += ("'" + (this.table.tableRows[i].columnList[j].description) + "')");
                        else
                            values += ("'"+(this.table.tableRows[i].columnList[j].description) + "', ");
                    }
                }
            }
            string sql = "INSERT INTO " + this.table.tableName + columnNames + values; 
            return sql;
        }
    }
}
