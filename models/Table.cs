using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/**
* Custom Class - Represents a table in the database
*/
namespace SpreadsheetReader.models
{
    public class Table
    {
        public Table(string tableName)
        {
            this.tableName = tableName; 
        }

        public string tableName { set; get; }
        public List<RowItem> tableRows { set; get; }
    }
}