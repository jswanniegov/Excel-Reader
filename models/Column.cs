using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetReader.models
{
    public class Column
    {
        public string description { set; get; }

        public Column(string description)
        {
            this.description = description; 
        }
    }
}
