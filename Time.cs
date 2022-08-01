using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetReader
{
    public class Time
    {
        public DateTime value { set; get; }
        public Time()
        {
            this.value = DateTime.Now;
        }
        public String GetTimestamp()
        {
            return this.value.ToString("MM/dd/yyyy hh:mm tt");
        }
    }
}
