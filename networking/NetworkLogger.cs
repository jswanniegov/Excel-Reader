using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetReader.networking
{
    public class NetworkLogger
    { 
        public NetworkLogger(string m, bool error)
        {
            if (error)
                writeText(m, "ConnectionErrorLog");
            else
                writeText(m, "ConnectionLog");
        }

        public void writeText(string m, string file)
        {
            string dir = new GetDirectory().dir;

            string errorfilePath = @dir + @"\" + "SpreadsheetReader" + @"\" + "networking" + @"\" + file + ".txt";

            Time time = new Time();

            using (StreamWriter sw = File.AppendText(errorfilePath))
            {
                sw.WriteLine("-> " + time.GetTimestamp() + "; " + m);
            }
        }
    }
}
