using SpreadsheetReader.networking;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetReader
{
    class Program
    {
        static void Main(string[] args)
        {
            string path;
            if (args.Length == 0)
                path = @"C:\Users\S2028566\Desktop\Test\vrcTest5.xlsx";
            else
                path = args[0]; 
                

            if(path != "")
                new App(path);
                
            else
                new NetworkLogger("No path, or inoccrect path specified", true);
            
        }
    }
}
