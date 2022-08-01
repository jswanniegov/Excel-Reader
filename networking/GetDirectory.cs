using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetReader.networking
{
    
    class GetDirectory
    {
        public string dir { get; }

        public GetDirectory()
        {
            this.dir = System.IO.Path.GetFullPath(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\.."));
        }

    }
}
