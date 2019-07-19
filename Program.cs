using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StockStatusMPOUS
{
    class Program
    {
        private static string locations = "";
        static void Main(string[] args)
        {
            locations = args[0];
            Process();
        }

        private static void Process()
        {
            Process proc = new Process(locations);

        }
    }
}
