using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ServiceTester
{
    class Program
    {
        static void Main(string[] args)
        {
            TASDownloadService.TASService serv = new TASDownloadService.TASService();
            serv.StartService();
            Console.ReadKey();
            serv.StopService();
        }
    }
}
