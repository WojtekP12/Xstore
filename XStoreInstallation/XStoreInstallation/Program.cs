using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.SqlServer.Server;

namespace XStoreInstallation
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Xstore Installer v.3.10.24   10.03.2016");
            Console.WriteLine("List of changes: \n The problem regarding the copying pos file has been fixed. (Hopefully)");
            SubMethods s = new SubMethods();

            s.LoadConfig();
            s.SetStoreCode();
            s.SetStoreType();
            s.SetEvironmentType();
            s.TestParameters();

            s.Install();

            Console.WriteLine("\n Have a nice day!");
            Console.ReadKey();
        }
    }
}
