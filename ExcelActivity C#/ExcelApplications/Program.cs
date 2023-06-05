using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Automaiton
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Program Start");
            Application app = new Application();
            app.Visible = true;


        }
    }
}