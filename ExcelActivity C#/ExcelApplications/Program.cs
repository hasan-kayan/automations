using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// For Excel applications C# has  Microsoft.Office.Interop.Excel namespace, namespace using interface architecture
// Namespace includes Application class which represents Excel Application
// 





namespace Excel_Automaiton
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Program Start");
            Application app = new Application(); // Creating new instance of the ' Application Class '. If you have to declare again remember you it will give error 
            // Go on the error and choose namespace 
            app.Visible = true; // Set visible property of the "Application" object to "true"
            


        }
    }
}