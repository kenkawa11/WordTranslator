using AngleSharp.Media;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


namespace WordTranslator
{
    class Program
    {
        static void Main(string[] args)
        {
            MainAsync().Wait();
        }

        private static async System.Threading.Tasks.Task MainAsync()
        {
            var TransEngine= new Google(@"C:\test\test.docx");
            await TransEngine.AsynWdProcess();
            var iii = 1;
        }
    }
}
