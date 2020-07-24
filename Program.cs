using System;
using System.Windows.Forms;
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
            var TransEngine= new DeepL();
            await TransEngine.AsynWdProcess(@"C:\test\test.docx");
        }
    }
}
