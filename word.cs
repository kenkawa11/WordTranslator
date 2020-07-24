using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Net.Http;
using System.IO;
using AngleSharp.Html.Parser;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Task = System.Threading.Tasks.Task;
using System.Web;
using Microsoft.Office.Core;
using Shape = Microsoft.Office.Interop.Word.Shape;
using System.Security.Cryptography.X509Certificates;
using AngleSharp.Common;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Word.Application;

namespace WordTranslator
{
    class Word : IDisposable
    {
        private Microsoft.Office.Interop.Word.Application word;
        public Document document;
        public Word(string filePath)
        {
            word = new Application { Visible = true };
            document = word.Documents.Open(filePath);
        }

        public void Dispose()
        {
            document.Close();
            document = null;
            word.Quit();
            word = null;
        }
        public void Replace(string search, string replace)
        {
            Find find = word.Selection.Find;
            find.ClearFormatting();
            find.Text = search;
            find.Replacement.ClearFormatting();
            find.Replacement.Text = replace;
            find.Execute(Replace: WdReplace.wdReplaceAll);
        }
        public void SaveAs(string savePath)
        {
            document.SaveAs2(savePath);
        }
        public void SavePdf(string savePath)
        {
            document.SaveAs2(savePath, FileFormat: WdSaveFormat.wdFormatPDF);
        }
    }


    public abstract class BaseEngine
    {

        protected static IWebDriver driver;
        protected int maxLength = 300;

        public BaseEngine(string fn)
        {
            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            var options = new ChromeOptions();
            //options.AddArgument("--headless");
            driver = new ChromeDriver(driverService, options);
        }
        public async Task AsynWdProcess()
        {
            var word = new Word(@"C:\test\test.docx");
            var document = word.document;
            var para = document.Paragraphs;

            var inshapes = document.InlineShapes;
            for (var i = 1; i <= inshapes.Count; i++)
            {
                inshapes[i].ConvertToShape().WrapFormat.Type = WdWrapType.wdWrapTopBottom;
            }

            var shapeCol = document.Shapes;
            for (var i = 1; i <= shapeCol.Count; i++)
            {
                if (shapeCol[i].WrapFormat.Type == WdWrapType.wdWrapInline)
                {
                    shapeCol[i].WrapFormat.Type = WdWrapType.wdWrapTopBottom;
                }
            }

            for (var i = 1; i <= para.Count; i++)
            {
                var text = para[i].Range.Text;

                var delim = text.Substring(text.Length - 1);
                if (delim == "\n" || delim == "\r" || delim == "\a" || delim == "\f")
                {
                    if (text.Length > 1)
                    {
                        var delim2 = text.Substring(text.Length - 2, 1);
                        if (delim2 == "\n" || delim2 == "\r" || delim2 == "\a" || delim2 == "\f")
                        {
                            delim = delim2 + delim;
                        }
                    }

                }

                para[i].Range.Select();
                var engtext = await AsyncRetEngText(text);
                if(engtext!="")
                {
                    var awd = document.Range(para[i].Range.Start, para[i].Range.End - 1);
                    awd.Text = "";

                    awd.Select();
                    awd.Text = "";
                    para[i].Range.InsertBefore(engtext);
                }
            }

            var listTextshape = new List<Shape>();

            foreach(Shape v in shapeCol)
            {
                GetShapeText(v, ref listTextshape); 
            }

            foreach (Shape txtShape in listTextshape)
            {
                var shtext = txtShape.TextFrame.TextRange.Text;
                txtShape.TextFrame.TextRange.Text= await AsyncRetEngText(shtext);
            }
        }


        public void GetShapeText(Shape aShape, ref List<Shape> textlist)
        {
            aShape.Select();

            switch(aShape.Type)
            {
                case MsoShapeType.msoCanvas:
                     foreach(Shape v in aShape.CanvasItems)
                    {
                        GetShapeText(v,ref textlist);
                    }
                    break;

                case MsoShapeType.msoGroup:
                     foreach(Shape v in aShape.GroupItems)
                    {
                        GetShapeText(v,ref textlist);
                    }
                    break;
                default:
                    if (aShape.TextFrame != null && aShape.TextFrame.HasText != 0)
                    {
                        textlist.Add(aShape);
                    }
                    break;
            }
        }



        private void divideText(string text, ref List<string> divided)
        {
            int pos;
            pos = 0;
            while (true)
            {
                if(text.Length<=maxLength)
                {
                    divided.Add(text);
                    return;
                }

                var asd = text.Substring(0, maxLength);
                pos = asd.LastIndexOf("。");
                if (pos == -1)
                {
                    pos = asd.LastIndexOf("．");
                    if (pos == -1)
                    {
                        pos = asd.LastIndexOf(".");
                        if (pos == -1)
                        {
                            pos = asd.LastIndexOf("　");
                            if (pos == -1)
                            {
                                pos = asd.LastIndexOf(" ");
                                if (pos == -1)
                                {
                                    pos = maxLength-1;
                                }
                            }
                        }
                    }
                }
                divided.Add(text.Substring(0, pos+1));
                if(text.Length<=pos+1)
                {
                    return;
                }
                text = text.Substring(pos + 1);
            }
        }

        private async Task<string >AsyncRetEngText(string text)
        {
            var reg = new Regex("[\n　\r \t\vt\f\a]");
            var replaced = text.Trim();
            replaced = reg.Replace(replaced, "");

            var divided = new List<string>();
            
           


            if(replaced=="")
            {
                return "";
            }


            if(text.Length<maxLength)
            {
                var engtext = await AsyncTranslate(text);
                return engtext;

            }

            divideText(text, ref divided);

            var divengText = "";
            foreach (var v in divided)
            {
                var temp=await AsyncTranslate(v);
                divengText += temp;

            }
   
            return divengText;
            
        }
        public virtual async Task<string> AsyncTranslate(string text)
        {
            await Task.Delay(1);
            return text;
        }
    }

    public class Google:BaseEngine
    {
        private string targetUrl="https://translate.google.co.jp";
        private string lang= @"/#ja/en/";

        public Google(string fn) :base(fn)
        {

        }

        public async Task<string> AsyncTranslateSentence(string sntnc)
        {
            var regptn = "(tlid-translation translation.*?\"\">)<span.*?>(.+?)</span>";
            var targeturl = targetUrl +lang+ HttpUtility.UrlEncode(sntnc);
            driver.Navigate().GoToUrl(targeturl);
            var html = driver.PageSource;

            var reg = new Regex(regptn);

            await Task.Delay(300);

            var l = driver.FindElement(By.ClassName("result-shield-container")).Text;
            var m = reg.Match(html);
            return m.Groups[2].Value;

        }

        public async override Task<string> AsyncTranslate(string text)
        {
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();


            driver.Navigate().GoToUrl(targetUrl + lang);
            var cpbtnl = driver.FindElements(By.ClassName("copybutton")).Count;
            sw.Start();
            while(cpbtnl>0)
            {
                cpbtnl = driver.FindElements(By.ClassName("copybutton")).Count;
                await Task.Delay(100);
                if (driver.FindElement(By.Id("source")).Text=="")
                {
                    driver.FindElement(By.Id("source")).SendKeys("");
                    await Task.Delay(100);
                    driver.FindElement(By.ClassName("tlid-clear-source-text")).Click();
                }
                if (sw.ElapsedMilliseconds > 5000)
                {
                    sw.Stop();
                    sw.Reset();
                    return "";
                }     
            }

            sw.Stop();
            sw.Reset();

            driver.FindElement(By.Id("source")).SendKeys(text);
            sw.Start();


            while (driver.FindElements(By.ClassName("result-shield-container")).Count==0)
            {
                await Task.Delay(200);
                if (sw.ElapsedMilliseconds > 5000)
                {
                    sw.Stop();
                    sw.Reset();
                    return "";
                }
            }

            sw.Stop();
            sw.Reset();

            var translated = driver.FindElement(By.ClassName("result-shield-container")).Text;
            driver.FindElement(By.ClassName("tlid-clear-source-text")).Click();
            await Task.Delay(100);
            return translated;

        }

        public void DisposeGoogle()
        {
            driver.Quit();
        }
    }
}
