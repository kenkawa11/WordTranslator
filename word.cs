using System;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Task = System.Threading.Tasks.Task;
using System.Web;
using Microsoft.Office.Core;
using Shape = Microsoft.Office.Interop.Word.Shape;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Web.ModelBinding;
using System.Runtime.InteropServices;

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
        protected long timelimit = 5000;

        public BaseEngine()
        {
            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            var options = new ChromeOptions();
            //options.AddArgument("--headless");
            driver = new ChromeDriver(driverService, options);
        }
        public async Task AsynWdProcess(string fn)
        {
            var word = new Word(fn);
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

            string dir = Path.GetDirectoryName(fn);
            string FileName = Path.GetFileNameWithoutExtension(fn);
            word.SaveAs(dir + "translated_" + FileName + ".docx");

            word.Dispose();
        }


        protected void GetShapeText(Shape aShape, ref List<Shape> textlist)
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



        protected void DivideText(string text, ref List<string> divided)
        {
            int pos;
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

        protected async Task<string >AsyncRetEngText(string text)
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

            DivideText(text, ref divided);

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


        private string CheckDelim(string text)
        {
            string delim = text.Substring(text.Length - 1);
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
            return delim;
        }

        public void Dispose()
        {
            driver.Quit();
        }
    }


    public class Google:BaseEngine
    {
        private string targetUrl=@"https://translate.google.co.jp";
        private string lang= @"/#ja/en/";
        
        public Google()
        {
            maxLength = 4500;
            var handle = driver.WindowHandles;
        }
        public async Task<string> AsyncTranslateUrlGet(string sntnc)
        {
            var regptn = "(tlid-translation translation.*?\"\">)<span.*?>(.+?)</span>";
            var targeturl = targetUrl +lang+ HttpUtility.UrlEncode(sntnc);
            driver.Navigate().GoToUrl(targeturl);
            var html = driver.PageSource;
            var reg = new Regex(regptn);
            await Task.Delay(300);
            //var l = driver.FindElement(By.ClassName("result-shield-container")).Text;
            var m = reg.Match(html);
            return m.Groups[2].Value;
        }

        public async override Task<string> AsyncTranslate(string text)
        {
            Stopwatch sw = new Stopwatch();

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
                if (sw.ElapsedMilliseconds > timelimit)
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
                if (sw.ElapsedMilliseconds > timelimit)
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

    }
    public class DeepL:BaseEngine
    {
        private string targetUrl = @"https://www.deepl.com/ja/translator#en/ja/";
        public DeepL()
        {
            maxLength = 3000;
        }
        public async override Task<string> AsyncTranslate(string text)
        {
            Stopwatch sw = new Stopwatch();
            driver.Navigate().GoToUrl(targetUrl);
            var  JudgeLength = 30;

 
            driver.FindElement(By.ClassName("lmt__source_textarea")).SendKeys(text);
            var button_css = "div.lmt__target_toolbar__copy button";

            var button = driver.FindElement(By.CssSelector(button_css));


            //Clipboard.Clear();
            //string ClipText="";
            //while(ClipText=="")
            //{
            //    button.Click();
            //    await Task.Delay(500);
            //    if(Clipboard.ContainsText())
            //    {
            //        ClipText = Clipboard.GetText();
            //    }       
            //}

            //var translated = ClipText;

            





            //string s = "";
            //string existText = "";
            //sw.Start();
            //while (s == "" || existText == prev)
            //{
            //    s = driver.FindElement(By.ClassName("lmt__target_textarea")).Text;
            //    if(s.Length> JudgeLength)
            //    {
            //        existText=s.Substring(0,JudgeLength);
            //    }
            //    else
            //    {
            //        existText = s;
            //    }

            //    await Task.Delay(200);
            //    if (sw.ElapsedMilliseconds > timelimit)
            //    {
            //        sw.Stop();
            //        sw.Reset();
            //        return "";
            //    }
            //} 


            //sw.Stop();
            //sw.Reset();



            //await Task.Delay(3000);
            //var translatedCol = driver.FindElements(By.ClassName("lmt__target_textarea"));
            ////var col= driver.FindElement(By.ClassName("lmt__target_textarea")).Text;
            //var translated = driver.FindElement(By.ClassName("lmt__target_textarea")).Text;
            //var translated2 = driver.FindElement(By.CssSelector("#dl_translator > div.lmt__sides_container > div.lmt__side_container.lmt__side_container--target > div.lmt__textarea_container > div.lmt__inner_textarea_container > textarea")).Text;
            //var translated3 = driver.FindElements(By.XPath("//*[@id='dl_translator']/div[1]/div[4]/div[3]/div[1]/textarea"));
            //var transcol2 = driver.FindElements(By.ClassName("lmt__textarea"));

            //var translated = translatedCol[1].Text;

            if (driver.FindElements(By.ClassName("lmt__clear_text_button")).Count!=0)
            {
                driver.FindElement(By.ClassName("lmt__clear_text_button")).Click();
            }

            await Task.Delay(200);


            return translated;
        }

    }
}
