using System;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Ghostscript.NET;
using Ghostscript.NET.Rasterizer;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OpenQA.Selenium;
using TechTalk.SpecFlow;
using Tesseract;

namespace WebApp
{
    public class FunctionRequest
    {
        public string name { get; set; }
        public string element { get; set; }
        public string elementone { get; set; }
        public string elementtwo { get; set; }
        public string elementthree { get; set; }
        public string elementfour { get; set; }
        public string elementfive { get; set; }
        public string elementsix { get; set; }
    }

    public class FunctionResponse
    {
        public string Value { get; set; }
        public string Valueone { get; set; }
        public string Valuetwo { get; set; }
        public string Valuethree { get; set; }
        public string Valuefour { get; set; }
        public string Valuefive { get; set; }
        public string Valuesix { get; set; }
    }






    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            // parse query parameter
            //string element = req.GetQueryNameValuePairs()
            // .FirstOrDefault(q => string.Compare(q.Key, "element", true) == 0).Value;

            string myJson = "{'element': 'line_2_13','elementone':'line_2_14'}";
            var _object = JsonConvert.DeserializeObject<FunctionRequest>(myJson);
            string element = _object.element;
            string elementone = _object.elementone;



            //Get request body
            dynamic data = await req.Content.ReadAsAsync<object>();

            //Set name to query string or body 


            Ocr Obj = new Ocr();
           
            // PDFContentPost Post = new PDFContentPost();
            string actual_CCC_InterestRates = Ocr.OCR_PDF_READ("SPLIT", "2", element);
            string actual_CCC_InterestRatesone = Ocr.OCR_PDF_READ_NO_Space("SPLIT", "2", elementone);


            if (Obj == null)
            {
                // Get request body
                //dynamic data = await req.Content.ReadAsAsync<object>();
               // element = data?.name;
            }

            var response = new FunctionResponse();
            response.Value = actual_CCC_InterestRates;
            response.Valueone = actual_CCC_InterestRatesone;

            return element == null
                ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
                : req.CreateResponse(HttpStatusCode.OK, response);



        }















        public class Ocr
        {
            // Base core = new Base();
            //public static string SolutionPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Locati‌​on)));
            public static string SolutionPath = "C:\\MASTER\\Azure\\App01\\App01\\App01\\bin";

            public static void ConvertorFromPdFtoData(string filelocationlocation, string datafilelocation)
            {
                Stream str = File.OpenRead(filelocationlocation);
                string tessdataPath = SolutionPath + "\\Tessdata\\";
                var engine = new TesseractEngine(tessdataPath, "eng", EngineMode.Default);
                for (int i = 1; i <= GetPdfPageCount(str); i++)
                {
                    using (var process = engine.Process(Pix.LoadTiffFromMemory(PdfToTiff(str, i, datafilelocation))))
                    {
                        //File.WriteAllText(string.Format(datafilelocation, i, "txt"), process.GetText());
                        File.WriteAllText(string.Format(datafilelocation, i, "html"), process.GetHOCRText(1));
                        //File.WriteAllText(string.Format(datafilelocation, i, "xlsx"), process.GetHOCRText(1));
                    }
                }

            }


            public static int GetPdfPageCount(Stream data)
            {
                string path64 = (SolutionPath + "\\gsdll64.dll");
                string path32 = (SolutionPath + "\\gsdll32.dll");
                {

                    try
                    {
                        GhostscriptVersionInfo gvi = new GhostscriptVersionInfo(path64);
                        GhostscriptRasterizer pdfengine = new GhostscriptRasterizer();
                        pdfengine.Open(data, gvi, true);
                        return pdfengine.PageCount;
                    }
                    catch
                    {
                        GhostscriptVersionInfo gviCatch = new GhostscriptVersionInfo(path32);
                        GhostscriptRasterizer pdfenginecatch = new GhostscriptRasterizer();

                        pdfenginecatch.Open(data, gviCatch, true);
                        return pdfenginecatch.PageCount;

                    }
                }


            }
            public static byte[] PdfToTiff(Stream data, int i, string datafilelocation)
            {
                string path64 = (SolutionPath + "\\gsdll64.dll");
                string path32 = (SolutionPath + "\\gsdll32.dll");
                try
                {
                    GhostscriptVersionInfo gvi = new GhostscriptVersionInfo(path64);
                    using (var pdfengine = new GhostscriptRasterizer())
                    {
                        pdfengine.Open(data, gvi, true);
                        var image = pdfengine.GetPage(300, 300, i);
                        var outputPngPath = Path.Combine(string.Format(datafilelocation, i, "tiff"));
                        var result = new MemoryStream();
                        image.Save(result, System.Drawing.Imaging.ImageFormat.Tiff);
                        result.Position = 0;
                        return result.ToArray();
                    }
                }
                catch
                {
                    GhostscriptVersionInfo gvi = new GhostscriptVersionInfo(path32);
                    using (var pdfengine = new GhostscriptRasterizer())
                    {
                        pdfengine.Open(data, gvi, true);
                        var image = pdfengine.GetPage(300, 300, i);
                        var outputPngPath = Path.Combine(string.Format(datafilelocation, i, "tiff"));
                        var result = new MemoryStream();
                        image.Save(result, System.Drawing.Imaging.ImageFormat.Tiff);
                        result.Position = 0;
                        return result.ToArray();
                    }
                }
            }


            public static string ReadImage(string filelocationlocation, string PDFType, string PDFpage, string xpath)
            {
                //string path = "C:/demo/demoZ.html";
                string sub = filelocationlocation + "\\" + PDFType + "__Page -";
                string fullpath = sub + PDFpage + ".html";

                StreamReader theFile = new StreamReader(fullpath);
                string content = theFile.ReadToEnd();
                theFile.Close();

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(content);
                string xpathroot = ("//span[@id='" + xpath + "']");

                var myDiv = doc.DocumentNode.SelectNodes(xpathroot);
                string hhh = (myDiv[0].InnerText);
                return hhh;
            }

            public static string ReadImageParagraph(string filelocationlocation, string PDFType, string PDFpage, string xpath)
            {
                //string path = "C:/demo/demoZ.html";
                string sub = filelocationlocation + "\\" + PDFType + "__Page -";
                string fullpath = sub + PDFpage + ".html";

                StreamReader theFile = new StreamReader(fullpath);
                string content = theFile.ReadToEnd();
                theFile.Close();

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(content);
                //string xpathroot = ("//span[@id='" + xpath + "']");

                string xpathroot = ("//p[@id='" + xpath + "']");

                var myDiv = doc.DocumentNode.SelectNodes(xpathroot);
                string hhh = (myDiv[0].InnerText);
                string rrr = hhh.Replace(" \n", string.Empty).Replace("\n", string.Empty).Replace("       ", string.Empty).Trim();


                return rrr;
            }
            //?? PROJECT SPECIFIC DATA ????


            public class PDFContentPost
            {
                public static string FirstLine { get; set; } = "line_2_12";
                public static string AddressFirstLine { get; set; } = "line_2_11";
                public static string AdequateInformation { get; set; } = "line_2_20";
                public static string refuseAUthorise { get; set; } = "par_2_22";
                public static string GeneralCCCFeesInterest { get; set; } = "par_2_8";

                public static string CCAInterestRates { get; set; } = "par_2_9";
                public static string CCACreditLimit { get; set; } = "par_2_16";

                public static string CCATAP { get; set; } = "par_2_8";
                public static string CCABalanceTransferInterest { get; set; } = "par_2_33";

                //START Concat NewDay Address
                public static string CCANewDayAddress0 { get; set; } = "word_2_46";
                public static string CCANewDayAddress1 { get; set; } = "word_2_47";
                public static string CCANewDayAddress2 { get; set; } = "word_2_48";
                public static string CCANewDayAddress3 { get; set; } = "word_2_49";
                public static string CCANewDayAddress4 { get; set; } = "word_2_50";
                public static string CCANewDayAddress5 { get; set; } = "word_2_51";
                //END Concat NewDay Address

                //Account Holder Name /Address
                public static string AccountHolderName { get; set; } = "par_2_4";


                //CI Name /Address
                public static string CINameAddress { get; set; } = "par_2_5";


                //Rates
                public static string Rate_Purchase_Standard_rate_per_annum_simple_rate { get; set; } = "word_2_381";
                public static string Rate_Purchase_Standard_rate_per_annum_Compond_Rate { get; set; } = "word_2_385";

                public static string Rate_CashAdvance_Simple_Rate { get; set; } = "word_2_334";
                public static string Rate_CashAdvance_CompoundRate { get; set; } = "word_2_338";

                public static string Rate_MoneyTransfer_Simple_Rate { get; set; } = "word_2_302";
                public static string Rate_MoneyTransfer_CompoundRate { get; set; } = "word_2_306";

                public static string Rate_Rurchase_Promotion { get; set; } = "word_2_310";

                public static string Rate_BalanceTransfer_Promotion { get; set; } = "word_2_359";

                public static string Rate_Purchase_Promotional_Preriod { get; set; } = "word_2_3";
                public static string Rate_BT_Preriod { get; set; } = "par_2_2";




                //Satic Data

                public static string ChargesPercentage { get; set; } = "word_2_178";

                public static string ChargesTransactionFees { get; set; } = "word_2_229";

                public static string LatePaymentFees { get; set; } = "word_2_327";

                public static string OverlimitFees { get; set; } = "word_2_327";

                public static string ReturnedPaymentFees { get; set; } = "word_2_381";

                public static string TraceFees { get; set; } = "word_2_392";

                public static string TransactionFeeForiegn { get; set; } = "word_2_497";

                public static string StatementCopy { get; set; } = "word_2_521";

                public static string RightofWithdrawals { get; set; } = "word_2_48";

                public static string WithDrawalsPhone { get; set; } = "line_2_7";

                public static string ComplaintsPhone { get; set; } = "line_2_26";

                public static string ComplaintsEmail { get; set; } = "word_2_423";

                public static string CompanyRegistrationNumber { get; set; } = "word_2_182";

            }



            public static string PDF_CCA_GET_AccountHolderName()
            {
                string Obj = OCR_PDF_READ_NO_Space("Credit_card_agreement", "1", PDFContentPost.AccountHolderName);
                string[] SObj = Obj.Split(new string[] { " and " }, StringSplitOptions.None);
                string sValueArray = SObj[1];
                string sElement = Regex.Replace(sValueArray, @"(°)", "").Replace(" ", string.Empty);
                return sElement;
            }


            public static string PDF_CCA_GET_NewDayAddress()
            {
                string Obj_NewDayAddress = OCR_PDF_READ_NO_Space("Credit_card_agreement", "1", PDFContentPost.AccountHolderName);
                string[] SObj_NewDayAddress = Obj_NewDayAddress.Split(new string[] { " of ", " and " }, StringSplitOptions.None);
                string sElement = SObj_NewDayAddress[1];
                return sElement;
            }

            public static string PDF_CCA_CI_Address()
            {
                string Obj_NewDayAddress = OCR_PDF_READ_NO_Space("Credit_card_agreement", "1", PDFContentPost.CINameAddress);
                string[] SObj_NewDayAddress = Obj_NewDayAddress.Split(new string[] { " of " }, StringSplitOptions.None);
                string sElement = SObj_NewDayAddress[1];
                return sElement;
            }


            public static string PDF_CCA_CI_Name()
            {
                string Obj_NewDayAddress = OCR_PDF_READ_NO_Space("Credit_card_agreement", "1", PDFContentPost.CINameAddress);
                string[] SObj_NewDayAddress = Obj_NewDayAddress.Split(new string[] { " was ", " of " }, StringSplitOptions.None);
                string sElement = SObj_NewDayAddress[1];
                return sElement;
            }











































            public static string getLatestfilename(string path)
            {
                try
                {
                    var directory = new DirectoryInfo(path);
                    var myFile = (from f in directory.GetFiles()
                                  orderby f.LastWriteTime descending
                                  select f).First().ToString();
                    return myFile;
                }
                catch
                {
                    return null;
                }
            }



            public void OCR_PDF_Generate(IWebDriver _driver, string PDFType)
            {
                switch (PDFType)
                {
                    case "PRE_Contract_CreditInformation":
                        Thread.Sleep(5000);
                        _driver.FindElement(By.Id("postContractCreditInformationPDF")).Click();
                        Thread.Sleep(5000);
                        break;
                    case "Adequate_Explanation":
                        Thread.Sleep(5000);
                        _driver.FindElement(By.Id("postAdequateExplanationPDF")).Click();
                        Thread.Sleep(5000);
                        break;
                    case "General_credit_card_conditions":
                        Thread.Sleep(5000);
                        _driver.FindElement(By.Id("creditCardConditionsPDF")).Click();
                        Thread.Sleep(5000);
                        break;
                    case "Credit_card_agreement":
                        Thread.Sleep(5000);
                        _driver.FindElement(By.Id("creditCardAgreementPDF")).Click();
                        Thread.Sleep(5000);
                        break;

                }

            }


            public string ReadPDF(string PDFType, string Page, string Xpath)

            {

                string PDFLoation = ScenarioContext.Current["PDFpath"].ToString();
                string newfilename = getLatestfilename(PDFLoation);
                string sValue = null;
                if (newfilename.Contains("_"))
                {
                    string TrimNewfile = newfilename.Substring(newfilename.LastIndexOf('_') + 1).ToString();
                    var directory = new DirectoryInfo(PDFLoation);
                    string concatSource = PDFLoation + "\\" + TrimNewfile;
                    string concatDes = PDFLoation + "\\" + PDFType + "_" + TrimNewfile;
                    Thread.Sleep(3000);
                    File.Move(concatSource, concatDes);
                    string dataloc = PDFLoation + "Page -{0}.{1}";
                    Ocr.ConvertorFromPdFtoData(concatDes, dataloc);
                    //return dataloc;
                }
                else
                {
                    string concatSource = PDFLoation + "\\" + newfilename;
                    string concatDes = PDFLoation + "\\" + PDFType + "_" + newfilename;
                    Thread.Sleep(3000);
                    File.Move(concatSource, concatDes);
                    string dataloc = PDFLoation + "\\" + "Page -{0}.{1}";
                    Ocr.ConvertorFromPdFtoData(concatDes, dataloc);
                    //return dataloc;
                }

                if (Xpath.StartsWith("Line") || Xpath.StartsWith("word"))
                {
                    string hhh = Ocr.ReadImage(PDFLoation, PDFType, Page, Xpath);
                    sValue = hhh;
                }

                else
                {
                    string hhh = Ocr.ReadImageParagraph(PDFLoation, PDFType, Page, Xpath);
                    sValue = hhh;
                }


                return sValue;

            }


            public void OCR_PDF_Extract(string PDFType)

            {

                //string PDFLoation = ScenarioContext.Current["PDFpath"].ToString();
                string PDFLoation = "C:\\demo";
                string newfilename = getLatestfilename(PDFLoation);
                //string sValue = null;
                if (newfilename.Contains("_"))
                {
                    string TrimNewfile = newfilename.Substring(newfilename.LastIndexOf('_') + 1).ToString();
                    var directory = new DirectoryInfo(PDFLoation);
                    string concatSource = PDFLoation + "\\" + TrimNewfile;
                    string concatDes = PDFLoation + "\\" + PDFType + "_" + TrimNewfile;
                    Thread.Sleep(3000);
                    //File.Move(concatSource, concatDes);
                    //string dataloc = PDFLoation + PDFType + "__Page -{0}.{1}";
                    string adda = "\\SPLIT";
                    string dataloc = PDFLoation + adda + "__Page -{0}.{1}";
                    Ocr.ConvertorFromPdFtoData(concatDes, dataloc);
                    //return dataloc;
                }
                else
                {
                    string concatSource = PDFLoation + "\\" + newfilename;
                    string concatDes = PDFLoation + "\\" + PDFType + "_" + newfilename;
                    Thread.Sleep(3000);
                    //File.Move(concatSource, concatDes);
                    string dataloc = PDFLoation + "\\" + PDFType + "__Page -{0}.{1}";
                    Ocr.ConvertorFromPdFtoData(concatDes, dataloc);
                    //return dataloc;
                }


            }

            public static string OCR_PDF_READ(string PDFType, string Page, string Xpath)

            {

                //string PDFLoation = ScenarioContext.Current["PDFpath"].ToString();
                string PDFLoation = "C:\\demo";
                string newfilename = getLatestfilename(PDFLoation);
                string sValue = null;


                if (Xpath.StartsWith("line") || Xpath.StartsWith("word"))
                {
                    string hhh = Ocr.ReadImage(PDFLoation, PDFType, Page, Xpath);

                    sValue = Regex.Replace(hhh, @"(°)", "").Replace(" ", string.Empty);
                }

                else
                {
                    string hhh = Ocr.ReadImageParagraph(PDFLoation, PDFType, Page, Xpath);
                    sValue = Regex.Replace(hhh, @"(°)", "").Replace(" ", string.Empty);
                }
                return sValue;
            }

            public static string OCR_PDF_READ_NO_Space(string PDFType, string Page, string Xpath)

            {

                //string PDFLoation = ScenarioContext.Current["PDFpath"].ToString();
                string PDFLoation = "C:\\demo";
                string newfilename = getLatestfilename(PDFLoation);
                string sValue = null;


                if (Xpath.StartsWith("line") || Xpath.StartsWith("word"))
                {
                    string hhh = Ocr.ReadImage(PDFLoation, PDFType, Page, Xpath);

                    sValue = Regex.Replace(hhh, @"(°)", "");
                }

                else
                {
                    string hhh = Ocr.ReadImageParagraph(PDFLoation, PDFType, Page, Xpath);
                    sValue = Regex.Replace(hhh, @"(°)", "");
                }
                return sValue;
            }


            public static void OCR_PDF_Validate(string actualText, string expectedText)
            {

                string expectedtrim = (expectedText.Replace(" ", string.Empty)).ToLower();
                string actualTexta = (actualText.Replace(" ", string.Empty)).ToLower();
                if (actualTexta.Contains(expectedtrim))
                    Console.WriteLine("PDF Comparion Passed.. -- EXPECTED----" + expectedText + "ACTUAL----" + actualTexta);
                else
                {
                    Console.WriteLine("Failed - PDF value not matched with expected .. -- ACTUAL---- " + actualText + "EXPECTED----" + expectedText);
                    //Assert.Fail("PDF doesn't contain the text...  " + expectedText);
                }
            }


            public static void OCR_PDF_ValidateContent(string actualText, string expectedText, string content)
            {

                string expectedtrim = (expectedText.Replace(" ", string.Empty)).ToLower();
                string actualTexta = (actualText.Replace(" ", string.Empty)).ToLower();
                if (actualTexta.Contains(expectedtrim))
                {
                    Console.BackgroundColor = ConsoleColor.Blue;
                    Console.WriteLine("PDF Comparion" + "{***" + content + "***}" + "has Passed -- with Actual value.... " + actualText);
                }
                else
                {
                    Console.WriteLine("PDF contains the Actual content as.. -- " + actualText);
                }
                Console.ResetColor();
            }


            public static string CellvaluePDFStaticData(string PDFType, string Item)
            {
                XSSFWorkbook hssfwb;
                string test = null;
                string sValue = null;
                string excelPath = null;
                //var featureTitle = FeatureContext.Current.FeatureInfo.Title.Split('-')[0];
                //var scenarioName = ScenarioContext.Current.ScenarioInfo.Title;
                //var testdataName = FeatureContext.Current["TestdataName"];

                excelPath = SolutionPath + "/PDF/" + "PDF_StaticData" + ".xlsx";

                using (FileStream file = new FileStream(@excelPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    hssfwb = new XSSFWorkbook(file);
                }
                ISheet sheet = hssfwb.GetSheet(PDFType);
                var range = "A1:EE1000";
                var cellRange = NPOI.SS.Util.CellRangeAddress.ValueOf(range);

                for (var a = cellRange.FirstRow; a <= sheet.LastRowNum; a++)
                {
                    if (sheet.GetRow(a).GetCell(0).StringCellValue.Equals(Item)) //&& sheet.GetRow(a).GetCell(1).StringCellValue.Equals(Iteration))
                    {
                        for (var j = cellRange.FirstRow; j <= cellRange.LastRow; j++)
                        {
                            var row = sheet.GetRow(j);
                            test = row.GetCell(0).StringCellValue;

                            if (test.Equals(Item))
                            {
                                sValue = row.GetCell(1).StringCellValue;//sheet.GetRow(a).GetCell(j+1).StringCellValue;
                                break;
                            }
                        }
                        break;
                    }
                }
                return sValue;
            }

            public static void PDF_CCA_Validate_StaticData()

            {
                string ChargesPercentage = OCR_PDF_READ("Credit_card_agreement", "4", PDFContentPost.ChargesPercentage);
                string testA = CellvaluePDFStaticData("CCA", "ChargesPercentage");
                OCR_PDF_Validate(ChargesPercentage, testA);

                string ChargesTransactionFees = OCR_PDF_READ("Credit_card_agreement", "4", PDFContentPost.ChargesTransactionFees);
                string testB = CellvaluePDFStaticData("CCA", "ChargesTransactionFees");
                OCR_PDF_Validate(ChargesTransactionFees, testB);

                string LatePaymentFees = OCR_PDF_READ("Credit_card_agreement", "4", PDFContentPost.LatePaymentFees);
                string testC = CellvaluePDFStaticData("CCA", "LatePaymentFees");
                OCR_PDF_Validate(LatePaymentFees, testC);


                string OverlimitFees = OCR_PDF_READ("Credit_card_agreement", "4", PDFContentPost.OverlimitFees);
                string testD = CellvaluePDFStaticData("CCA", "OverlimitFees");
                OCR_PDF_Validate(OverlimitFees, testD);

                string ReturnedPaymentFees = OCR_PDF_READ("Credit_card_agreement", "4", PDFContentPost.ReturnedPaymentFees);
                string testF = CellvaluePDFStaticData("CCA", "ReturnedPaymentFees");
                OCR_PDF_Validate(ReturnedPaymentFees, testF);


                string TraceFees = OCR_PDF_READ("Credit_card_agreement", "4", PDFContentPost.TraceFees);
                string testG = CellvaluePDFStaticData("CCA", "TraceFees");
                OCR_PDF_Validate(TraceFees, testG);


                string TransactionFeeForiegn = OCR_PDF_READ("Credit_card_agreement", "4", PDFContentPost.TransactionFeeForiegn);
                string testH = CellvaluePDFStaticData("CCA", "TransactionFeeForiegn");
                OCR_PDF_Validate(TransactionFeeForiegn, testH);

                string StatementCopy = OCR_PDF_READ("Credit_card_agreement", "4", PDFContentPost.StatementCopy);
                string testI = CellvaluePDFStaticData("CCA", "StatementCopy");
                OCR_PDF_Validate(StatementCopy, testI);

                string RightofWithdrawals = OCR_PDF_READ("Credit_card_agreement", "5", PDFContentPost.RightofWithdrawals);
                string testJ = CellvaluePDFStaticData("CCA", "RightofWithdrawals");
                OCR_PDF_Validate(RightofWithdrawals, testJ);

                string WithDrawalsPhone = OCR_PDF_READ("Credit_card_agreement", "5", PDFContentPost.WithDrawalsPhone);
                string testK = CellvaluePDFStaticData("CCA", "WithDrawalsPhone");
                OCR_PDF_Validate(WithDrawalsPhone, testK);

                string ComplaintsPhone = OCR_PDF_READ("Credit_card_agreement", "5", PDFContentPost.ComplaintsPhone);
                string testL = CellvaluePDFStaticData("CCA", "ComplaintsPhone");
                OCR_PDF_Validate(ComplaintsPhone, testL);

                string ComplaintsEmail = OCR_PDF_READ("Credit_card_agreement", "5", PDFContentPost.ComplaintsEmail);
                string testM = CellvaluePDFStaticData("CCA", "ComplaintsEmail");
                OCR_PDF_Validate(ComplaintsEmail, testM);

                string CompanyRegistrationNumber = OCR_PDF_READ("Credit_card_agreement", "14", PDFContentPost.CompanyRegistrationNumber);
                string testN = CellvaluePDFStaticData("CCA", "CompanyRegistrationNumber");
                OCR_PDF_Validate(CompanyRegistrationNumber, testN);

                string Headera = OCR_PDF_READ("Credit_card_agreement", "1", "par_2_1");
                string testa = CellvaluePDFStaticData("CCA", "Header");
                OCR_PDF_Validate(Headera, testa);

                string PARTIES = OCR_PDF_READ("Credit_card_agreement", "1", "line_2_5");
                string testb = CellvaluePDFStaticData("CCA", "PARTIES");
                OCR_PDF_Validate(PARTIES, testb);

                string Creditlimit = OCR_PDF_READ("Credit_card_agreement", "1", "par_2_6");
                string CreditlimitA = CellvaluePDFStaticData("CCA", "Creditlimit");
                OCR_PDF_Validate(Creditlimit, CreditlimitA);


                string Yourpayments = OCR_PDF_READ("Credit_card_agreement", "2", "par_2_9");
                string YourpaymentsA = CellvaluePDFStaticData("CCA", "Yourpayments");
                OCR_PDF_Validate(Yourpayments, YourpaymentsA);

                string APR = OCR_PDF_READ("Credit_card_agreement", "4", "par_2_7");
                string APRA = CellvaluePDFStaticData("CCA", "APR");
                OCR_PDF_Validate(APR, APRA);

                string Charges = OCR_PDF_READ("Credit_card_agreement", "4", "par_2_9");
                string ChargesA = CellvaluePDFStaticData("CCA", "Charges");
                OCR_PDF_Validate(Charges, ChargesA);

                string withdrawal = OCR_PDF_READ("Credit_card_agreement", "5", "par_2_3");
                string withdrawalA = CellvaluePDFStaticData("CCA", "withdrawal");
                OCR_PDF_Validate(withdrawal, withdrawalA);

                string payments = OCR_PDF_READ("Credit_card_agreement", "5", "par_2_6");
                string paymentsA = CellvaluePDFStaticData("CCA", "payments");
                OCR_PDF_Validate(payments, paymentsA);

                string end = OCR_PDF_READ("Credit_card_agreement", "5", "par_2_9");
                string endA = CellvaluePDFStaticData("CCA", "end");
                OCR_PDF_Validate(end, endA);


                string keyinformation = OCR_PDF_READ("Credit_card_agreement", "5", "par_2_12");
                string keyinformationA = CellvaluePDFStaticData("CCA", "Keyinformation");
                OCR_PDF_Validate(keyinformation, keyinformationA);

                string Complaints = OCR_PDF_READ("Credit_card_agreement", "5", "par_2_17");
                string ComplaintsA = CellvaluePDFStaticData("CCA", "Complaints");
                OCR_PDF_Validate(Complaints, ComplaintsA);

                string special = OCR_PDF_READ("Credit_card_agreement", "6", "par_2_4");
                string specialA = CellvaluePDFStaticData("CCA", "special");
                OCR_PDF_Validate(special, specialA);

                string Using = OCR_PDF_READ("Credit_card_agreement", "6", "par_2_1");
                string UsingA = CellvaluePDFStaticData("CCA", "Using");
                OCR_PDF_Validate(Using, UsingA);

                string stoppingpayments = OCR_PDF_READ("Credit_card_agreement", "7", "par_2_1");
                string stoppingpaymentsA = CellvaluePDFStaticData("CCA", "stoppingpayments");
                OCR_PDF_Validate(stoppingpayments, stoppingpaymentsA);

                string Refunds = OCR_PDF_READ("Credit_card_agreement", "10", "par_2_1");
                string RefundsA = CellvaluePDFStaticData("CCA", "Refunds");
                OCR_PDF_Validate(Refunds, RefundsA);


                string security = OCR_PDF_READ("Credit_card_agreement", "10", "par_2_1");
                string securityA = CellvaluePDFStaticData("CCA", "security");
                OCR_PDF_Validate(security, securityA);

                string stolen = OCR_PDF_READ("Credit_card_agreement", "10", "par_2_1");
                string stolenA = CellvaluePDFStaticData("CCA", "stolen");
                OCR_PDF_Validate(stolen, stolenA);

                string Ending = OCR_PDF_READ("Credit_card_agreement", "11", "par_2_1");
                string EndingA = CellvaluePDFStaticData("CCA", "Ending");
                OCR_PDF_Validate(Ending, EndingA);



            }






























        }

        }
}
