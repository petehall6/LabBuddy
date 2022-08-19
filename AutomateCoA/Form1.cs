using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Remote;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using Microsoft.Office.Interop.Word;
using Spire.Pdf;



namespace AutomateCoA
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //open and edit QRR
        //TODO make new directory with QRR and labels .docs.
        //TODO set as default directory
        private void QRRFindBtn_Click(object sender, EventArgs e)
        {

            OpenFileDialog _fileDialog = new OpenFileDialog();
            _fileDialog.InitialDirectory = "Z:\\SJShare\\SJCOMMON\\DI\\MIC\\COMPONENT DOCS & VALIDATIONS\\Quarantine Release\\Digital QRR - Test Set";

            if (_fileDialog.ShowDialog() == DialogResult.OK)
            {
                //launch and open an instance of word
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                object fileName = _fileDialog.FileName;
                object readOnly = false;
                object isVisible = true;
                //easy way to handle random refs I dont care about
                object iDontCare = System.Reflection.Missing.Value;

                wordApp.Visible = true;

                Microsoft.Office.Interop.Word.Document qrrDoc = wordApp.Documents.Open(ref fileName, ref iDontCare, ref readOnly,
                                                                                        ref iDontCare, ref iDontCare, ref iDontCare,
                                                                                        ref iDontCare, ref iDontCare, ref iDontCare,
                                                                                        ref iDontCare, ref iDontCare, ref isVisible,
                                                                                        ref iDontCare, ref iDontCare, ref iDontCare, ref iDontCare);
                qrrDoc.Activate();
            }
        }


        //find and save CoA
        //TODO add other vendors
        private void FetchBtn1_Click(object sender, EventArgs e)
        {
            //iniate web browser with console hidden.  Browser will close when driver is exited.
            ChromeOptions options = new ChromeOptions();
            options.LeaveBrowserRunning = false;
            
            new DriverManager().SetUpDriver(new ChromeConfig());
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            IWebDriver mydriver = new ChromeDriver(service, options);

            //when waiting on elements to load
            WebDriverWait wait = new WebDriverWait(mydriver, TimeSpan.FromSeconds(10));

            //get input from text/comboboxes
            switch (VendorBox1.Text.ToUpper())
            {
                case "FISHER":
                    
                    mydriver.Url = "https://www.fishersci.com/us/en/catalog/search/certificates.html";
                    //catalog number
                    mydriver.FindElement(By.XPath("/html/body/section/section/div/div/div/div[2]/div[1]/div/form/ul/li[1]/input")).SendKeys(ItemBox1.Text);
                    //lot number
                    mydriver.FindElement(By.XPath("/html/body/section/section/div/div/div/div[2]/div[1]/div/form/ul/li[2]/input")).SendKeys(LotBox1.Text);
                    //document type.  HAVE TO HIT RETURN TO CONFIRM SELECTION OR IT ALL GOES TO HELL!!!!
                    mydriver.FindElement(By.XPath("/html/body/section/section/div/div/div/div[2]/div[1]/div/form/ul/li[3]/select")).SendKeys("C");
                    mydriver.FindElement(By.XPath("/html/body/section/section/div/div/div/div[2]/div[1]/div/form/ul/li[3]/select")).SendKeys(OpenQA.Selenium.Keys.Return);
                    //search button
                    mydriver.FindElement(By.CssSelector("#certificate-search-form > form > ul > li:nth-child(4) > input")).Click();
                    //certificate inside of table     
                    mydriver.FindElement(By.XPath("/html/body/section/section/div/div/div/div[3]/table/tbody/tr/td[6]/a")).Click();
                             

                    //close window but not exit webdriver
                    mydriver.Close();
                    break;

                case "SIGMA":
                    mydriver.Url = "https://www.sigmaaldrich.com/US/en/search";

                    //click on cookie accept button
                    IWebElement cookiebtn = wait.Until(drvr => drvr.FindElement(By.CssSelector("#onetrust-close-btn-container > button")));
                    cookiebtn.Click();
             
                    //product number
                    mydriver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div/div/div[2]/div/div[1]/form/div/div[1]/div/div/div/input")).SendKeys(ItemBox1.Text);
                    //lot number
                    mydriver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div/div/div[2]/div/div[1]/form/div/div[3]/div/div/div/input")).SendKeys(LotBox1.Text);
                    //search button
                    mydriver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div/div/div[2]/div/div[1]/form/div/button/span")).Click();

                    //TODO why cant I close the tab??!?! without closing all the thigns
                    //mydriver.Close();

                    break;

                case "VWR":
                    mydriver.Url = "https://us.vwr.com/store/search/searchCerts.jsp?tabId=certSearch";

                    //part number

                    mydriver.FindElement(By.XPath("//*[@id='certSearchPartNumber']")).SendKeys(ItemBox1.Text);
                    //lot number
                    mydriver.FindElement(By.XPath("//*[@id='certSearchLotNumber']")).SendKeys(LotBox1.Text);
                    //search button
                    mydriver.FindElement(By.XPath("//*[@id='certSearch']")).Click();
                    //certificate inside of table
                    mydriver.FindElement(By.XPath("//*[@id='content']/div/div[2]/div[2]/div/div/table/tbody[1]/tr/td[2]/a")).Click();              

                    //mydriver.Close();

                    break;

                case "BD":
                    mydriver.Url = "https://regdocs.bd.com/regdocs/qcinfo";


                    mydriver.FindElement(By.XPath("//*[@id='page']/div/section[2]/div/div/div/div/div/form/div[2]/div[1]/div/div[1]/input")).SendKeys(ItemBox1.Text);
                    //lot number
                    mydriver.FindElement(By.XPath("//*[@id='page']/div/section[2]/div/div/div/div/div/form/div[2]/div[1]/div/div[2]/input")).SendKeys(LotBox1.Text);
                    //search button
                    mydriver.FindElement(By.XPath("//*[@id='page']/div/section[2]/div/div/div/div/div/form/button[2]")).Click();
                    //certificate inside of table

                    //click on cookie accept button
                    IWebElement certbtn = wait.Until(drvr => drvr.FindElement(By.XPath("//*[@id='page']/div/section[1]/div/div[2]/div[1]/div/div[2]/table/tbody/tr[2]/td[5]/i")));
                    certbtn.Click();

                    break;

                case "MILLIPORE":
                    mydriver.Url = "https://www.emdmillipore.com/US/en/documents/Z.qb.qB.tecAAAFDDJUsznLq,nav#s_mbqBTWQAAAFKyVsIik_d";
                    //defeats automation blocking. BOOYAH!
                    mydriver.Manage().Cookies.DeleteAllCookies();

                    //catalog number  
                    mydriver.FindElement(By.XPath("//*[@id='COAOrderNumber']")).SendKeys(ItemBox1.Text);
                    //lot number  
                    mydriver.FindElement(By.XPath("//*[@id='COABatchNumber']")).SendKeys(LotBox1.Text);
                    //mydriver.FindElement(By.XPath("//*[@id='COABatchNumber']")).SendKeys(OpenQA.Selenium.Keys.Return);
                    //search button
                    mydriver.FindElement(By.XPath("//*[@id='s_mbqBTWQAAAFKyVsIik_d_find']")).SendKeys(OpenQA.Selenium.Keys.Return);
                    
                    //certificate inside of table
                    //this will break  results is tied to individual coa.  leaving window open for user to manually click and download
                    //certbtn = wait.Until(drvr => drvr.FindElement(By.CssSelector("#result-FvCbqBDaYAAAFErYZWba4i > table > tbody > tr > td:nth-child(1) > a")));
                    //certbtn.Click();

                    break;
            }
        }

        //clear CoA Buttons
        private void clrCoABtn_Click(object sender, EventArgs e)
        {
            ItemBox1.Text = "Item#";
            LotBox1.Text = "Lot#";
            VendorBox1.Text = "Vendor";
            QRRFileNameBx.Text = "";
            CoAFileNameBx.Text = "";
            FinalPDFBx.Text = "";

        }


        //cant think of a better way to scope
        private string qrrPDFPath;
        private string coaPDFPath;
        //merge and save QRR and CoA PDF
        private void QRRPDFBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog _fileDialog = new OpenFileDialog();
            _fileDialog.ShowDialog();

            qrrPDFPath = _fileDialog.FileName.ToString();

            QRRFileNameBx.Text = qrrPDFPath;

            
        }

        private void CoAPDFBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog _fileDialog = new OpenFileDialog();
            _fileDialog.ShowDialog();

            coaPDFPath = _fileDialog.FileName.ToString();

            CoAFileNameBx.Text = coaPDFPath;
        }

        //merge, rename and save final PDF
        private void SaveFinalPDFBtn_Click(object sender, EventArgs e)
        {
            //store paths in string array
            string[] pdfPaths = { qrrPDFPath, coaPDFPath };

            //load all pdfs into an PdfDocument obj and store all objs in PdfDocument array
            PdfDocument[] pdfs = new PdfDocument[pdfPaths.Length];

            for (int i=0; i< pdfPaths.Length; i++)
            {
                pdfs[i] = new PdfDocument(pdfPaths[i]);
            }

            PdfDocument finalPdf = new PdfDocument();

            //just first page of QRR and all pages of CoA
            finalPdf.InsertPage(pdfs[0], 0);
            finalPdf.AppendPage(pdfs[1]);

            string finalPdfPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            

            string finalPdfName = FinalPDFBx.Text;

            finalPdf.SaveToFile(finalPdfPath+"\\"+ finalPdfName + ".pdf");
            Process.Start(finalPdfPath + "\\" + finalPdfName +".pdf");





        }
    }
}
