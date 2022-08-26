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
using Word = Microsoft.Office.Interop.Word;
using Spire.Pdf;
using BarcodeLib;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


/*TODO
 * redo layout
 * pick a color pallete 
 * design
 * make new directory with QRR and labels .docs.
 */




namespace AutomateCoA
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        Font myFont = new Font("Arial", 20, FontStyle.Bold);

        //cant think of a better way to scope
        private string qrrPDFPath;
        private string coaPDFPath;

        //makes barcode generator work somehow.  DONT TOUCH
        public bool ThumbnailCallback()
        {
            return false;
        }



        //QRR/COA Tab///////////////////////////////////////////////////////////////////////////////////////////////

        //open and edit QRR
        private void FindQRR()
        {
            OpenFileDialog _fileDialog = new OpenFileDialog();
            _fileDialog.InitialDirectory = "Z:\\SJShare\\SJCOMMON\\DI\\MIC\\COMPONENT DOCS & VALIDATIONS\\Quarantine Release\\Digital QRR - Test Set";

            if (_fileDialog.ShowDialog() == DialogResult.OK)
            {
                //launch and open an instance of word
                Word.Application wordApp = new Word.Application();
                object fileName = _fileDialog.FileName;
                object readOnly = false;
                object isVisible = true;
                //easy way to handle random refs I dont care about
                object iDontCare = System.Reflection.Missing.Value;

                wordApp.Visible = true;

                Word.Document qrrDoc = wordApp.Documents.Open(ref fileName, ref iDontCare, ref readOnly,
                                                                                        ref iDontCare, ref iDontCare, ref iDontCare,
                                                                                        ref iDontCare, ref iDontCare, ref iDontCare,
                                                                                        ref iDontCare, ref iDontCare, ref isVisible,
                                                                                        ref iDontCare, ref iDontCare, ref iDontCare, ref iDontCare);
                //opens doc with track changes turned off
                qrrDoc.TrackRevisions = false;
                qrrDoc.Activate();
            }
        }

        //find and save CoA
        private void FindCoA()
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



            //go find the stuff from vendors
            try
            {
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
                        //lot number and execture search
                        mydriver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div/div/div[2]/div/div[1]/form/div/div[3]/div/div/div/input")).SendKeys(LotBox1.Text);
                        mydriver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div/div/div[2]/div/div[1]/form/div/div[3]/div/div/div/input")).SendKeys(OpenQA.Selenium.Keys.Return);

                        //TODO why cant I close the tab??!?! without closing all the things
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
            catch (Exception Exception)
            {
                int errorLength = 100;

                var myError = Exception;
                mydriver.Quit();
                MessageBox.Show("Tell Pete this happened: " + myError.ToString().Substring(0, errorLength));
            }
        }

        //find QRR PDF
        private void QRRPDF()
        {
            OpenFileDialog _fileDialog = new OpenFileDialog();
            _fileDialog.ShowDialog();

            qrrPDFPath = _fileDialog.FileName.ToString();

            //shows end of file name for easier verification
            QRRFileNameBx.Text = qrrPDFPath;
            QRRFileNameBx.SelectionStart = QRRFileNameBx.Text.Length;
            QRRFileNameBx.SelectionLength = 0;
        }

        //find CoAPDF
        private void CoAPDF()
        {
            OpenFileDialog _fileDialog = new OpenFileDialog();
            _fileDialog.ShowDialog();

            coaPDFPath = _fileDialog.FileName.ToString();

            CoAFileNameBx.Text = coaPDFPath;
            CoAFileNameBx.SelectionStart = CoAFileNameBx.Text.Length;
            CoAFileNameBx.SelectionLength = 0;
        }

        //combine QRR and COA then save to desktop
        private void SaveFinalPDF()
        {
            try
            {
                //store paths in string array
                string[] pdfPaths = { qrrPDFPath, coaPDFPath };

                //load all pdfs into an PdfDocument obj and store all objs in PdfDocument array
                PdfDocument[] pdfs = new PdfDocument[pdfPaths.Length];

                for (int i = 0; i < pdfPaths.Length; i++)
                {
                    pdfs[i] = new PdfDocument(pdfPaths[i]);
                }

                PdfDocument finalPdf = new PdfDocument();

                //just first page of QRR and all pages of CoA
                finalPdf.InsertPage(pdfs[0], 0);
                finalPdf.AppendPage(pdfs[1]);

                string finalPdfPathDesktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string finalPdfPathFiixFolder = "Z:\\SJShare\\SJCOMMON\\DI\\MIC\\COMPONENT DOCS & VALIDATIONS\\Fiix Scanned Documents";
                string finalPdfName = FinalPDFBx.Text;

                finalPdf.SaveToFile(finalPdfPathFiixFolder + "\\" + finalPdfName + ".pdf");
                finalPdf.SaveToFile(finalPdfPathDesktop + "\\" + finalPdfName + ".pdf");
                Process.Start(finalPdfPathDesktop + "\\" + finalPdfName + ".pdf");
            }

            catch (Exception)
            {
                MessageBox.Show("Please select both .pdfs");
            }

        }

        //clear controls
        void ClearControls()
        {
            foreach (Control control in PTSTAB.SelectedTab.Controls)
            {
                if (control is System.Windows.Forms.TextBox)
                {
                    //cast to textbox to use clear method
                    ((System.Windows.Forms.TextBox)control).Clear();

                }

                if (control is System.Windows.Forms.ComboBox)
                {
                    ((System.Windows.Forms.ComboBox)control).SelectedIndex = -1;
                }

                if (control is PictureBox)
                {
                    ((PictureBox)control).Image = null;
                }
            }

        }

        //interface events
        private void QRRFindBtn_Click_1(object sender, EventArgs e)
        {
            FindQRR();
        }

        private void FetchBtn1_Click_1(object sender, EventArgs e)
        {
            FindCoA();

        }

        private void QRRPDFBtn_Click(object sender, EventArgs e)
        {
            QRRPDF();
        }

        private void CoAPDFBtn_Click(object sender, EventArgs e)
        {
            CoAPDF();
        }

        private void SaveFinalPDFBtn_Click(object sender, EventArgs e)
        {
            SaveFinalPDF();
        }

        private void clrCoABtn_Click(object sender, EventArgs e)
        {
            ClearControls();
        }


        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



        //Barcode Tab

        //makes barcodes
        private void GenerateCodes()
        {


            //link textbox control to picturebox controls...felt good
            Dictionary<Control, Control> barDict = new Dictionary<Control, Control>();

            barDict.Add(lot1, lot1code);
            barDict.Add(bud1, bud1code);
            barDict.Add(lot2, lot2code);
            barDict.Add(bud2, bud2code);
            barDict.Add(lot3, lot3code);
            barDict.Add(bud3, bud3code);
            barDict.Add(lot4, lot4code);
            barDict.Add(bud4, bud4code);
            barDict.Add(lot5, lot5code);
            barDict.Add(bud5, bud5code);
            barDict.Add(lot6, lot6code);
            barDict.Add(bud6, bud6code);


            foreach (var barText in barDict.Keys)
            {
                if (barText.Text != "")
                {

                    BarcodeLib.Barcode barcode = new BarcodeLib.Barcode()
                    {
                        IncludeLabel = true,
                        LabelFont = myFont,
                        Alignment = AlignmentPositions.CENTER,
                        Width = 280,
                        Height = 70,
                        //RotateFlipType = RotateFlipType.RotateNoneFlipNone,
                        BackColor = Color.White,
                        ForeColor = Color.Black,
                    };

                    Image img = barcode.Encode(TYPE.CODE128B, barText.Text);
                    Image.GetThumbnailImageAbort myCallback = new Image.GetThumbnailImageAbort(ThumbnailCallback);

                    PictureBox pb = barDict[barText] as PictureBox;
                    pb.Image = img;
                }
            }
        }


        //interface events
        private void MakeBarCodeBtn_Click_1(object sender, EventArgs e)
        {
            GenerateCodes();
        }

        private void clearButton_Click_1(object sender, EventArgs e)
        {
            ClearControls();
        }

        //right click menu
        private void copyToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Clipboard.Clear();

            ToolStripItem menuItem = sender as ToolStripItem;

            if (menuItem != null)
            {


                ContextMenuStrip owner = menuItem.Owner as ContextMenuStrip;
                if (owner != null)
                {
                    // Get the control that is displaying this context menu
                    Control sourceControl = owner.SourceControl;
                    PictureBox pb = sourceControl as PictureBox;
                    Image img = pb.Image;
                    if (img != null)
                    {
                        Clipboard.SetImage(img);
                    }
                }
            }
        }



        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



        //PTS Tab
        private void GeneratePTSCodes()
        {

            Dictionary<Control, Control> ptsDict = new Dictionary<Control, Control>();

            ptsDict.Add(PTS_initial, PTS_initialCode);
            ptsDict.Add(PTS_lot, PTS_lotCode);
            ptsDict.Add(PTS_sampleName, PTS_sampleNameCode);
            ptsDict.Add(PTS_sampleLot, PTS_sampleLotCode);
            ptsDict.Add(PTS_dil, PTS_dilCode);

            foreach (var barText in ptsDict.Keys)
            {
                if (barText.Text != "")
                {

                    BarcodeLib.Barcode barcode = new BarcodeLib.Barcode()
                    {
                        IncludeLabel = true,
                        LabelFont = myFont,
                        Alignment = AlignmentPositions.CENTER,
                        Width = 280,
                        Height = 70,
                        //RotateFlipType = RotateFlipType.RotateNoneFlipNone,
                        BackColor = Color.White,
                        ForeColor = Color.Black,
                    };

                    Image img = barcode.Encode(TYPE.CODE128B, barText.Text);
                    Image.GetThumbnailImageAbort myCallback = new Image.GetThumbnailImageAbort(ThumbnailCallback);

                    PictureBox pb = ptsDict[barText] as PictureBox;
                    pb.Image = img;
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            GeneratePTSCodes();
        }

        private void ClearPTS_Click(object sender, EventArgs e)
        {
            ClearControls();
        }
    }
}
