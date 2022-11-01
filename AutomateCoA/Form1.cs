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
using System.Runtime.Remoting.Channels;
using System.Runtime.CompilerServices;
using Microsoft.Office.Core;
//using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace AutomateCoA
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            PTSTAB.DrawItem += new DrawItemEventHandler(PTSTAB_DrawItem);
        }


        Font myFont = new Font("Arial", 20, FontStyle.Bold);

        //cant think of a better way to scope
        private string qrrPDFPath;
        private string coaPDFPath;
        private bool bcrMode = false;
        

        //makes barcode generator work somehow.  DONT TOUCH
        public bool ThumbnailCallback()
        {
            return false;
        }

        private void PTSTAB_DrawItem(Object sender, System.Windows.Forms.DrawItemEventArgs e)
        {

            //System.Drawing


            Graphics g = e.Graphics;
            Brush  _textBrush;
            Brush _activetabBrush;
            Brush _tabBrush;                   
;
            //Get the item from the collection
            TabPage _tabPage = PTSTAB.TabPages[e.Index];

            //Get the real bounds for the tab rectangle
            // Get the real bounds for the tab rectangle.
            Rectangle _tabBounds = PTSTAB.GetTabRect(e.Index);

            if (e.State == DrawItemState.Selected)
            {

                // Draw a different background color, and don't paint a focus rectangle.
                _textBrush = new SolidBrush(Color.FromArgb(245,245,245));
                _activetabBrush = new SolidBrush(Color.FromArgb(222, 143, 110));
                g.FillRectangle(_activetabBrush, e.Bounds);
            }
            
            else
            {
                _tabBrush = new SolidBrush(Color.FromArgb(55,73,94));
                _textBrush = new SolidBrush(Color.FromArgb(245, 245, 245));
                g.FillRectangle(_tabBrush, e.Bounds);
            }
            

            // Use our own font.
            Font _tabFont = new Font("Arial", 10.0f, FontStyle.Bold, GraphicsUnit.Pixel);

            // Draw string. Center the text.
            StringFormat _stringFlags = new StringFormat();
            _stringFlags.Alignment = StringAlignment.Center;
            _stringFlags.LineAlignment = StringAlignment.Center;
            g.DrawString(_tabPage.Text, _tabFont, _textBrush, _tabBounds, new StringFormat(_stringFlags));
        }

        //QRR/COA Tab///////////////////////////////////////////////////////////////////////////////////////////////

        //open and edit QRR
        private void FindQRR()
        {
            
            if(bcrMode == false)
            {
                OpenFileDialog _fileDialog = new OpenFileDialog();
                _fileDialog.InitialDirectory = "Z:\\SJShare\\SJCOMMON\\DI\\MIC\\COMPONENT DOCS & VALIDATIONS\\Quarantine Release\\Digital QRR - Lab Buddy";

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
            if(bcrMode == true)
            {
                OpenFileDialog _fileDialog = new OpenFileDialog();
                _fileDialog.InitialDirectory = "Z:\\SJShare\\SJCOMMON\\DI\\MIC\\COMPONENT DOCS & VALIDATIONS\\Quarantine Release\\Digital BCR - Lab Buddy";

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
        }

        //find and save CoA
        private void FindCoA()
        {
            
            try
            {
                //iniate web browser with console hidden.  Browser will close when driver is exited.
                ChromeOptions options = new ChromeOptions();
                options.LeaveBrowserRunning = false;
                ChromeDriverService service = ChromeDriverService.CreateDefaultService();
                new DriverManager().SetUpDriver(new ChromeConfig());


                service.HideCommandPromptWindow = true;
                IWebDriver myDriver = new ChromeDriver(service, options);

                //when waiting on elements to load
                WebDriverWait wait = new WebDriverWait(myDriver, TimeSpan.FromSeconds(10));

                //get stuff from vendors

                //get input from text/comboboxes

                switch (VendorBox1.Text.ToUpper())
                {
                    case "FISHER":


                        myDriver.Url = "https://www.fishersci.com/us/en/catalog/search/certificates.html";
                        //catalog number
                        myDriver.FindElement(By.XPath("/html/body/section/section/div/div/div/div[2]/div[1]/div/form/ul/li[1]/input")).SendKeys(ItemBox1.Text);
                        //lot number
                        myDriver.FindElement(By.XPath("/html/body/section/section/div/div/div/div[2]/div[1]/div/form/ul/li[2]/input")).SendKeys(LotBox1.Text);
                        //document type.  HAVE TO HIT RETURN TO CONFIRM SELECTION OR IT ALL GOES TO HELL!!!!
                        myDriver.FindElement(By.XPath("/html/body/section/section/div/div/div/div[2]/div[1]/div/form/ul/li[3]/select")).SendKeys("C");
                        myDriver.FindElement(By.XPath("/html/body/section/section/div/div/div/div[2]/div[1]/div/form/ul/li[3]/select")).SendKeys(OpenQA.Selenium.Keys.Return);
                        //search button
                        myDriver.FindElement(By.CssSelector("#certificate-search-form > form > ul > li:nth-child(4) > input")).Click();
                        //certificate inside of table     
                        myDriver.FindElement(By.XPath("/html/body/section/section/div/div/div/div[3]/table/tbody/tr/td[6]/a")).Click();

                        //close window but not exit webdriver
                        myDriver.Close();
                        break;

                    case "SIGMA":
                        myDriver.Url = "https://www.sigmaaldrich.com/US/en/search";

                        //click on cookie accept button
                        IWebElement cookiebtn = wait.Until(drvr => drvr.FindElement(By.CssSelector("#onetrust-close-btn-container > button")));
                        cookiebtn.Click();

                        //product number
                        myDriver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div/div/div[2]/div/div[1]/form/div/div[1]/div/div/div/input")).SendKeys(ItemBox1.Text);
                        //lot number and execture search
                        myDriver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div/div/div[2]/div/div[1]/form/div/div[3]/div/div/div/input")).SendKeys(LotBox1.Text);
                        myDriver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div/div/div[2]/div/div[1]/form/div/div[3]/div/div/div/input")).SendKeys(OpenQA.Selenium.Keys.Return);

                        //TODO why cant I close the tab??!?! without closing all the things
                        //myDriver.Close();

                        break;

                    case "VWR":
                        myDriver.Url = "https://us.vwr.com/store/search/searchCerts.jsp?tabId=certSearch";

                        //part number

                        myDriver.FindElement(By.XPath("//*[@id='certSearchPartNumber']")).SendKeys(ItemBox1.Text);
                        //lot number
                        myDriver.FindElement(By.XPath("//*[@id='certSearchLotNumber']")).SendKeys(LotBox1.Text);
                        //search button
                        myDriver.FindElement(By.XPath("//*[@id='certSearch']")).Click();
                        //certificate inside of table
                        myDriver.FindElement(By.XPath("//*[@id='content']/div/div[2]/div[2]/div/div/table/tbody[1]/tr/td[2]/a")).Click();

                        //myDriver.Close();

                        break;

                    case "BD":
                        myDriver.Url = "https://regdocs.bd.com/regdocs/qcinfo";


                        myDriver.FindElement(By.XPath("//*[@id='page']/div/section[2]/div/div/div/div/div/form/div[2]/div[1]/div/div[1]/input")).SendKeys(ItemBox1.Text);
                        //lot number
                        myDriver.FindElement(By.XPath("//*[@id='page']/div/section[2]/div/div/div/div/div/form/div[2]/div[1]/div/div[2]/input")).SendKeys(LotBox1.Text);
                        //search button
                        myDriver.FindElement(By.XPath("//*[@id='page']/div/section[2]/div/div/div/div/div/form/button[2]")).Click();
                        //certificate inside of table

                        //click on cookie accept button
                        IWebElement certbtn = wait.Until(drvr => drvr.FindElement(By.XPath("//*[@id='page']/div/section[1]/div/div[2]/div[1]/div/div[2]/table/tbody/tr[2]/td[5]/i")));
                        certbtn.Click();

                        break;

                    case "MILLIPORE":
                        myDriver.Url = "https://www.emdmillipore.com/US/en/documents/Z.qb.qB.tecAAAFDDJUsznLq,nav#s_mbqBTWQAAAFKyVsIik_d";
                        //defeats automation blocking. BOOYAH!
                        myDriver.Manage().Cookies.DeleteAllCookies();

                        //catalog number  
                        myDriver.FindElement(By.XPath("//*[@id='COAOrderNumber']")).SendKeys(ItemBox1.Text);
                        //lot number  
                        myDriver.FindElement(By.XPath("//*[@id='COABatchNumber']")).SendKeys(LotBox1.Text);
                        //myDriver.FindElement(By.XPath("//*[@id='COABatchNumber']")).SendKeys(OpenQA.Selenium.Keys.Return);
                        //search button
                        myDriver.FindElement(By.XPath("//*[@id='s_mbqBTWQAAAFKyVsIik_d_find']")).SendKeys(OpenQA.Selenium.Keys.Return);

                        //certificate inside of table
                        //this will break  results is tied to individual coa.  leaving window open for user to manually click and download
                        //certbtn = wait.Until(drvr => drvr.FindElement(By.CssSelector("#result-FvCbqBDaYAAAFErYZWba4i > table > tbody > tr > td:nth-child(1) > a")));
                        //certbtn.Click();

                        break;
                }
            }
            catch (InvalidOperationException e)
            {
                MessageBox.Show(e.Message);
                
                if (e.Message.ToUpper().Contains("VERSION OF CHROME"))
                {
                    MessageBox.Show("Please Update To The Most Current Version of Chrome.");
                }               
            }
        }

        //find QRR PDF
        private void QRRPDF()
        {
            OpenFileDialog _fileDialog = new OpenFileDialog();
            _fileDialog.Filter = "|*.pdf";
            _fileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
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
            string pdfLocation;

            if (bcrMode == true)
            {
                pdfLocation = "Z:\\SJShare\\SJCOMMON\\DI\\MIC\\Attachment Bucket";
            }
            else
            {
                pdfLocation = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
           
            OpenFileDialog _fileDialog = new OpenFileDialog();
            _fileDialog.Filter = "|*.pdf";
            _fileDialog.InitialDirectory = pdfLocation;
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
            PdfDocument firstPDF = new PdfDocument(qrrPDFPath);
            PdfDocument secondPDF = new PdfDocument(coaPDFPath);
            PdfDocument finalPdf = new PdfDocument();

            //gets rid of label page so only QRR/BCR info pages are saved
            int firstPDFLen = firstPDF.Pages.Count;

           
                //So this looks weird.  I think it has to do with the insertpage range array initializing at zero and being exclusive
                //so firstPDFLen has to be one less than the length you want??  Either way it works dont touch
                if (firstPDFLen == 1)
                {
                    firstPDFLen = firstPDFLen - 1;
                }
                else
                {
                    firstPDFLen = firstPDFLen - 2;
                }

                finalPdf.InsertPageRange(firstPDF, 0, firstPDFLen);
                finalPdf.AppendPage(secondPDF);

                string finalPdfPathDesktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string finalPdfPathFiixFolder = "Z:\\SJShare\\SJCOMMON\\DI\\MIC\\COMPONENT DOCS & VALIDATIONS\\Fiix Scanned Documents\\Added to FIIX";
                string finalPdfName = FinalPDFBx.Text;

                if(finalPdfName != "")
                {
                    finalPdf.SaveToFile(finalPdfPathFiixFolder + "\\" + finalPdfName + ".pdf");
                    finalPdf.SaveToFile(finalPdfPathDesktop + "\\" + finalPdfName + ".pdf");
                    Process.Start(finalPdfPathDesktop + "\\" + finalPdfName + ".pdf");
                }
                else
                {
                    MessageBox.Show("Please enter final pdf name.");
                }

               
            }

            catch (Exception)
            {
                MessageBox.Show("Please select two pdfs");
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

        private void Btn_Enter(object sender, EventArgs e)
        {
            Button activeButton = (Button)sender;
                        
            int width =activeButton.Width;
            int height = activeButton.Height;
            int scaleUp = 10;

            activeButton.Size = new System.Drawing.Size(width + scaleUp, height + scaleUp);
            activeButton.Font = new Font("", 13, FontStyle.Bold);
        }

        private void Btn_Leave(object sender, EventArgs e)
        {
            Button activeButton = (Button)sender;
            
            int width = activeButton.Width;
            int height = activeButton.Height;
            int scaleDown = -10;

            activeButton.Size = new System.Drawing.Size(width + scaleDown, height + scaleDown);
            activeButton.Font = new Font("", 11, FontStyle.Regular);            
        }

        private void QrrRadBtn_CheckedChanged(object sender, EventArgs e)
        {
            QRRFindBtn.Text = "Choose QRR";
            bcrMode = false;
            ItemBox1.Visible = true;
            LotBox1.Visible = true;
            VendorBox1.Visible = true;
            label1.Visible = true;
            label2.Visible = true;
            label3.Visible = true;
            FetchBtn1.Visible = true;
            label28.Visible = true;
            label29.Visible = true;
            QRRPDFBtn.Text = "Choose QRR PDF";
            CoAPDFBtn.Text = "Choose CoA PDF";
            label30.Text = "4)";
            label31.Text = "5)";
        }

        private void BcrRadBtn_CheckedChanged(object sender, EventArgs e)
        {
            QRRFindBtn.Text = "Choose BCR";
            bcrMode = true;
            ItemBox1.Visible = false;
            LotBox1.Visible = false;
            VendorBox1.Visible = false;
            label1.Visible = false;
            label2.Visible = false;
            label3.Visible = false;
            FetchBtn1.Visible = false;
            label28.Visible = false;
            label29.Visible = false;
            QRRPDFBtn.Text = "Choose BCR PDF";
            CoAPDFBtn.Text = "Attachment Bucket Data";
            label30.Text = "2)";
            label31.Text = "3)";
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
