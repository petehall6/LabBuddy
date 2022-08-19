namespace AutomateCoA
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.FetchBtn1 = new System.Windows.Forms.Button();
            this.ItemBox1 = new System.Windows.Forms.TextBox();
            this.LotBox1 = new System.Windows.Forms.TextBox();
            this.VendorBox1 = new System.Windows.Forms.ComboBox();
            this.clrCoABtn = new System.Windows.Forms.Button();
            this.QRRFindBtn = new System.Windows.Forms.Button();
            this.QRRLabel = new System.Windows.Forms.Label();
            this.CoALabel = new System.Windows.Forms.Label();
            this.ComboPDFSection = new System.Windows.Forms.Label();
            this.QRRFileNameBx = new System.Windows.Forms.TextBox();
            this.QRRPDFBtn = new System.Windows.Forms.Button();
            this.CoAPDFBtn = new System.Windows.Forms.Button();
            this.CoAFileNameBx = new System.Windows.Forms.TextBox();
            this.SaveFinalPDFBtn = new System.Windows.Forms.Button();
            this.PDFNameLabel = new System.Windows.Forms.Label();
            this.FinalPDFBx = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // FetchBtn1
            // 
            resources.ApplyResources(this.FetchBtn1, "FetchBtn1");
            this.FetchBtn1.Name = "FetchBtn1";
            this.FetchBtn1.UseVisualStyleBackColor = true;
            this.FetchBtn1.Click += new System.EventHandler(this.FetchBtn1_Click);
            // 
            // ItemBox1
            // 
            this.ItemBox1.AcceptsTab = true;
            resources.ApplyResources(this.ItemBox1, "ItemBox1");
            this.ItemBox1.Name = "ItemBox1";
            // 
            // LotBox1
            // 
            this.LotBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            resources.ApplyResources(this.LotBox1, "LotBox1");
            this.LotBox1.Name = "LotBox1";
            // 
            // VendorBox1
            // 
            resources.ApplyResources(this.VendorBox1, "VendorBox1");
            this.VendorBox1.FormattingEnabled = true;
            this.VendorBox1.Items.AddRange(new object[] {
            resources.GetString("VendorBox1.Items"),
            resources.GetString("VendorBox1.Items1"),
            resources.GetString("VendorBox1.Items2"),
            resources.GetString("VendorBox1.Items3"),
            resources.GetString("VendorBox1.Items4")});
            this.VendorBox1.Name = "VendorBox1";
            // 
            // clrCoABtn
            // 
            resources.ApplyResources(this.clrCoABtn, "clrCoABtn");
            this.clrCoABtn.Name = "clrCoABtn";
            this.clrCoABtn.UseVisualStyleBackColor = true;
            this.clrCoABtn.Click += new System.EventHandler(this.clrCoABtn_Click);
            // 
            // QRRFindBtn
            // 
            resources.ApplyResources(this.QRRFindBtn, "QRRFindBtn");
            this.QRRFindBtn.Name = "QRRFindBtn";
            this.QRRFindBtn.UseVisualStyleBackColor = true;
            this.QRRFindBtn.Click += new System.EventHandler(this.QRRFindBtn_Click);
            // 
            // QRRLabel
            // 
            resources.ApplyResources(this.QRRLabel, "QRRLabel");
            this.QRRLabel.Name = "QRRLabel";
            // 
            // CoALabel
            // 
            resources.ApplyResources(this.CoALabel, "CoALabel");
            this.CoALabel.Name = "CoALabel";
            // 
            // ComboPDFSection
            // 
            resources.ApplyResources(this.ComboPDFSection, "ComboPDFSection");
            this.ComboPDFSection.Name = "ComboPDFSection";
            // 
            // QRRFileNameBx
            // 
            resources.ApplyResources(this.QRRFileNameBx, "QRRFileNameBx");
            this.QRRFileNameBx.Name = "QRRFileNameBx";
            // 
            // QRRPDFBtn
            // 
            resources.ApplyResources(this.QRRPDFBtn, "QRRPDFBtn");
            this.QRRPDFBtn.Name = "QRRPDFBtn";
            this.QRRPDFBtn.UseVisualStyleBackColor = true;
            this.QRRPDFBtn.Click += new System.EventHandler(this.QRRPDFBtn_Click);
            // 
            // CoAPDFBtn
            // 
            resources.ApplyResources(this.CoAPDFBtn, "CoAPDFBtn");
            this.CoAPDFBtn.Name = "CoAPDFBtn";
            this.CoAPDFBtn.UseVisualStyleBackColor = true;
            this.CoAPDFBtn.Click += new System.EventHandler(this.CoAPDFBtn_Click);
            // 
            // CoAFileNameBx
            // 
            resources.ApplyResources(this.CoAFileNameBx, "CoAFileNameBx");
            this.CoAFileNameBx.Name = "CoAFileNameBx";
            // 
            // SaveFinalPDFBtn
            // 
            resources.ApplyResources(this.SaveFinalPDFBtn, "SaveFinalPDFBtn");
            this.SaveFinalPDFBtn.Name = "SaveFinalPDFBtn";
            this.SaveFinalPDFBtn.UseVisualStyleBackColor = true;
            this.SaveFinalPDFBtn.Click += new System.EventHandler(this.SaveFinalPDFBtn_Click);
            // 
            // PDFNameLabel
            // 
            resources.ApplyResources(this.PDFNameLabel, "PDFNameLabel");
            this.PDFNameLabel.Name = "PDFNameLabel";
            // 
            // FinalPDFBx
            // 
            resources.ApplyResources(this.FinalPDFBx, "FinalPDFBx");
            this.FinalPDFBx.Name = "FinalPDFBx";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // Form1
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label1);
            this.Controls.Add(this.SaveFinalPDFBtn);
            this.Controls.Add(this.PDFNameLabel);
            this.Controls.Add(this.FinalPDFBx);
            this.Controls.Add(this.CoAFileNameBx);
            this.Controls.Add(this.CoAPDFBtn);
            this.Controls.Add(this.QRRPDFBtn);
            this.Controls.Add(this.QRRFileNameBx);
            this.Controls.Add(this.ComboPDFSection);
            this.Controls.Add(this.CoALabel);
            this.Controls.Add(this.QRRLabel);
            this.Controls.Add(this.QRRFindBtn);
            this.Controls.Add(this.clrCoABtn);
            this.Controls.Add(this.VendorBox1);
            this.Controls.Add(this.LotBox1);
            this.Controls.Add(this.ItemBox1);
            this.Controls.Add(this.FetchBtn1);
            this.Name = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button FetchBtn1;
        private System.Windows.Forms.TextBox ItemBox1;
        private System.Windows.Forms.TextBox LotBox1;
        private System.Windows.Forms.ComboBox VendorBox1;
        private System.Windows.Forms.Button clrCoABtn;
        private System.Windows.Forms.Button QRRFindBtn;
        private System.Windows.Forms.Label QRRLabel;
        private System.Windows.Forms.Label CoALabel;
        private System.Windows.Forms.Label ComboPDFSection;
        private System.Windows.Forms.TextBox QRRFileNameBx;
        private System.Windows.Forms.Button QRRPDFBtn;
        private System.Windows.Forms.Button CoAPDFBtn;
        private System.Windows.Forms.TextBox CoAFileNameBx;
        private System.Windows.Forms.Button SaveFinalPDFBtn;
        private System.Windows.Forms.Label PDFNameLabel;
        private System.Windows.Forms.TextBox FinalPDFBx;
        private System.Windows.Forms.Label label1;
    }
}

