namespace Products3._0
{
    partial class Instructions
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Instructions));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.listBox1);
            this.groupBox1.Location = new System.Drawing.Point(5, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(921, 226);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Instructions";
            // 
            // listBox1
            // 
            this.listBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBox1.FormattingEnabled = true;
            this.listBox1.HorizontalScrollbar = true;
            this.listBox1.ItemHeight = 20;
            this.listBox1.Items.AddRange(new object[] {
            "*Install Microsoft Office 2007 or above  .",
            "***The Tool Work Perfectly With Microsoft office 2019 Pro Plus Retail ,and no nee" +
                "d to install external tools",
            "*if installed Office 2007 -> install AccessDatabaseEngine/ 64bit according to you" +
                "r computer system ,from directory folder.",
            "1- The original excel file(xls.xlsx) is found in directory ,is opened automatical" +
                "ly.",
            "2- please use original file , can edit this file . add,delete rows is possible.",
            "3- Don\'t use the new excel file you created.",
            "4- use(.) not (,) in the Price Column .",
            "5- in Vat textbox use just numbers without (.) -> example you put (17) that mean " +
                "17%=0.17",
            "6- price should be between 0 and 800000.",
            "7- count should be between 0 and 1000.",
            "8- Save the file in another location (Don\'t lose the  original excel file).",
            "9- before send a mail make sure you change (Allow less secure apps: ON) from E-ma" +
                "il ",
            "setting->Sign-in & security->Apps with account access.",
            "10- before send a mail in the directory folder ,change the logo picture to your ," +
                "with the same name (logo.jpg)"});
            this.listBox1.Location = new System.Drawing.Point(5, 16);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(911, 204);
            this.listBox1.TabIndex = 0;
            // 
            // Instructions
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(929, 233);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Instructions";
            this.Text = "Price Offer 1.0";
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ListBox listBox1;
    }
}