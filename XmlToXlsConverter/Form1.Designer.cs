
namespace XmlToXlsConverter
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.txtXmlFilePath = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.chkCustomName = new System.Windows.Forms.CheckBox();
            this.txtCustomFileName = new System.Windows.Forms.TextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.button2 = new System.Windows.Forms.Button();
            this.OFD = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(212, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select XML file to convert";
            // 
            // txtXmlFilePath
            // 
            this.txtXmlFilePath.Location = new System.Drawing.Point(13, 42);
            this.txtXmlFilePath.Name = "txtXmlFilePath";
            this.txtXmlFilePath.Size = new System.Drawing.Size(323, 31);
            this.txtXmlFilePath.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(350, 39);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(112, 34);
            this.button1.TabIndex = 2;
            this.button1.Text = "Browse";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // chkCustomName
            // 
            this.chkCustomName.AutoSize = true;
            this.chkCustomName.Location = new System.Drawing.Point(13, 95);
            this.chkCustomName.Name = "chkCustomName";
            this.chkCustomName.Size = new System.Drawing.Size(159, 29);
            this.chkCustomName.TabIndex = 3;
            this.chkCustomName.Text = "Excel File Name";
            this.chkCustomName.UseVisualStyleBackColor = true;
            // 
            // txtCustomFileName
            // 
            this.txtCustomFileName.Location = new System.Drawing.Point(179, 95);
            this.txtCustomFileName.Name = "txtCustomFileName";
            this.txtCustomFileName.Size = new System.Drawing.Size(283, 31);
            this.txtCustomFileName.TabIndex = 4;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(13, 148);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(323, 34);
            this.progressBar1.TabIndex = 5;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(350, 147);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(112, 34);
            this.button2.TabIndex = 6;
            this.button2.Text = "Convert";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // OFD
            // 
            this.OFD.Filter = "XML File (*.xml)|*.xml|All files (*.*)|*.*";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(474, 197);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.txtCustomFileName);
            this.Controls.Add(this.chkCustomName);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtXmlFilePath);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Xml to Xls Converter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtXmlFilePath;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox chkCustomName;
        private System.Windows.Forms.TextBox txtCustomFileName;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.OpenFileDialog OFD;
    }
}

