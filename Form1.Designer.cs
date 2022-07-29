namespace EoscarProduction
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
            this.btn_Production = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.txtExcelFile = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_browse_fraud = new System.Windows.Forms.Button();
            this.txtFraudProduction = new System.Windows.Forms.TextBox();
            this.btn_Fraud = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.btn_browse_Banckruptcy = new System.Windows.Forms.Button();
            this.txtBanckruptcyProduction = new System.Windows.Forms.TextBox();
            this.btn_Banckruptcy = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.btn_browse_link = new System.Windows.Forms.Button();
            this.txt_link = new System.Windows.Forms.TextBox();
            this.btn_link = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.txtOscarProduction = new System.Windows.Forms.TextBox();
            this.btn_getProduction = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btn_Production
            // 
            this.btn_Production.Location = new System.Drawing.Point(551, 35);
            this.btn_Production.Name = "btn_Production";
            this.btn_Production.Size = new System.Drawing.Size(96, 26);
            this.btn_Production.TabIndex = 0;
            this.btn_Production.Text = "Get Production";
            this.btn_Production.UseVisualStyleBackColor = true;
            this.btn_Production.Click += new System.EventHandler(this.btn_Production_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // txtExcelFile
            // 
            this.txtExcelFile.Location = new System.Drawing.Point(12, 39);
            this.txtExcelFile.Name = "txtExcelFile";
            this.txtExcelFile.Size = new System.Drawing.Size(479, 20);
            this.txtExcelFile.TabIndex = 1;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(505, 39);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(31, 19);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(125, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "EOscar Production Excel";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 62);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(117, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Fraud Production Excel";
            // 
            // btn_browse_fraud
            // 
            this.btn_browse_fraud.Location = new System.Drawing.Point(505, 81);
            this.btn_browse_fraud.Name = "btn_browse_fraud";
            this.btn_browse_fraud.Size = new System.Drawing.Size(31, 19);
            this.btn_browse_fraud.TabIndex = 6;
            this.btn_browse_fraud.Text = "...";
            this.btn_browse_fraud.UseVisualStyleBackColor = true;
            this.btn_browse_fraud.Click += new System.EventHandler(this.btn_browse_fraud_Click);
            // 
            // txtFraudProduction
            // 
            this.txtFraudProduction.Location = new System.Drawing.Point(12, 81);
            this.txtFraudProduction.Name = "txtFraudProduction";
            this.txtFraudProduction.Size = new System.Drawing.Size(479, 20);
            this.txtFraudProduction.TabIndex = 5;
            // 
            // btn_Fraud
            // 
            this.btn_Fraud.Location = new System.Drawing.Point(551, 77);
            this.btn_Fraud.Name = "btn_Fraud";
            this.btn_Fraud.Size = new System.Drawing.Size(96, 26);
            this.btn_Fraud.TabIndex = 4;
            this.btn_Fraud.Text = "Get Production";
            this.btn_Fraud.UseVisualStyleBackColor = true;
            this.btn_Fraud.Click += new System.EventHandler(this.btn_Fraud_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 111);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(150, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "Banckruptcy Production Excel";
            // 
            // btn_browse_Banckruptcy
            // 
            this.btn_browse_Banckruptcy.Location = new System.Drawing.Point(505, 130);
            this.btn_browse_Banckruptcy.Name = "btn_browse_Banckruptcy";
            this.btn_browse_Banckruptcy.Size = new System.Drawing.Size(31, 19);
            this.btn_browse_Banckruptcy.TabIndex = 10;
            this.btn_browse_Banckruptcy.Text = "...";
            this.btn_browse_Banckruptcy.UseVisualStyleBackColor = true;
            this.btn_browse_Banckruptcy.Click += new System.EventHandler(this.btn_browse_Banckruptcy_Click);
            // 
            // txtBanckruptcyProduction
            // 
            this.txtBanckruptcyProduction.Location = new System.Drawing.Point(12, 130);
            this.txtBanckruptcyProduction.Name = "txtBanckruptcyProduction";
            this.txtBanckruptcyProduction.Size = new System.Drawing.Size(479, 20);
            this.txtBanckruptcyProduction.TabIndex = 9;
            // 
            // btn_Banckruptcy
            // 
            this.btn_Banckruptcy.Location = new System.Drawing.Point(551, 126);
            this.btn_Banckruptcy.Name = "btn_Banckruptcy";
            this.btn_Banckruptcy.Size = new System.Drawing.Size(96, 26);
            this.btn_Banckruptcy.TabIndex = 8;
            this.btn_Banckruptcy.Text = "Get Production";
            this.btn_Banckruptcy.UseVisualStyleBackColor = true;
            this.btn_Banckruptcy.Click += new System.EventHandler(this.btn_Banckruptcy_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(16, 161);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(110, 13);
            this.label4.TabIndex = 15;
            this.label4.Text = "Link Production Excel";
            // 
            // btn_browse_link
            // 
            this.btn_browse_link.Location = new System.Drawing.Point(505, 180);
            this.btn_browse_link.Name = "btn_browse_link";
            this.btn_browse_link.Size = new System.Drawing.Size(31, 19);
            this.btn_browse_link.TabIndex = 14;
            this.btn_browse_link.Text = "...";
            this.btn_browse_link.UseVisualStyleBackColor = true;
            this.btn_browse_link.Click += new System.EventHandler(this.btn_browse_link_Click);
            // 
            // txt_link
            // 
            this.txt_link.Location = new System.Drawing.Point(12, 180);
            this.txt_link.Name = "txt_link";
            this.txt_link.Size = new System.Drawing.Size(479, 20);
            this.txt_link.TabIndex = 13;
            // 
            // btn_link
            // 
            this.btn_link.Location = new System.Drawing.Point(551, 176);
            this.btn_link.Name = "btn_link";
            this.btn_link.Size = new System.Drawing.Size(96, 26);
            this.btn_link.TabIndex = 12;
            this.btn_link.Text = "Get Production";
            this.btn_link.UseVisualStyleBackColor = true;
            this.btn_link.Click += new System.EventHandler(this.btn_link_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(572, 275);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 42);
            this.button1.TabIndex = 16;
            this.button1.Text = "Get Production";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtOscarProduction
            // 
            this.txtOscarProduction.Location = new System.Drawing.Point(12, 297);
            this.txtOscarProduction.Name = "txtOscarProduction";
            this.txtOscarProduction.Size = new System.Drawing.Size(479, 20);
            this.txtOscarProduction.TabIndex = 17;
            // 
            // btn_getProduction
            // 
            this.btn_getProduction.Location = new System.Drawing.Point(505, 298);
            this.btn_getProduction.Name = "btn_getProduction";
            this.btn_getProduction.Size = new System.Drawing.Size(31, 19);
            this.btn_getProduction.TabIndex = 18;
            this.btn_getProduction.Text = "...";
            this.btn_getProduction.UseVisualStyleBackColor = true;
            this.btn_getProduction.Click += new System.EventHandler(this.btn_getProduction_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btn_getProduction);
            this.Controls.Add(this.txtOscarProduction);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btn_browse_link);
            this.Controls.Add(this.txt_link);
            this.Controls.Add(this.btn_link);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btn_browse_Banckruptcy);
            this.Controls.Add(this.txtBanckruptcyProduction);
            this.Controls.Add(this.btn_Banckruptcy);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btn_browse_fraud);
            this.Controls.Add(this.txtFraudProduction);
            this.Controls.Add(this.btn_Fraud);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtExcelFile);
            this.Controls.Add(this.btn_Production);
            this.Name = "Form1";
            this.Text = "Process EOscar";
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button btn_Production;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.TextBox txtExcelFile;
		private System.Windows.Forms.Button btnBrowse;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btn_browse_fraud;
		private System.Windows.Forms.TextBox txtFraudProduction;
		private System.Windows.Forms.Button btn_Fraud;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Button btn_browse_Banckruptcy;
		private System.Windows.Forms.TextBox txtBanckruptcyProduction;
		private System.Windows.Forms.Button btn_Banckruptcy;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button btn_browse_link;
		private System.Windows.Forms.TextBox txt_link;
		private System.Windows.Forms.Button btn_link;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtOscarProduction;
        private System.Windows.Forms.Button btn_getProduction;
    }
}

