using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO; 
using xls = Microsoft.Office.Interop.Excel;


namespace EoscarProduction
{
	public partial class Form1 : Form
	{

		public string prodPath = @"D:\Joel Files\Payroll\Eoscar.csv";
		public string holidayPath = @"D:\Joel Files\Payroll\holiday.csv";
		public string clientErrorPath = @"D:\Joel Files\Payroll\ClientError.csv";
		public string otherIncomePath = @"D:\Joel Files\Payroll\OtherIncome.csv";
		public string fraudPath = @"D:\Joel Files\Payroll\";
		public string banckruptcyPath = @"D:\Joel Files\Payroll\banckruptcy.csv";
		public string linkPath = @"D:\Joel Files\Payroll\link.csv";
		public string outclientpath = @"D:\Joel Files\Payroll\";
		public string outputPath = @"D:\Joel Files\Payroll\";

		public Form1()
		{
			InitializeComponent();
		}


		private void LoadExcel(String path)
		{
			xls.Application xl = new xls.Application();
			xls.Workbooks wbs = xl.Workbooks;
			xls.Workbook wb = wbs.Open(path);

			xls.Worksheet ws = wb.Worksheets["Production"];

			long row = 1;
			bool isErrorReport = false;
			decimal Amount = 0;
			StreamWriter fsProd = new StreamWriter(prodPath, false);
			StreamWriter fsHoliday = new StreamWriter(holidayPath, false);
			StreamWriter fsClientError = new StreamWriter(clientErrorPath, false);
			StreamWriter fsOthereIncome = new StreamWriter(otherIncomePath, false);

			while (true)
			{
				this.Text = string.Format("Creating Production Files: {0}", row);

				xls.Range r = ws.Range["A" + row, "R" + row];

				if (r.Cells[1, 6].Text == "Client Reported")
				{
					isErrorReport = true;
				}

				//do client errors 
				if (isErrorReport && decimal.TryParse(r.Cells[1, 14].Text, out Amount))
				{
					if (Amount != 0)
					{
						//4 userid, 
						fsClientError.WriteLine(string.Format("{0},{1}", r.Cells[1, 4].Text, Amount));
					}
				}

				if (!isErrorReport)
				{
					if (r.Cells[1, 14].Text != string.Empty) //N
					{
						//do total amount here
						fsProd.WriteLine(string.Format("{0},{1}", r.Cells[1, 14].Text, r.Cells[1, 13].Text));
					}
					else if (r.Cells[1, 11].Text != string.Empty) //holiday - K 11
					{
						if (decimal.TryParse(r.Cells[1, 11].Text, out Amount))
						{
							if (Amount != 0)
							{
								//do holiday here
								//2 -- date, 11-holiday,1- userid
								fsHoliday.WriteLine(string.Format("{0},{1},{2:0000}", r.Cells[1, 2].Text, Amount, r.Cells[1, 1].Text));
							}
						}
					}
				}

				if (isErrorReport && ws.Range["N" + row].Value == null)
				{
					//close range
					Marshal.FinalReleaseComObject(r);
					break;
				}

				//close range
				Marshal.FinalReleaseComObject(r);
				row += 1;
			}


			//release all com objects 
			Marshal.FinalReleaseComObject(ws);

			//otherincome
			try
			{
				ws = wb.Worksheets["Other Income"];
				row = 3;
				int blank = 0;
				int userid = 0;
				string remarks = string.Empty;
				decimal amount = 0;
				int prevUserid = 0; 
				while (true)
				{
					this.Text = string.Format("Creating Other Income Files: {0}", row);
					xls.Range r = ws.Range["A" + row, "I" + row];

					if (r.Cells[1, 1].Text + r.Cells[1, 2].Text != string.Empty)
					{
						blank = 0;
						if (int.TryParse(r.Cells[1, 1].Text, out userid))
						{
							if (prevUserid != userid)
							{
								if (prevUserid != 0)
								{
									fsOthereIncome.WriteLine(string.Format("{0:0000}, {1}, {2}", prevUserid, remarks, amount));
								}
								prevUserid = userid;
								amount = decimal.Parse(r.Cells[8].Text);
							}
							else
							{
								amount += decimal.Parse(r.Cells[8].Text);
							}
						}
						else
						{
						
							if (prevUserid != 0)
							{
								fsOthereIncome.WriteLine(string.Format("{0:0000}, {1}, {2}", prevUserid, remarks, amount));
							}
							prevUserid = 0;
							amount = 0;
							remarks = r.Cells[1, 1].Text + r.Cells[1, 2].Text;
						}
					}
					else
					{
						blank += 1;
						if (blank > 5)
						{
							//write last person
							fsOthereIncome.WriteLine(string.Format("{0:0000}, {1}, {2}", prevUserid, remarks, amount));

							//close range
							Marshal.FinalReleaseComObject(r);
							break;
						}
					}
					Marshal.FinalReleaseComObject(r);
					row += 1;
				}
			}
			catch (Exception)
			{
			}

			//release all com objects 
			Marshal.FinalReleaseComObject(ws);

			wb.Close();
			Marshal.FinalReleaseComObject(wb);

			xl.Quit();
			Marshal.FinalReleaseComObject(xl);

			fsProd.Close();
			fsHoliday.Close();
			fsClientError.Close();
			fsOthereIncome.Close();

			MessageBox.Show("Done");
		}


		private void LoadExcelOutclient(String path)
		{
			xls.Application xl = new xls.Application();
			xls.Workbooks wbs = xl.Workbooks;
			xls.Workbook wb = wbs.Open(path);

			xls.Worksheet ws = wb.Worksheets["Production"];

			long row = 1;
			bool isErrorReport = false;
			decimal Amount = 0;
			StreamWriter fsProd = new StreamWriter(string.Format("{0}outclient.csv",outclientpath), false);
			StreamWriter fsHoliday = new StreamWriter(string.Format("{0}outclientholiday.csv", outclientpath), false);
			StreamWriter fsClientError = new StreamWriter(string.Format("{0}outclienterror.csv", outclientpath), false);
			StreamWriter fsOthereIncome = new StreamWriter(string.Format("{0}outclientotherincome.csv", outclientpath), false);

			while (true)
			{
				this.Text = string.Format("Creating Production Files: {0}", row);

				xls.Range r = ws.Range["A" + row, "R" + row];

				if (r.Cells[1, 6].Text == "Client Reported")
				{
					isErrorReport = true;
				}

				//do client errors 
				if (isErrorReport && decimal.TryParse(r.Cells[1, 14].Text, out Amount))
				{
					if (Amount != 0)
					{
						//3 userid, 
						fsClientError.WriteLine(string.Format("{0},{1}", r.Cells[1, 4].Text, Amount));
					}
				}

				if (!isErrorReport)
				{
					if (r.Cells[1, 14].Text != string.Empty) //N
					{
						//do total amount here
						fsProd.WriteLine(string.Format("{0},{1}", r.Cells[1, 14].Text, r.Cells[1, 13].Text));
					}
					else if (r.Cells[1, 11].Text != string.Empty) //holiday - K 11
					{
						if (decimal.TryParse(r.Cells[1, 11].Text, out Amount))
						{
							if (Amount != 0)
							{
								//do holiday here
								//2 -- date, 11-holiday,1- userid
								fsHoliday.WriteLine(string.Format("{0},{1},{2:0000}", r.Cells[1, 2].Text, Amount, r.Cells[1, 1].Text));
							}
						}
					}
				}

				if (isErrorReport && ws.Range["N" + row].Value == null)
				{
					//close range
					Marshal.FinalReleaseComObject(r);
					break;
				}

				//close range
				Marshal.FinalReleaseComObject(r);
				row += 1;
			}


			//release all com objects 
			Marshal.FinalReleaseComObject(ws);

			//otherincome
			try
			{
				ws = wb.Worksheets["Other Income"];
				row = 3;
				int blank = 0;
				int userid = 0;
				string remarks = string.Empty;
				decimal amount = 0;
				int prevUserid = 0;
				while (true)
				{
					this.Text = string.Format("Creating Other Income Files: {0}", row);
					xls.Range r = ws.Range["A" + row, "I" + row];

					if (r.Cells[1, 1].Text + r.Cells[1, 2].Text != string.Empty)
					{
						blank = 0;
						if (int.TryParse(r.Cells[1, 1].Text, out userid))
						{
							if (prevUserid != userid)
							{
								if (prevUserid != 0)
								{
									fsOthereIncome.WriteLine(string.Format("{0:0000}, {1}, {2}", prevUserid, remarks, amount));
								}
								prevUserid = userid;
								amount = decimal.Parse(r.Cells[8].Text);
							}
							else
							{
								amount += decimal.Parse(r.Cells[8].Text);
							}
						}
						else
						{

							if (prevUserid != 0)
							{
								fsOthereIncome.WriteLine(string.Format("{0:0000}, {1}, {2}", prevUserid, remarks, amount));
							}
							prevUserid = 0;
							amount = 0;
							remarks = r.Cells[1, 1].Text + r.Cells[1, 2].Text;
						}
					}
					else
					{
						blank += 1;
						if (blank > 5)
						{
							//write last person
							fsOthereIncome.WriteLine(string.Format("{0:0000}, {1}, {2}", prevUserid, remarks, amount));

							//close range
							Marshal.FinalReleaseComObject(r);
							break;
						}
					}
					Marshal.FinalReleaseComObject(r);
					row += 1;
				}
			}
			catch (Exception)
			{
			}

			//release all com objects 
			Marshal.FinalReleaseComObject(ws);

			wb.Close();
			Marshal.FinalReleaseComObject(wb);

			xl.Quit();
			Marshal.FinalReleaseComObject(xl);

			fsProd.Close();
			fsHoliday.Close();
			fsClientError.Close();
			fsOthereIncome.Close();

			MessageBox.Show("Done");
		}

		private void btn_Production_Click(object sender, EventArgs e)
		{
			//LoadExcel(@"C:\Joel Files\Payroll\20200511\20200426-20200510 eOscar PS.xlsx");
			switch (((Button)sender).Name) {
				case "btn_outclient":
					LoadExcelOutclient(txtOutclient.Text);
					break;
				default:
					LoadExcel(txtExcelFile.Text);
					break; 
			}
			
		}

		private void btnBrowse_Click(object sender, EventArgs e)
		{
			openFileDialog1.CheckFileExists = true;
			switch (((Button)sender).Name)
			{
				case "btn_browse_paper":
					openFileDialog1.Filter = "Paper Bankruptcy|*Paper*.xlsx";
					if (openFileDialog1.ShowDialog() == DialogResult.OK) txt_Paper.Text = openFileDialog1.FileName;
					break;
				case "btn_browse_Ptask73":
					openFileDialog1.Filter = "PTask73|*Ptask73*.xlsx";
					if (openFileDialog1.ShowDialog() == DialogResult.OK) txt_Ptask73.Text = openFileDialog1.FileName;
					break;
				case "btn_browse_deceased":
					openFileDialog1.Filter = "Deceased|*Deceased*.xlsx";
					if (openFileDialog1.ShowDialog() == DialogResult.OK) txt_deceased.Text = openFileDialog1.FileName;
					break;
				case "btn_browse_outclient":
					openFileDialog1.Filter = "Outclient|*outclient*.xlsx";
					if (openFileDialog1.ShowDialog() == DialogResult.OK) txtOutclient.Text = openFileDialog1.FileName;
					break;
				case "btn_browse_audit_aged":
					openFileDialog1.Filter = "Audit Aged|*Audit Aged*.xlsx";
					if (openFileDialog1.ShowDialog() == DialogResult.OK)
					{
						txt_audit_aged.Text = openFileDialog1.FileName;
					}
					break;
				default:
					openFileDialog1.Filter = "Eoscar PS|*Eoscar*.xlsx";
					if (openFileDialog1.ShowDialog() == DialogResult.OK)
					{
						txtExcelFile.Text = openFileDialog1.FileName;
					}
					break;
			}

		}

		private void btn_browse_fraud_Click(object sender, EventArgs e) 
		{
			openFileDialog1.Filter = "Fraud PS|*Fraud*.xlsx";
			openFileDialog1.CheckFileExists = true;
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				txtFraudProduction.Text = openFileDialog1.FileName;
			}
		}

		private void getPayroll(string excelPath , string output, string errorPath = null ) {

			if (excelPath == string.Empty) { return; }

			StreamWriter fraud = File.CreateText(output);
			bool isOutputError = false;
			StreamWriter outputError = null; 
			if (!(errorPath is  null) )
			{
				outputError = File.CreateText(errorPath);
				isOutputError = true; 
			}

			xls.Application xl = new xls.Application();
			xls.Workbooks wbs = xl.Workbooks;
			xls.Workbook wb = wbs.Open(excelPath);
			xls.Worksheet ws = wb.Worksheets["Production"];

			int blankCounter = 0;
			int iRow = 2;
			int errorRow = 1;
			bool onErrorRange = false;
			string userid = "";
			decimal amount = 0;

			while (true)
			{
				xls.Range r = ws.Range["A" + iRow, "H" + iRow];


				if (r.Cells[1, 1].Text == string.Empty)
				{
					blankCounter += 1;

					if (userid != string.Empty ) fraud.WriteLine(string.Format("{0},{1}", userid, amount));
					userid = string.Empty;
				}
				else
				{
					blankCounter = 1;


					//do client errors 
					if (isOutputError && onErrorRange && decimal.TryParse(r.Cells[1, errorRow].Text, out amount))
					{
						if (amount != 0)
						{
							//3 userid, 
							outputError.WriteLine(string.Format("{0},{1}", r.Cells[1, 1].Text, amount));
						}
					}

					if (!onErrorRange)
					{
						//check if on error range 
						if (ws.Range["F"+iRow].Text == "Error Amount")
						{
							onErrorRange = true;
							errorRow = 6; 
						}
						else if (ws.Range["G" + iRow].Text == "Error Amount")
						{
							onErrorRange = true;
							errorRow = 11; 
						}
						else { 
						
							if (userid == r.Cells[1, 2].Text && userid != string.Empty)
							{
								amount += decimal.Parse(r.Cells[1, 8].Text);
							}
							else
							{
								//insert on text file
								if (userid != string.Empty) fraud.WriteLine(string.Format("{0},{1}", userid, amount));

								userid = r.Cells[1, 2].Text;
								amount = decimal.Parse(r.Cells[1, 8].Text);
							}
						}
					}
				}

				if (blankCounter > 10)
				{
					break;
				}
				iRow += 1;
			}
			//write last person
			if(userid != "") fraud.WriteLine(string.Format("{0},{1}", userid, amount));

			//release all com objects 
			Marshal.FinalReleaseComObject(ws);

			wb.Close();
			Marshal.FinalReleaseComObject(wb);

			xl.Quit();
			Marshal.FinalReleaseComObject(xl);

			fraud.Close();
			outputError.Close();

			MessageBox.Show("Done");
		}

		private void btn_Fraud_Click(object sender, EventArgs e)
		{
			switch (((Button)sender).Name)
			{
				case "btn_link":
					getPayroll(txt_link.Text, outputPath + "Link.csv", outputPath + "LinkError.csv");
					break;
				case "btn_Banckruptcy":
					getPayroll(txtBanckruptcyProduction.Text, outputPath + "Bankruptcy.csv", outputPath + "BankruptcyError.csv");
					break; 
				case "btn_paper":
					getPayroll(txt_Paper.Text, outputPath + "Paper.csv", outputPath + "PaperError.csv");
					break;
				case "btn_Ptask73":
					getPayroll(txt_Ptask73.Text, outputPath + "Ptask73.csv", outputPath + "Ptask73Error.csv");
					break;
				case "btn_deceased":
					getPayroll(txt_deceased.Text, outputPath + "deceased.csv", outputPath + "deceasedError.csv");
					break;
				case "btn_audit_aged":
					getPayroll(txt_audit_aged.Text, outputPath + "AuditAge.csv", outputPath + "AuditAgedError.csv"); 
					break;
				default:
					getPayroll(txtFraudProduction.Text, outputPath + "Fraud.csv", outputPath + "fraud_error.csv");
					break;
			}
		}

		private void btn_browse_Banckruptcy_Click(object sender, EventArgs e)
		{
			openFileDialog1.Filter = "Bankruptcy PS|*Bankruptcy*.xlsx";
			openFileDialog1.CheckFileExists = true;
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				txtBanckruptcyProduction.Text = openFileDialog1.FileName;
			}
		}

		private void btn_Banckruptcy_Click(object sender, EventArgs e)
		{
			if (txtBanckruptcyProduction.Text.Trim() == string.Empty) { return; }

			StreamWriter fraud = File.CreateText(banckruptcyPath);

			xls.Application xl = new xls.Application();
			xls.Workbooks wbs = xl.Workbooks;
			xls.Workbook wb = wbs.Open(txtBanckruptcyProduction.Text);
			xls.Worksheet ws = wb.Worksheets["Production"];

			int blankCounter = 0;
			int iRow = 2;

			string userid = "";
			decimal amount = 0;

			while (true)
			{
				xls.Range r = ws.Range["A" + iRow, "H" + iRow];


				if (r.Cells[1, 1].Text == string.Empty)
				{
					blankCounter += 1;
				}
				else
				{
					blankCounter = 1;
					if (userid == r.Cells[1, 2].Text && userid != string.Empty)
					{
						amount += decimal.Parse(r.Cells[1, 8].Text);
					}
					else
					{

						//insert on text file
						fraud.WriteLine(string.Format("{0},{1}", userid, amount));

						userid = r.Cells[1, 2].Text;
						amount = decimal.Parse(r.Cells[1, 8].Text);
					}
				}

				if (blankCounter > 5)
				{
					break;
				}
				iRow += 1;
			}
			//write last person
			fraud.WriteLine(string.Format("{0},{1}", userid, amount));

			//release all com objects 
			Marshal.FinalReleaseComObject(ws);

			wb.Close();
			Marshal.FinalReleaseComObject(wb);

			xl.Quit();
			Marshal.FinalReleaseComObject(xl);

			fraud.Close();

			MessageBox.Show("Done");
		}

		private void btn_browse_link_Click(object sender, EventArgs e)
		{
			openFileDialog1.Filter = "Link PS|*Link*.xlsx";
			openFileDialog1.CheckFileExists = true;
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				txt_link.Text = openFileDialog1.FileName;
			}
		}

		private void btn_link_Click(object sender, EventArgs e)
		{
			if (txt_link.Text.Trim() == string.Empty) { return; }

			StreamWriter fraud = File.CreateText(linkPath);

			xls.Application xl = new xls.Application();
			xls.Workbooks wbs = xl.Workbooks;
			xls.Workbook wb = wbs.Open(txt_link.Text);
			xls.Worksheet ws = wb.Worksheets["Production"];

			int blankCounter = 0;
			int iRow = 2;

			string userid = "";
			decimal amount = 0;

			while (true)
			{
				xls.Range r = ws.Range["A" + iRow, "H" + iRow];


				if (r.Cells[1, 1].Text == string.Empty)
				{
					blankCounter += 1;
				}
				else
				{
					blankCounter = 1;
					if (userid == r.Cells[1, 2].Text && userid != string.Empty)
					{
						amount += decimal.Parse(r.Cells[1, 8].Text);
					}
					else
					{

						//insert on text file
						fraud.WriteLine(string.Format("{0},{1}", userid, amount));

						userid = r.Cells[1, 2].Text;
						amount = decimal.Parse(r.Cells[1, 8].Text);
					}
				}

				if (blankCounter > 5)
				{
					break;
				}
				iRow += 1;
			}
			//write last person
			fraud.WriteLine(string.Format("{0},{1}", userid, amount));

			//release all com objects 
			Marshal.FinalReleaseComObject(ws);

			wb.Close();
			Marshal.FinalReleaseComObject(wb);

			xl.Quit();
			Marshal.FinalReleaseComObject(xl);

			fraud.Close();

			MessageBox.Show("Done");
		}

	}
	
}
