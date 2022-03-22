using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Net.Mail;

namespace SendMailThue
{
    public partial class MainForm : Form
    {
        private List<Company> companies = new List<Company>();
        public MainForm()
        {
            InitializeComponent();
            this.progressBar.Minimum = 0;
            this.progressBar.Value = 0;
            this.progressBar.Step = 1;
            this.progressBar.Visible = false;

        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var file = openFileDialog.FileName;
                    handleExcelFile(file);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
            }
        }
        private void handleExcelFile(string filePath) 
        {
            //txtFilePath.Text = filePath;
            //companies = getCompanys(filePath);
            //companies.Add(new Company { Name = "name 1", Email = "Email 1", DuNo = "Du no 1" }) ;
            //companies.Add(new Company { Name = "name 1", Email = "Email 1", DuNo = "Du no 1" }) ;
            //companies.Add(new Company { Name = "name 1", Email = "Email 1", DuNo = "Du no 1" }) ;
            //dgvCompany.DataSource = companies;
            //dgvCompany.DefaultCellStyle.ForeColor = Color.Black;
            //lblCompanyCount.Text = "Total:  " + companies.Count;
        }


        private static List<Company> getCompanys(string excelFile)
        {

            List<Company> cpns = new List<Company>();

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel._Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;
            try
            {
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(excelFile);
                xlWorksheet = xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;


                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                var values = (xlRange.Value as Object[,]);


                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        try
                        {
                            //new line
                            string cellValue = Convert.ToString(values[i, j]);
                            int a = 10;
                        }
                        catch (Exception e) { }

                        //add useful things here!   
                    }


                }
            }
            catch (Exception ex) { }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                if (xlWorkbook != null)
                {
                    xlWorkbook.Close();
                }
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                if (xlApp != null)
                {
                    xlApp.Quit();
                }
               
                Marshal.ReleaseComObject(xlApp);
            }

            return cpns;
        }

        private void btnSendMail_Click(object sender, EventArgs e)
        {
            btnSendMail.Enabled = false;
            progressBar.Maximum = companies.Count;
            progressBar.Visible = true;
            foreach (Company com in companies) {
                progressBar.PerformStep();
                //sendMail(com);
                System.Threading.Thread.Sleep(2000);
                
            }

            companies.Clear();
            //txtFilePath.Text = "";
            lblCompanyCount.Text = "";
            dgvCompany.DataSource = null;
            btnSendMail.Enabled = true;
            progressBar.Visible = false;
        }

        private static void sendMail(Company company) {

            var fromAddress = new MailAddress("from@gmail.com", "From Name");
            var toAddress = new MailAddress(company.Email, "To Name");
            const string fromPassword = "fromPassword";
            const string subject = "Subject";
            const string body = "Body";

            var smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
            };
            using (var message = new MailMessage(fromAddress, toAddress)
            {
                Subject = subject,
                Body = body
            })
            {
                smtp.Send(message);
            }
        }

        private void dgvCompany_DragOver(object sender, DragEventArgs e)
        {
            var a = 10;
        }

        private void dgvCompany_DragDrop(object sender, DragEventArgs e)
        {
            var a = 10;
        }

        private void dgvCompany_DragEnter(object sender, DragEventArgs e)
        {
            var a = 10;

        }
    }
}
