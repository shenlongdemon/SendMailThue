using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Z.Expressions;
using System.Drawing;
using System.Threading.Tasks;
using System.IO;
using SendMailThue.Extentions;

namespace SendMailThue
{
    public partial class SendMailForm : Form
    {

        delegate void UpdateCompaniesDelegate();
        delegate void UpdateCompanyEmailsDelegate();
        delegate void UpdateCompaniesLoadingDelegate(bool isShow);
        delegate void UpdateCompanyEmailsLoadingDelegate(bool isShow);
        delegate void UpdateSendEmailsLoadingDelegate(bool isShow);

        private List<Company> companies = new List<Company>();
        private List<CompanyEmail> companyEmails = new List<CompanyEmail>();

        private string companyFile = "";
        private string companyEmailFile = Storage.CompanyEmailFile;
        private string donDocWordFile = FileUtils.ExecuteAppDir + @"\donDoc.doc";
        public SendMailForm()
        {
            InitializeComponent();
            companiesLoading.Visible = false;
            companyEmailsLoading.Visible = false;
            sendMailLoading.Visible = false;
            dgvCompany.DoubleBuffered(true);
        }

        private void SendMail()
        {
            if (!chkExcelEmail.Checked && !chkWordEmail.Checked)
            {
                return;
            }
            List<Company> validCompanies = companies.Where((c) => c.Email != "" && (c.AttachExcel || c.AttachWord)).ToList();
            if (validCompanies.Count == 0)
            {
                return;
            }
            showSendMailLoading(true);

            foreach (Company company in validCompanies)
            {
                string excelFile = "";
                if (company.AttachExcel)
                {
                    excelFile = CompanyUtils.GetCompanyExcelFile(company);
                    FileUtils.WaitForFile(excelFile);
                    excelFile = FileUtils.CopyFile(excelFile, FileUtils.ExcelDir + @"\Thông báo kết quả.xls");
                }
                string wordFile = "";
                if (company.AttachWord)
                {
                    wordFile = CompanyUtils.GetCompanyWordFile(company);
                    FileUtils.WaitForFile(wordFile);
                    wordFile = FileUtils.CopyFile(wordFile, FileUtils.WordDir + @"\Công văn đôn đốc.doc");
                }
                EmailUtils.SendGMail(company.Email, "hông báo kết quả đóng BHXH tháng " + DateTime.Now.ToString("MM/yyyy"), "", new List<string> { excelFile, wordFile });
            }
            showSendMailLoading(false);

        }
        private void Print()
        {
            try
            {
                if (!chkWordPrint.Checked)
                {
                    return;
                }
                List<Company> validCompanies = companies.Where((c) => c.Email != "" && c.AttachWord).ToList();
                if (validCompanies.Count == 0)
                {
                    return;
                }

                using (PrintDialog printDialog1 = new PrintDialog())
                {
                    //printDialog1.PrinterSettings.PrinterName = printer;
                    if (printDialog1.ShowDialog() == DialogResult.OK)
                    {
                        foreach (Company company in validCompanies)
                        {
                            System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo(CompanyUtils.GetCompanyWordFile(company));
                            //info.Arguments = „\““ + printDialog1.PrinterSettings.PrinterName + „\““;
                            //info.Arguments = „\““ +printer + „\““;
                            info.CreateNoWindow = true;
                            info.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                            info.UseShellExecute = true;
                            info.Verb = "PrintTo";
                            System.Diagnostics.Process.Start(info);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorUtils.ShowError(ex, true);
            }

        }

        private void LoadCompaniesFromFile()
        {
            try
            {
                showCompanyLoading(true);

                CompanyUtils.GetCompaniesFromExcelFile(companyFile, donDocWordFile, (result) =>
                {
                    try
                    {
                        companies = result;
                        showCompanyLoading(false);
                        MMapMailCompany();
                    }
                    catch (Exception ex)
                    {
                        ErrorUtils.ShowError(ex, "GetCompaniesFromExcelFile result", result);
                    }
                });
            }
            catch (Exception ex)
            {
                ErrorUtils.ShowError(ex, "LoadCompaniesFromFile");
            }
        }

        private void LoadCompanyEmailFromFile()
        {
            try
            {
                showCompanyEmailsLoading(true);
                companyEmails = CompanyUtils.GetCompanyEmailsFromFile(companyEmailFile);
                showCompanyEmailsLoading(false);
                MMapMailCompany();
            }
            catch (Exception ex)
            {
                ErrorUtils.ShowError(ex, true);
            }
        }

        private void MMapMailCompany()
        {
            if (companies.Count >= 0 && companyEmails.Count >= 0)
            {
                foreach (Company company in companies)
                {
                    CompanyEmail cm = companyEmails.Find(x => x.MaDonVi == company.MaDonVi);
                    if (cm != null)
                    {
                        company.Email = cm.Email;
                    }
                }
            }
            try
            {
                UpdateCompaniesDataSource();
                UpdateCompanyEmailsDataSource();
            }
            catch (Exception ex)
            {
                ErrorUtils.ShowError(ex, true);
            }
        }

        void showSendMailLoading(bool isShow)
        {
            if (InvokeRequired)
            {
                this.Invoke(new UpdateSendEmailsLoadingDelegate(showSendMailLoading), isShow);
                return;
            }
            this.sendMailLoading.IsEnabled = isShow;
            this.sendMailLoading.Visible = isShow;
            btnSendMail.Visible = !isShow;
        }

        void showCompanyLoading(bool isShow)
        {
            if (InvokeRequired)
            {
                this.Invoke(new UpdateCompaniesLoadingDelegate(showCompanyLoading), isShow);
                return;
            }
            this.companiesLoading.IsEnabled = isShow;
            this.companiesLoading.Visible = isShow;
        }

        void showCompanyEmailsLoading(bool isShow)
        {
            if (InvokeRequired)
            {
                this.Invoke(new UpdateCompanyEmailsLoadingDelegate(showCompanyEmailsLoading), isShow);
                return;
            }
            this.companyEmailsLoading.IsEnabled = isShow;
            this.companyEmailsLoading.Visible = isShow;
        }

        void UpdateCompaniesDataSource()
        {
            if (InvokeRequired)
            {
                this.Invoke(new UpdateCompaniesDelegate(UpdateCompaniesDataSource));
                return;
            }
            //this.dgvCompany.DataSource = new BindingSource(companies, ""); ;
            this.dgvCompany.DataSource = new BindingSource(companies, "");


            this.dgvCompany.Columns["TenDonVi"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            this.dgvCompany.Columns["TongNo"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            this.dgvCompany.Columns["TongNo"].DefaultCellStyle.Format = Utility.MoneyFormat;
            this.dgvCompany.Columns["TongNo"].DefaultCellStyle.Font = new Font("Tahoma", 12);
            this.dgvCompany.Columns["TongNo"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

            this.dgvCompany.Columns["DaDongDenThang"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.dgvCompany.Columns["TongSoThangNo"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            this.dgvCompany.Columns["TongSoThangNo"].DefaultCellStyle.Font = new Font("Tahoma", 12);

            this.dgvCompany.Columns["AttachExcel"].ReadOnly = true;
            this.dgvCompany.Columns["AttachWord"].ReadOnly = true;


            //int columnIndex = dgvCompany.Columns.Count;
            //if (dgvCompany.Columns["Export"] == null)
            //{
            //    DataGridViewButtonColumn buttonColSaveAs = new DataGridViewButtonColumn();
            //    buttonColSaveAs.Name = "Export";
            //    buttonColSaveAs.Text = "Save As";
            //    buttonColSaveAs.UseColumnTextForButtonValue = true;
            //    dgvCompany.Columns.Insert(columnIndex, buttonColSaveAs);
            //}
        }
        void UpdateCompanyEmailsDataSource()
        {
            if (InvokeRequired)
            {
                this.Invoke(new UpdateCompanyEmailsDelegate(UpdateCompanyEmailsDataSource));
                return;
            }
            BindingSource binding = new BindingSource();
            binding.DataSource = companyEmails;
            dgvEmail.DataSource = null;
            dgvEmail.DataSource = binding;
        }



        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void dgvCompany_DragDrop(object sender, DragEventArgs e)
        {
            string[] fileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (fileList.Length != 1)
            {
                return;
            }
            companyFile = fileList[0];
            Thread trd = new Thread(new ThreadStart(this.LoadCompaniesFromFile));
            trd.IsBackground = true;
            trd.Start();
        }

        private void dgvEmail_DragDrop(object sender, DragEventArgs e)
        {
            string[] fileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (fileList.Length != 1)
            {
                return;
            }
            companyEmailFile = fileList[0];
            Storage.CompanyEmailFile = companyEmailFile;
            StartLoadCompanyEmailFromFile();
        }

        void StartLoadCompanyEmailFromFile()
        {
            Thread trd = new Thread(new ThreadStart(this.LoadCompanyEmailFromFile));
            trd.IsBackground = true;
            trd.Start();
        }


        private void SendMailForm_Load(object sender, EventArgs e)
        {
            StartLoadCompanyEmailFromFile();
            LoadDonDocWordFile();
        }

        private void LoadDonDocWordFile()
        {
            if (!FileUtils.IsFileExist(donDocWordFile))
            {
                File.WriteAllBytes(donDocWordFile, Storage.DonDocWordFile);
            }
        }

        private void dgvCompany_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridViewRow oRow = dgvCompany.Rows[e.RowIndex];
            if (dgvCompany.Columns[e.ColumnIndex].Name == "AttachExcel")
            {
                bool isAttachExcel = oRow.Cells["AttachExcel"].Value != null
                                       && (bool)oRow.Cells["AttachExcel"].Value == true;

                bool isAttachWord = oRow.Cells["AttachWord"].Value != null
                                       && (bool)oRow.Cells["AttachWord"].Value == true;

                if (isAttachExcel && isAttachWord)
                {
                    oRow.DefaultCellStyle.BackColor = Color.FromArgb(255, 179, 179);
                }
                else if (isAttachExcel)
                {
                    oRow.DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204);
                }
                else if (isAttachWord)
                {
                    oRow.DefaultCellStyle.BackColor = Color.FromArgb(179, 217, 255);
                }
            }
        }

        private void ApplyCondition()
        {
            int numberForExcel = 0;
            int numberForWord = 0;

            try
            {
                numberForExcel = int.Parse(txtSendExcelCodition.Text.Trim());
            }
            catch (Exception ex) { }
            try
            {
                numberForWord = int.Parse(txtSendWordCodition.Text.Trim());
            }
            catch (Exception ex) { }

            foreach (Company company in companies)
            {
                if (company.Email.Trim() == "" && numberForExcel == 0 && numberForWord == 0)
                {
                    company.AttachExcel = false;
                    company.AttachWord = false;
                    continue;
                }

                company.AttachExcel = numberForExcel > 0 && company.TongSoThangNo >= numberForExcel;

                company.AttachWord = numberForWord > 0 && company.TongSoThangNo >= numberForWord;
            }
            UpdateCompaniesDataSource();
        }

        #region ignore

        private void btnWordonDoc_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void companiesLoading_Load(object sender, EventArgs e)
        {

        }

        private async Task GoogleLogin()
        {
            //string token = await GoogleAuth.Login(this);
            //return token;
            EmailUtils.SendGMail("shenlongdemon@gmail.com", "subject", "body", new List<string> { @"C:\Users\ASUS\Downloads\donDoc.doc" });
        }

        private void txtSendExcelCodition_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                ApplyCondition();
            }
        }

        private void txtSendWordCodition_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                ApplyCondition();
            }
        }

        private void dgvCompany_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void dgvEmail_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void btnKillExcelProcess_Click(object sender, EventArgs e)
        {
            var excelProcesses = from p in Process.GetProcessesByName("EXCEL")
                                 select p;

            foreach (var process in excelProcesses)
            {
                process.Kill();
            }

            var wordProcesses = from p in Process.GetProcessesByName("WINWORD")
                                select p;

            foreach (var process in wordProcesses)
            {
                process.Kill();
            }
        }

        private void txtSendWordCodition_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnSendMail_Click(object sender, EventArgs e)
        {
            Thread sendMailTask = new Thread(new ThreadStart(this.SendMail));
            sendMailTask.IsBackground = true;
            sendMailTask.Start();

            Thread printTask = new Thread(new ThreadStart(this.Print));
            printTask.IsBackground = true;
            printTask.Start();
        }

        private void dgvCompany_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
        }

        private void dgvCompany_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.ColumnIndex == dgvCompany.Columns["Export"].Index)
            //{
            //    var row = dgvCompany.Rows[e.RowIndex];
            //    string MaDonVi = row.Cells["MaDonVi"].Value.ToString();
            //    string TenDonVi = row.Cells["TenDonVi"].Value.ToString();
            //    string range = row.Cells["Range"].Value.ToString();
            //    exportFileDialog.FileName = MaDonVi + "-" + TenDonVi;
            //    if (exportFileDialog.ShowDialog() == DialogResult.OK)
            //    {
            //        ExportFiles(range, exportFileDialog.FileName);
            //    }
            //}
        }

        private void ExportFiles(string fromFileName, string savePath)
        {
            FileUtils.CopyFile(FileUtils.ExcelDir + @"\" + fromFileName + ".xls", savePath + ".xls");
            FileUtils.CopyFile(FileUtils.WordDir + @"\" + fromFileName + ".doc", savePath + ".doc");

        }

        #endregion

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void dgvEmail_CellEnter(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void StartUpdateCompanyEmail()
        {
            Thread trd = new Thread(new ThreadStart(this.UpdateCompanyEmails));
            trd.IsBackground = true;
            trd.Start();
        }

        private void UpdateCompanyEmails()
        {
            try
            {
                companyEmails = (dgvEmail.DataSource as BindingSource).DataSource as List<CompanyEmail>;
                MMapMailCompany();
                CompanyUtils.SaveCompanyEmailsToExcelFile(companyEmailFile, companyEmails);
            }
            catch (Exception ex)
            {
                ErrorUtils.ShowError(ex, true);
            }
        }

        private void dgvEmail_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                StartUpdateCompanyEmail();
            }
        }

        private void dgvEmail_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                StartUpdateCompanyEmail();
            }
            catch (Exception ex) { }
        }
    }
}
