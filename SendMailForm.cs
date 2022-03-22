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

namespace SendMailThue
{
    public partial class SendMailForm : Form
    {
        
        delegate void UpdateCompaniesDelegate();
        delegate void UpdateCompanyEmailsDelegate();
        delegate void UpdateCompaniesLoadingDelegate(bool isShow);
        delegate void UpdateCompanyEmailsLoadingDelegate(bool isShow);

        private List<Company> companies = new List<Company>();
        private List<CompanyEmail> companyEmails = new List<CompanyEmail>();

        private string companyFile = "";
        private string companyEmailFile = Storage.CompanyEmailFile;
        private string donDocWordFile = Storage.DonDocWordFile;
        public SendMailForm()
        {
            InitializeComponent();
            companiesLoading.Visible = false;
            companyEmailsLoading.Visible = false;
            LoadCompanyEmailFromFile();
            UpdateForDonDocWordFileUI();
        }

        private void LoadCompaniesFromFile() {
            showCompanyLoading(true);
            companies = CompanyUtils.GetCompaniesFromExcelFile(companyFile, donDocWordFile);
            showCompanyLoading(false);
            MMapMailCompany();
        }

   

        private void LoadCompanyEmailFromFile()
        {
            showCompanyEmailsLoading(true);
            companyEmails = CompanyUtils.GetCompanyEmailsFromFile(companyEmailFile);
            showCompanyEmailsLoading(false);
            MMapMailCompany();
        
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
            UpdateCompaniesDataSource();
            UpdateCompanyEmailsDataSource();
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

        void UpdateForDonDocWordFileUI()
        {
            bool exits =  File.Exists(donDocWordFile);
            if (exits)
            {
                btnWordonDoc.BackColor = Color.DodgerBlue;
                btnWordonDoc.ForeColor = Color.White;
                btnWordonDoc.Text = Path.GetFileName(donDocWordFile);
            }
            else {
                btnWordonDoc.BackColor = Color.Red;
                btnWordonDoc.ForeColor = Color.White;
                btnWordonDoc.Text = "Kéo thả file \"Đôn đốc\" vào đây";
            }
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
            this.dgvCompany.Columns["TongNo"].DefaultCellStyle.Format = "#,##0";
            this.dgvCompany.Columns["DaDongDenThang"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            this.dgvCompany.Columns["TongSoThangNo"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        } 
        void UpdateCompanyEmailsDataSource()
        {
            if (InvokeRequired)
            {
                this.Invoke(new UpdateCompanyEmailsDelegate(UpdateCompanyEmailsDataSource));
                return;
            }
            this.dgvEmail.DataSource = new BindingSource(companyEmails, ""); ;
        }

       

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void dgvCompany_DragDrop(object sender, DragEventArgs e)
        {
            string[] fileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (fileList.Length != 1) {
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
            Thread trd = new Thread(new ThreadStart(this.LoadCompanyEmailFromFile));
            trd.IsBackground = true;
            trd.Start();
        }

        private void dgvCompany_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void dgvEmail_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void SendMailForm_Load(object sender, EventArgs e)
        {

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
                else if (isAttachExcel) {
                    oRow.DefaultCellStyle.BackColor = Color.FromArgb(204, 255, 204);
                }
                else if (isAttachWord) {
                    oRow.DefaultCellStyle.BackColor = Color.FromArgb(179, 217, 255);
                }
                
            }

       
        }

        

        private void btnSendMail_Click(object sender, EventArgs e)
        {

        }

        private void btnOpenGuild_Click(object sender, EventArgs e)
        {
            //GuidForm gf = new GuidForm();
            //gf.Show();
            GoogleLogin();
        }

        private async Task<string> GoogleLogin() 
        {
            string token = await GoogleAuth.Login(this);
            return token;
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

        private void ApplyCondition()
        {
            foreach (Company company in companies)
            {
                bool attachExcel = false;
                try
                {
                    attachExcel = Eval.Execute<bool>("company." + txtSendExcelCodition.Text, new { company = company });
                }
                catch (NullReferenceException nrex) { } 
                catch (Exception ex) { }
                company.AttachExcel = attachExcel; 

                bool attachWord = false;
                try
                {
                    attachWord = Eval.Execute<bool>("company." + txtSendWordCodition.Text, new { company = company });
                }
                catch (NullReferenceException nrex) { }
                catch (Exception ex) { }
                company.AttachWord = attachWord;
            }
            UpdateCompaniesDataSource();
        }

        private void btnWordonDoc_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void btnWordonDoc_DragDrop(object sender, DragEventArgs e)
        {
            string[] fileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (fileList.Length != 1)
            {
                return;
            }
            donDocWordFile = fileList[0];
            Storage.DonDocWordFile = donDocWordFile;
            UpdateForDonDocWordFileUI();
        }
    }
}
