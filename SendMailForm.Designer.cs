
namespace SendMailThue
{
    partial class SendMailForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SendMailForm));
            this.dgvEmail = new System.Windows.Forms.DataGridView();
            this.dgvCompany = new System.Windows.Forms.DataGridView();
            this.btnSendMail = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.line1 = new Unclassified.UI.Line();
            this.line4 = new Unclassified.UI.Line();
            this.line3 = new Unclassified.UI.Line();
            this.line2 = new Unclassified.UI.Line();
            this.chkWordPrint = new System.Windows.Forms.CheckBox();
            this.chkExcelPrint = new System.Windows.Forms.CheckBox();
            this.chkWordEmail = new System.Windows.Forms.CheckBox();
            this.chkExcelEmail = new System.Windows.Forms.CheckBox();
            this.txtSendWordCodition = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtSendExcelCodition = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnKillExcelProcess = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.companiesLoading = new AnimOfDots.Circular();
            this.companyEmailsLoading = new AnimOfDots.Circular();
            this.exportFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.sendMailLoading = new AnimOfDots.Circular();
            ((System.ComponentModel.ISupportInitialize)(this.dgvEmail)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCompany)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgvEmail
            // 
            this.dgvEmail.AllowDrop = true;
            this.dgvEmail.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dgvEmail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvEmail.Location = new System.Drawing.Point(12, 153);
            this.dgvEmail.MultiSelect = false;
            this.dgvEmail.Name = "dgvEmail";
            this.dgvEmail.RowHeadersWidth = 51;
            this.dgvEmail.RowTemplate.Height = 29;
            this.dgvEmail.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvEmail.Size = new System.Drawing.Size(273, 635);
            this.dgvEmail.TabIndex = 2;
            this.dgvEmail.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvEmail_CellEnter);
            this.dgvEmail.UserDeletedRow += new System.Windows.Forms.DataGridViewRowEventHandler(this.dgvEmail_UserDeletedRow);
            this.dgvEmail.DragDrop += new System.Windows.Forms.DragEventHandler(this.dgvEmail_DragDrop);
            this.dgvEmail.DragEnter += new System.Windows.Forms.DragEventHandler(this.dgvEmail_DragEnter);
            this.dgvEmail.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.dgvEmail_PreviewKeyDown);
            // 
            // dgvCompany
            // 
            this.dgvCompany.AllowDrop = true;
            this.dgvCompany.AllowUserToAddRows = false;
            this.dgvCompany.AllowUserToDeleteRows = false;
            this.dgvCompany.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvCompany.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvCompany.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dgvCompany.ColumnHeadersHeight = 29;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvCompany.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgvCompany.Location = new System.Drawing.Point(291, 153);
            this.dgvCompany.Name = "dgvCompany";
            this.dgvCompany.RowHeadersWidth = 51;
            this.dgvCompany.RowTemplate.Height = 29;
            this.dgvCompany.Size = new System.Drawing.Size(1097, 635);
            this.dgvCompany.TabIndex = 3;
            this.dgvCompany.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvCompany_CellClick);
            this.dgvCompany.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvCompany_CellFormatting);
            this.dgvCompany.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.dgvCompany_DataBindingComplete);
            this.dgvCompany.DragDrop += new System.Windows.Forms.DragEventHandler(this.dgvCompany_DragDrop);
            this.dgvCompany.DragEnter += new System.Windows.Forms.DragEventHandler(this.dgvCompany_DragEnter);
            // 
            // btnSendMail
            // 
            this.btnSendMail.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnSendMail.BackgroundImage")));
            this.btnSendMail.Location = new System.Drawing.Point(1255, 13);
            this.btnSendMail.Name = "btnSendMail";
            this.btnSendMail.Size = new System.Drawing.Size(133, 133);
            this.btnSendMail.TabIndex = 4;
            this.btnSendMail.UseVisualStyleBackColor = true;
            this.btnSendMail.Click += new System.EventHandler(this.btnSendMail_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.line1);
            this.groupBox2.Controls.Add(this.line4);
            this.groupBox2.Controls.Add(this.line3);
            this.groupBox2.Controls.Add(this.line2);
            this.groupBox2.Controls.Add(this.chkWordPrint);
            this.groupBox2.Controls.Add(this.chkExcelPrint);
            this.groupBox2.Controls.Add(this.chkWordEmail);
            this.groupBox2.Controls.Add(this.chkExcelEmail);
            this.groupBox2.Controls.Add(this.txtSendWordCodition);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.txtSendExcelCodition);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.groupBox2.Location = new System.Drawing.Point(291, 13);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(958, 134);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Info";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label5.Location = new System.Drawing.Point(694, 9);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(43, 20);
            this.label5.TabIndex = 27;
            this.label5.Text = "Print";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label2.Location = new System.Drawing.Point(617, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 20);
            this.label2.TabIndex = 26;
            this.label2.Text = "Email";
            // 
            // line1
            // 
            this.line1.BorderColor = System.Drawing.Color.Green;
            this.line1.Dark3dColor = System.Drawing.SystemColors.ControlDark;
            this.line1.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            this.line1.Light3dColor = System.Drawing.SystemColors.ControlLightLight;
            this.line1.Location = new System.Drawing.Point(445, 53);
            this.line1.Name = "line1";
            this.line1.Orientation = Unclassified.UI.LineOrientation.DiagonalUp;
            this.line1.Size = new System.Drawing.Size(68, 24);
            this.line1.TabIndex = 25;
            this.line1.TabStop = false;
            // 
            // line4
            // 
            this.line4.BorderColor = System.Drawing.Color.Navy;
            this.line4.Dark3dColor = System.Drawing.SystemColors.ControlDark;
            this.line4.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            this.line4.Light3dColor = System.Drawing.SystemColors.ControlLightLight;
            this.line4.Location = new System.Drawing.Point(446, 77);
            this.line4.Name = "line4";
            this.line4.Orientation = Unclassified.UI.LineOrientation.DiagonalDown;
            this.line4.Size = new System.Drawing.Size(68, 24);
            this.line4.TabIndex = 24;
            this.line4.TabStop = false;
            // 
            // line3
            // 
            this.line3.BorderColor = System.Drawing.Color.Green;
            this.line3.Dark3dColor = System.Drawing.SystemColors.ControlDark;
            this.line3.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            this.line3.Light3dColor = System.Drawing.SystemColors.ControlLightLight;
            this.line3.Location = new System.Drawing.Point(210, 54);
            this.line3.Name = "line3";
            this.line3.Orientation = Unclassified.UI.LineOrientation.DiagonalDown;
            this.line3.Size = new System.Drawing.Size(68, 24);
            this.line3.TabIndex = 23;
            this.line3.TabStop = false;
            // 
            // line2
            // 
            this.line2.BorderColor = System.Drawing.Color.Navy;
            this.line2.Dark3dColor = System.Drawing.SystemColors.ControlDark;
            this.line2.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            this.line2.Light3dColor = System.Drawing.SystemColors.ControlLightLight;
            this.line2.Location = new System.Drawing.Point(211, 78);
            this.line2.Name = "line2";
            this.line2.Orientation = Unclassified.UI.LineOrientation.DiagonalUp;
            this.line2.Size = new System.Drawing.Size(68, 24);
            this.line2.TabIndex = 22;
            this.line2.TabStop = false;
            // 
            // chkWordPrint
            // 
            this.chkWordPrint.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.chkWordPrint.Location = new System.Drawing.Point(698, 93);
            this.chkWordPrint.Name = "chkWordPrint";
            this.chkWordPrint.Size = new System.Drawing.Size(30, 30);
            this.chkWordPrint.TabIndex = 20;
            this.chkWordPrint.UseVisualStyleBackColor = true;
            // 
            // chkExcelPrint
            // 
            this.chkExcelPrint.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.chkExcelPrint.Location = new System.Drawing.Point(698, 46);
            this.chkExcelPrint.Name = "chkExcelPrint";
            this.chkExcelPrint.Size = new System.Drawing.Size(30, 30);
            this.chkExcelPrint.TabIndex = 19;
            this.chkExcelPrint.UseVisualStyleBackColor = true;
            // 
            // chkWordEmail
            // 
            this.chkWordEmail.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.chkWordEmail.Location = new System.Drawing.Point(626, 93);
            this.chkWordEmail.Name = "chkWordEmail";
            this.chkWordEmail.Size = new System.Drawing.Size(30, 30);
            this.chkWordEmail.TabIndex = 18;
            this.chkWordEmail.UseVisualStyleBackColor = true;
            // 
            // chkExcelEmail
            // 
            this.chkExcelEmail.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.chkExcelEmail.Location = new System.Drawing.Point(626, 46);
            this.chkExcelEmail.Name = "chkExcelEmail";
            this.chkExcelEmail.Size = new System.Drawing.Size(30, 30);
            this.chkExcelEmail.TabIndex = 13;
            this.chkExcelEmail.UseVisualStyleBackColor = true;
            // 
            // txtSendWordCodition
            // 
            this.txtSendWordCodition.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.txtSendWordCodition.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.txtSendWordCodition.Location = new System.Drawing.Point(530, 83);
            this.txtSendWordCodition.Name = "txtSendWordCodition";
            this.txtSendWordCodition.Size = new System.Drawing.Size(45, 39);
            this.txtSendWordCodition.TabIndex = 17;
            this.txtSendWordCodition.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtSendWordCodition.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSendWordCodition_KeyPress);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label3.Location = new System.Drawing.Point(285, 66);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(158, 20);
            this.label3.TabIndex = 16;
            this.label3.Text = "Tổng số tháng nợ >=";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // txtSendExcelCodition
            // 
            this.txtSendExcelCodition.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.txtSendExcelCodition.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.txtSendExcelCodition.Location = new System.Drawing.Point(530, 34);
            this.txtSendExcelCodition.Name = "txtSendExcelCodition";
            this.txtSendExcelCodition.Size = new System.Drawing.Size(45, 39);
            this.txtSendExcelCodition.TabIndex = 15;
            this.txtSendExcelCodition.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtSendExcelCodition.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSendExcelCodition_KeyPress);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label7.ForeColor = System.Drawing.Color.Green;
            this.label7.Location = new System.Drawing.Point(23, 41);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(180, 20);
            this.label7.TabIndex = 14;
            this.label7.Text = "Thông báo kết quả đóng";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label4.ForeColor = System.Drawing.Color.Navy;
            this.label4.Location = new System.Drawing.Point(42, 89);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(156, 20);
            this.label4.TabIndex = 2;
            this.label4.Text = "Công văn đôn đốc nợ";
            // 
            // btnKillExcelProcess
            // 
            this.btnKillExcelProcess.BackColor = System.Drawing.Color.Red;
            this.btnKillExcelProcess.ForeColor = System.Drawing.Color.White;
            this.btnKillExcelProcess.Location = new System.Drawing.Point(12, 12);
            this.btnKillExcelProcess.Name = "btnKillExcelProcess";
            this.btnKillExcelProcess.Size = new System.Drawing.Size(117, 102);
            this.btnKillExcelProcess.TabIndex = 12;
            this.btnKillExcelProcess.Text = "Tắt tất cả Excell và Word";
            this.btnKillExcelProcess.UseVisualStyleBackColor = false;
            this.btnKillExcelProcess.Click += new System.EventHandler(this.btnKillExcelProcess_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 130);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(274, 20);
            this.label1.TabIndex = 7;
            this.label1.Text = "Kéo thả file \"Email\" vào khung bên dưới";
            // 
            // companiesLoading
            // 
            this.companiesLoading.AnimationSpeed = ((byte)(60));
            this.companiesLoading.BackColor = System.Drawing.Color.White;
            this.companiesLoading.IsEnabled = false;
            this.companiesLoading.Location = new System.Drawing.Point(801, 394);
            this.companiesLoading.Name = "companiesLoading";
            this.companiesLoading.Size = new System.Drawing.Size(155, 118);
            this.companiesLoading.TabIndex = 9;
            this.companiesLoading.Visible = false;
            this.companiesLoading.Load += new System.EventHandler(this.companiesLoading_Load);
            // 
            // companyEmailsLoading
            // 
            this.companyEmailsLoading.AnimationSpeed = ((byte)(60));
            this.companyEmailsLoading.BackColor = System.Drawing.Color.White;
            this.companyEmailsLoading.IsEnabled = false;
            this.companyEmailsLoading.Location = new System.Drawing.Point(45, 394);
            this.companyEmailsLoading.Name = "companyEmailsLoading";
            this.companyEmailsLoading.Size = new System.Drawing.Size(155, 118);
            this.companyEmailsLoading.TabIndex = 10;
            this.companyEmailsLoading.Visible = false;
            // 
            // sendMailLoading
            // 
            this.sendMailLoading.AnimationSpeed = ((byte)(60));
            this.sendMailLoading.BackColor = System.Drawing.Color.Transparent;
            this.sendMailLoading.IsEnabled = false;
            this.sendMailLoading.Location = new System.Drawing.Point(1255, 13);
            this.sendMailLoading.Name = "sendMailLoading";
            this.sendMailLoading.Size = new System.Drawing.Size(133, 133);
            this.sendMailLoading.TabIndex = 13;
            this.sendMailLoading.Visible = false;
            // 
            // SendMailForm
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1400, 800);
            this.Controls.Add(this.sendMailLoading);
            this.Controls.Add(this.companyEmailsLoading);
            this.Controls.Add(this.companiesLoading);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnSendMail);
            this.Controls.Add(this.dgvCompany);
            this.Controls.Add(this.dgvEmail);
            this.Controls.Add(this.btnKillExcelProcess);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SendMailForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SendMail";
            this.Load += new System.EventHandler(this.SendMailForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvEmail)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvCompany)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.DataGridView dgvEmail;
        private System.Windows.Forms.DataGridView dgvCompany;
        private System.Windows.Forms.Button btnSendMail;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private AnimOfDots.Circular companiesLoading;
        private AnimOfDots.Circular companyEmailsLoading;
        private System.Windows.Forms.Button btnKillExcelProcess;
        private System.Windows.Forms.TextBox txtSendExcelCodition;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.SaveFileDialog exportFileDialog;
        private System.Windows.Forms.TextBox txtSendWordCodition;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox chkWordPrint;
        private System.Windows.Forms.CheckBox chkExcelPrint;
        private System.Windows.Forms.CheckBox chkWordEmail;
        private System.Windows.Forms.CheckBox chkExcelEmail;
        private Unclassified.UI.Line line2;
        private Unclassified.UI.Line line1;
        private Unclassified.UI.Line line4;
        private Unclassified.UI.Line line3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label2;
        private AnimOfDots.Circular sendMailLoading;
    }
}