
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
            this.txtSendExcelCodition = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnKillExcelProcess = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtSendWordCodition = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.companiesLoading = new AnimOfDots.Circular();
            this.companyEmailsLoading = new AnimOfDots.Circular();
            this.btnWordonDoc = new System.Windows.Forms.Button();
            this.exportFileDialog = new System.Windows.Forms.SaveFileDialog();
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
            this.dgvEmail.Name = "dgvEmail";
            this.dgvEmail.RowHeadersWidth = 51;
            this.dgvEmail.RowTemplate.Height = 29;
            this.dgvEmail.Size = new System.Drawing.Size(273, 635);
            this.dgvEmail.TabIndex = 2;
            this.dgvEmail.DragDrop += new System.Windows.Forms.DragEventHandler(this.dgvEmail_DragDrop);
            this.dgvEmail.DragEnter += new System.Windows.Forms.DragEventHandler(this.dgvEmail_DragEnter);
            // 
            // dgvCompany
            // 
            this.dgvCompany.AllowDrop = true;
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
            this.groupBox2.Controls.Add(this.txtSendExcelCodition);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.textBox1);
            this.groupBox2.Controls.Add(this.btnKillExcelProcess);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.txtSendWordCodition);
            this.groupBox2.Location = new System.Drawing.Point(277, 13);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(972, 112);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Info";
            // 
            // txtSendExcelCodition
            // 
            this.txtSendExcelCodition.Location = new System.Drawing.Point(272, 20);
            this.txtSendExcelCodition.Name = "txtSendExcelCodition";
            this.txtSendExcelCodition.PlaceholderText = "Nhấn Enter để áp dụng điều kiện";
            this.txtSendExcelCodition.Size = new System.Drawing.Size(294, 27);
            this.txtSendExcelCodition.TabIndex = 15;
            this.txtSendExcelCodition.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSendExcelCodition_KeyPress);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(47, 23);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(160, 20);
            this.label7.TabIndex = 14;
            this.label7.Text = "Điều kiện gửi file excel";
            // 
            // textBox1
            // 
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(307, 94);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(196, 20);
            this.textBox1.TabIndex = 13;
            this.textBox1.Text = " TongSoThangNo >= 3";
            // 
            // btnKillExcelProcess
            // 
            this.btnKillExcelProcess.BackColor = System.Drawing.Color.Red;
            this.btnKillExcelProcess.ForeColor = System.Drawing.Color.White;
            this.btnKillExcelProcess.Location = new System.Drawing.Point(693, 0);
            this.btnKillExcelProcess.Name = "btnKillExcelProcess";
            this.btnKillExcelProcess.Size = new System.Drawing.Size(238, 43);
            this.btnKillExcelProcess.TabIndex = 12;
            this.btnKillExcelProcess.Text = "Tắt tất cả Excell và Word";
            this.btnKillExcelProcess.UseVisualStyleBackColor = false;
            this.btnKillExcelProcess.Click += new System.EventHandler(this.btnKillExcelProcess_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(273, 94);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(28, 20);
            this.label5.TabIndex = 3;
            this.label5.Text = "vd:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(47, 61);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(219, 20);
            this.label4.TabIndex = 2;
            this.label4.Text = "Điều kiện gửi file word đôn đốc";
            // 
            // txtSendWordCodition
            // 
            this.txtSendWordCodition.Location = new System.Drawing.Point(272, 58);
            this.txtSendWordCodition.Name = "txtSendWordCodition";
            this.txtSendWordCodition.PlaceholderText = "Nhấn Enter để áp dụng điều kiện";
            this.txtSendWordCodition.Size = new System.Drawing.Size(294, 27);
            this.txtSendWordCodition.TabIndex = 0;
            this.txtSendWordCodition.TextChanged += new System.EventHandler(this.txtSendWordCodition_TextChanged);
            this.txtSendWordCodition.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSendWordCodition_KeyPress);
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
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(344, 128);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(54, 20);
            this.label2.TabIndex = 8;
            this.label2.Text = "Report";
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
            // btnWordonDoc
            // 
            this.btnWordonDoc.AllowDrop = true;
            this.btnWordonDoc.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnWordonDoc.ForeColor = System.Drawing.Color.White;
            this.btnWordonDoc.Location = new System.Drawing.Point(12, 12);
            this.btnWordonDoc.Name = "btnWordonDoc";
            this.btnWordonDoc.Size = new System.Drawing.Size(215, 103);
            this.btnWordonDoc.TabIndex = 12;
            this.btnWordonDoc.Text = "Kéo thả file word \"Đôn đốc\" vào đây";
            this.btnWordonDoc.UseVisualStyleBackColor = false;
            this.btnWordonDoc.DragDrop += new System.Windows.Forms.DragEventHandler(this.btnWordonDoc_DragDrop);
            this.btnWordonDoc.DragEnter += new System.Windows.Forms.DragEventHandler(this.btnWordonDoc_DragEnter);
            // 
            // SendMailForm
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1400, 800);
            this.Controls.Add(this.btnWordonDoc);
            this.Controls.Add(this.companyEmailsLoading);
            this.Controls.Add(this.companiesLoading);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnSendMail);
            this.Controls.Add(this.dgvCompany);
            this.Controls.Add(this.dgvEmail);
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
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtSendWordCodition;
        private AnimOfDots.Circular companiesLoading;
        private AnimOfDots.Circular companyEmailsLoading;
        private System.Windows.Forms.Button btnKillExcelProcess;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox txtSendExcelCodition;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btnWordonDoc;
        private System.Windows.Forms.SaveFileDialog exportFileDialog;
    }
}