
namespace SendMailThue
{
    partial class GuidForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.lnkEnableLessSecureApps = new System.Windows.Forms.LinkLabel();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(29, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(201, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "Đăng nhập vào link bên dưới";
            // 
            // lnkEnableLessSecureApps
            // 
            this.lnkEnableLessSecureApps.AutoSize = true;
            this.lnkEnableLessSecureApps.Location = new System.Drawing.Point(65, 65);
            this.lnkEnableLessSecureApps.Name = "lnkEnableLessSecureApps";
            this.lnkEnableLessSecureApps.Size = new System.Drawing.Size(183, 20);
            this.lnkEnableLessSecureApps.TabIndex = 1;
            this.lnkEnableLessSecureApps.TabStop = true;
            this.lnkEnableLessSecureApps.Text = "Enable \"Less Secure Apps\"";
            this.lnkEnableLessSecureApps.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkEnableLessSecureApps_LinkClicked);
            // 
            // GuidForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(719, 539);
            this.Controls.Add(this.lnkEnableLessSecureApps);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "GuidForm";
            this.Text = "GuidForm";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.LinkLabel lnkEnableLessSecureApps;
    }
}