using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SendMailThue
{
    public partial class GuidForm : Form
    {
        public GuidForm()
        {
            InitializeComponent();
        }

        private void lnkEnableLessSecureApps_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.lnkEnableLessSecureApps.LinkVisited = true;
            string link = "https://www.google.com/settings/security/lesssecureapps";
            ProcessStartInfo psInfo = new ProcessStartInfo
            {
                FileName = link,
                UseShellExecute = true
            };
            Process.Start(psInfo);

        }
    }
}
