using System;
using System.Collections.Generic;
using System.Text;

namespace SendMailThue
{
    public class Storage
    {
        public static String CompanyEmailFile
        {
            get {
                return Settings.Default.CompanyEmailFile;
            }

            set {
                Settings.Default.CompanyEmailFile = value;
                Settings.Default.Save();
            }
        }
        public static String DonDocWordFile
        {
            get {
                return Settings.Default.DonDocWordFile;
            }

            set {
                Settings.Default.DonDocWordFile = value;
                Settings.Default.Save();
            }
        }
    }
}
