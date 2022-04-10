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

        public static byte[] DonDocWordFile
        {
            get {
                return Properties.Resources.donDoc;
            }
        }


    }
}
