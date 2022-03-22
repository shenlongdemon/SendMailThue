using System;
using System.Collections.Generic;
using System.Text;

namespace SendMailThue
{
    public class Company
    {
        public String MaDonVi { get; set; }
        public String TenDonVi { get; set; }
        public Int64 TongNo { get; set; }
        public String DaDongDenThang { get; set; }
        public int TongSoThangNo { get; set; }
        public String Email { get; set; }

        public String Excel { get; set; }
        public String Word { get; set; }

        public bool AttachWord { get; set; }
        public bool AttachExcel { get; set; }
    }
}
