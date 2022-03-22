using Itenso.TimePeriod;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace SendMailThue
{
    public class CompanyUtils
    {
        private static readonly int EMPTY_ROWS_BETWEEN_DONVI = 30;
        public static List<Company> GetCompaniesFromExcelFile(string excelFile, string donDocWordFile)
        {
            List<Company> companies = new List<Company>();
            List<string> fields = new List<string>
                {
                    "8_2", // ten don vi
                    "10_2", // ma don vi
                    "-20_7", // chuyen ky sau tong no
                    "-14_1", // Kết quả đơn vị đã đóng BHXH bắt buộc cho 5 lao động đến hết tháng 12/2021						
                };
            string excelSplitDir = System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName) + @"\excels";
            string wordSplitDir = System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName) + @"\words";

            List<List<string>> groups = ExcelUtils.GetGroupValuesByGroup(excelFile, fields, EMPTY_ROWS_BETWEEN_DONVI, excelSplitDir);
            foreach (List<string> groupValues in groups)
            {
                Company company = GetCompanyFromValues(groupValues);
                if (company != null) {
                    companies.Add(company);
                }
            }

            HandleDonDocWordFile(donDocWordFile, companies, wordSplitDir);

            return companies;
        }

        private static void HandleDonDocWordFile(string donDocWordFile, List<Company> companies, string wordSplitDir)
        {
            List<List<string[]>> replaces = new List<List<string[]>>();
            DateTime now = DateTime.Now;
            foreach (Company company in companies)
            {
                List<string[]> replace = new List<string[]>();

                replace.Add(new string[] { "{day}", now.Day + "" });
                replace.Add(new string[] { "{month}", (now.Month + 1) + "" });
                replace.Add(new string[] { "{year}", now.Year + "" });


                replaces.Add(replace);
            }
            List<string> outs = WordUtils.Replace(donDocWordFile, wordSplitDir, replaces);
            for (int i = 0; i < outs.Count; i++)
            {
                companies[i].Word = outs[i];
            }
        }

        public static List<Company> GetCompaniesFromFiles(List<string> files)
        {
            List<Company> companies = new List<Company>();

            foreach(string file in files)
            {
                Company company = GetCompanyFromFile(file);
                if (company != null) {
                    companies.Add(company);
                }
            }

            return companies;
        }

        public static Company GetCompanyFromValues(List<string> values)
        {
            Int64 tongNo = 0;
            try
            {
                tongNo = Int64.Parse(values[2]);
            }
            catch (Exception ex) { }

            int tongThangNo = 0;
            string month = values[3];
            try
            {
                month = values[3].Substring(values[3].Length - 7, 7);
 
            }
            catch (Exception ex) { }

             try
            {
                string[] ms = month.Split("/");
                DateTime md = new DateTime(int.Parse(ms[1]), int.Parse(ms[0]), 1);
                DateTime now = DateTime.Now;
                DateDiff dateDiff = new DateDiff(md, now);
                tongThangNo = dateDiff.Months;
            }
            catch (Exception ex) { }

            Company company = new Company
            {
                Email = "",
                TenDonVi = values[0],
                MaDonVi = values[1],
                TongNo = tongNo,
                DaDongDenThang = month,
                TongSoThangNo = tongThangNo,
                Excel =  values[4],
                Word =  "",
                AttachExcel = false,
                AttachWord = false,
            };
            return company;

        }

        public static Company GetCompanyFromFile(string file) 
        {
            try
            {
                List<string> fields = new List<string>
                {
                    "8-2", // ten don vi
                    "10-2", // ma don vi
                    "49-7", // chuyen ky sau tong no
                    "55-1", // Kết quả đơn vị đã đóng BHXH bắt buộc cho 5 lao động đến hết tháng 12/2021						
                };
                List<string> values = ExcelUtils.GetValuesFromFile(file, fields);
                return GetCompanyFromValues(values);
                
            }
            catch (Exception ex) { }
            return null;
        }

        public static List<CompanyEmail> GetCompanyEmailsFromFile(string file)
        {
            
            List<List<string>> table = ExcelUtils.GetDataFromFile(file, new int[] { 1,2}, true);
            List<CompanyEmail> companyEmails = new List<CompanyEmail>();
            
            foreach (List<string> row in table) {
                companyEmails.Add(new CompanyEmail { MaDonVi = row[0].ToString(), Email = row[1].ToString() });
            }

            return companyEmails;
        }
    }
}
