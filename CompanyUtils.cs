using Itenso.TimePeriod;
using System;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Threading;

namespace SendMailThue
{
    public class CompanyUtils
    {
        private static readonly int EMPTY_ROWS_BETWEEN_DONVI = 30;
        public static void GetCompaniesFromExcelFile(string excelFile, string donDocWordFile, Action<List<Company>> callback)
        {
            List<string> fields = new List<string>
                {
                    "8_2", // ten don vi
                    "10_2", // ma don vi
                    "-20_7", // chuyen ky sau tong no
                    "-14_1", // Kết quả đơn vị đã đóng BHXH bắt buộc cho 5 lao động đến hết tháng 12/2021						
                };
            string excelSplitDir = FileUtils.ExcelDir;
            string wordSplitDir = FileUtils.WordDir;

            ExcelUtils.GetGroupValuesByGroup(excelFile, fields, EMPTY_ROWS_BETWEEN_DONVI, excelSplitDir, (result) =>
            {
                List<Company> companies = new List<Company>();
                List<List<string>> errorGroupValues = new List<List<string>>();
                foreach (List<string> groupValues in result)
                {
                    List<string> processedGroup = groupValues.Select(p => p.Trim().Replace("\r\n", "").Replace("\n", "").Replace("\r", "").Replace(System.Environment.NewLine, "")).ToList();
                    try
                    {
                        Company company = GetCompanyFromValues(processedGroup);
                        if (company != null)
                        {
                            companies.Add(company);
                        }
                    }
                    catch (Exception ex) 
                    {
                        errorGroupValues.Add(processedGroup);
                    }
                }
                if (errorGroupValues.Count > 0) {
                    ErrorUtils.ShowError(new Exception("Cannot parse company"), "", errorGroupValues);
                }
                callback(companies);
                var t = new Thread(() =>
                {
                    
                    try
                    {
                        HandleDonDocWordFile(donDocWordFile, companies, wordSplitDir);
                    }
                    catch (Exception ex)
                    {
                        ErrorUtils.ShowError(ex, "GetCompaniesFromExcelFile HandleDonDocWordFile", companies);
                    }
                });
                t.IsBackground = true;
                t.Start();
            });
        }

        public static string GetCompanyExcelFile(Company company)
        {
            return FileUtils.ExcelDir + @"\" + company.Range + ".xls";
        }

        public static string GetCompanyWordFile(Company company)
        {
            return FileUtils.WordDir + @"\" + company.Range + ".doc";
        }

        private static void HandleDonDocWordFile(string donDocWordFile, List<Company> companies, string wordSplitDir)
        {
            bool isShowError = false;
            List<List<string[]>> replaces = new List<List<string[]>>();
            DateTime now = DateTime.Now;
            DateTime lastDateOfMonth = new DateTime(now.Year, now.Month, DateTime.DaysInMonth(now.Year, now.Month));
            foreach (Company company in companies)
            {
                try
                {
                    List<string[]> replace = new List<string[]>();
                    replace.Add(new string[] { "{fileName}", company.Range + "" });
                    replace.Add(new string[] { "{day}", now.Day + "" });
                    replace.Add(new string[] { "{month}", (now.Month) + "" });
                    replace.Add(new string[] { "{year}", now.Year + "" });
                    replace.Add(new string[] { "{tendonvi}", company.TenDonVi + "" });
                    replace.Add(new string[] { "{now}", now.ToString("dd/MM/yyyy") });
                    replace.Add(new string[] { "{tongon}", Utility.FormatMoney(company.TongNo) });
                    replace.Add(new string[] { "{tongnobangchu}", Utility.ConvertMoneyToString(company.TongNo) });
                    replace.Add(new string[] { "{tongthangno}", company.TongSoThangNo + "" });
                    replace.Add(new string[] { "{lastdateofmonth}", lastDateOfMonth.ToString("dd/MM/yyyy") + "" });
                    replaces.Add(replace);
                }
                catch (Exception ex)
                {
                    if (!isShowError)
                    {
                        ErrorUtils.ShowError(ex, "HandleDonDocWordFile", company);
                    }
                    isShowError = true;
                }
            }
            WordUtils.Replace(donDocWordFile, wordSplitDir, replaces);
        }

        //public static List<Company> GetCompaniesFromFiles(List<string> files)
        //{
        //    List<Company> companies = new List<Company>();

        //    foreach (string file in files)
        //    {
        //        Company company = GetCompanyFromFile(file);
        //        if (company != null)
        //        {
        //            companies.Add(company);
        //        }
        //    }

        //    return companies;
        //}

        public static Company GetCompanyFromValues(List<string> values)
        {

            Int64 tongNo = Int64.Parse(values[2]);

            string month = values[3].Substring(values[3].Length - 7, 7);

            string[] ms = month.Split("/");
            DateTime md = new DateTime(int.Parse(ms[1]), int.Parse(ms[0]), 1);
            DateTime now = DateTime.Now;
            DateDiff dateDiff = new DateDiff(md, now);
            int tongThangNo = dateDiff.Months;

            Company company = new Company
            {
                Email = "",
                TenDonVi = values[0],
                MaDonVi = values[1],
                TongNo = tongNo,
                DaDongDenThang = month,
                TongSoThangNo = tongThangNo,
                AttachExcel = false,
                AttachWord = false,
                Range = values[4],
            };
            return company;

        }

        //public static Company GetCompanyFromFile(string file)
        //{
        //    try
        //    {
        //        List<string> fields = new List<string>
        //        {
        //            "8-2", // ten don vi
        //            "10-2", // ma don vi
        //            "49-7", // chuyen ky sau tong no
        //            "55-1", // Kết quả đơn vị đã đóng BHXH bắt buộc cho 5 lao động đến hết tháng 12/2021						
        //        };
        //        List<string> values = ExcelUtils.GetValuesFromFile(file, fields);
        //        return GetCompanyFromValues(values);

        //    }
        //    catch (Exception ex) { }
        //    return null;
        //}

        public static List<CompanyEmail> GetCompanyEmailsFromFile(string file)
        {

            List<List<string>> table = ExcelUtils.GetDataFromFile(file, new int[] { 1, 2 }, true);
            List<CompanyEmail> companyEmails = new List<CompanyEmail>();

            foreach (List<string> row in table)
            {
                companyEmails.Add(new CompanyEmail { MaDonVi = row[0].ToString(), Email = row[1].ToString() });
            }

            return companyEmails;
        }

        public static void SaveCompanyEmailsToExcelFile(string companyEmailFile, List<CompanyEmail> companyEmails)
        {
            List<string[]> values = companyEmails.Select(p => new string[] { p.MaDonVi, p.Email }).ToList();
            ExcelUtils.SaveValuesToFile(companyEmailFile, values, true);
        }
    }
}