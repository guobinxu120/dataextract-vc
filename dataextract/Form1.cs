using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Net;
using OpenQA.Selenium.IE;
using System.Threading;

namespace dataextract
{
    public partial class Form1 : Form
    {
        IWebDriver driver;
        Microsoft.Office.Interop.Excel.Application xlApp;
        Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
        Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet, xlWorkSheet2, xlWorkSheet3, xlWorkSheet4, xlWorkSheet5, xlWorkSheet6, xlWorkSheet7, xlWorkSheet8;

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            string str = opendialog();
            txtExcel.Text = str;
            tour(2, txtExcel.Text);
        }
        public string opendialog()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse Text Files";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.DefaultExt = "txt";
            openFileDialog1.Filter = "Excel files (*.xls)|*.xlsx|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.ReadOnlyChecked = true;
            openFileDialog1.ShowReadOnly = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtExcel.Text = openFileDialog1.FileName;
            }
            return openFileDialog1.FileName;
        }
        public Form1()
        {
            InitializeComponent();
        }


        public void tour(int startvalue, string excelpath)
        {
            string pageurl = "", TripRateFor100 = "", sheet1price = "", sheet2price = "", sheet3price = "", sheet4price = "", sheet5price = "", sheet6price = "", sheet7price = "", sheet8price = "", TripRateFor200 = "", add1 = "", add2 = "", cancel2500for100 = "", cancel5000for100 = "", cancel10000for100 = "", cancelunlimitedfor100 = "", cancel2500for200 = "", cancel5000for200 = "", cancel10000for200 = "", cancelunlimitedfor200 = "";
            pageurl = "https://www.scti.co.nz";

            try
            {
                string pagesource = "", countryname = "";
                int idx;

                string chromedriver = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
                driver = new ChromeDriver(chromedriver);
                driver.Navigate().GoToUrl(pageurl);
                Microsoft.Office.Interop.Excel.Range range, range2, range3, range4, range5, range6, range7, range8;
                string videoId;
                int rCnt;
                int cCnt;
                int rw = 0;
                int cl = 2;
                List<string> listr = new List<string>();
                List<string> invalidNumber = new List<string>();
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlApp.DisplayAlerts = false;
                xlWorkBook = xlApp.Workbooks.Open(excelpath, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
                xlWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);
                xlWorkSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(4);
                xlWorkSheet5 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(5);
                xlWorkSheet6 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(6);
                xlWorkSheet7 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(7);
                xlWorkSheet8 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(8);
                range = xlWorkSheet.UsedRange;
                range2 = xlWorkSheet2.UsedRange;
                range3 = xlWorkSheet3.UsedRange;
                range4 = xlWorkSheet4.UsedRange;
                range5 = xlWorkSheet5.UsedRange;
                range6 = xlWorkSheet6.UsedRange;
                range7 = xlWorkSheet7.UsedRange;
                range8 = xlWorkSheet8.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;

                int errordate = 1, row = 1;
                rw = range.Rows.Count;
                cl = 30;
                for (row = 1; row <= rw; row++)
                {
                    string searchcountry = "";
                    searchcountry = ((range.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range).Value2).ToString();
                    try
                    {
                        driver.FindElement(By.Id("StartDate")).Click();
                        driver.FindElement(By.XPath("/html/body/div[14]/div[1]/table/tbody/tr[2]/td[4]")).Click();
                        IWebElement age = driver.FindElement(By.Id("TravellerAges_0__Age"));
                        age.SendKeys("35");
                        driver.FindElement(By.ClassName("select2-search__field")).Click();
                        Thread.Sleep(500);
                        pagesource = driver.PageSource;
                        Thread.Sleep(500);
                        idx = pagesource.IndexOf("select2-results__options");
                        pagesource = pagesource.Substring(idx);
                        idx = pagesource.IndexOf(">" + searchcountry);
                        pagesource = pagesource.Substring(idx);
                        idx = pagesource.IndexOf("id=");
                        pagesource = pagesource.Substring(idx + 4);
                        idx = pagesource.IndexOf('"');
                        countryname = pagesource.Substring(0, idx);
                        driver.FindElement(By.Id(countryname)).Click();
                        Thread.Sleep(500);
                        try
                        {
                           
                            
                            cCnt = 2;
                            for (int j = 2; j < 7; j++)
                            {
                                for (int i = 5; i < 8; i++)
                                {
                                    driver.FindElement(By.Id("EndDate")).Click();
                                    Thread.Sleep(1000);
                                    driver.FindElement(By.XPath("/html/body/div[14]/div[1]/table/tbody/tr[" + j + "]/td[" + i + "]")).Click();
                                    IWebElement submit = driver.FindElement(By.TagName("button"));
                                    submit.Click();
                                    /////page2 
                                    Thread.Sleep(3000);
                                    pagesource = driver.PageSource;
                                    idx = pagesource.IndexOf("option-cost body-font bold");
                                    pagesource = pagesource.Substring(idx + 2);
                                    idx = pagesource.IndexOf("</h2>");
                                    TripRateFor100 = pagesource.Substring(0, idx);
                                    TripRateFor100 = TripRateFor100.Replace("<span class=" + '"' + "small" + '"' + ">", "");
                                    TripRateFor100 = TripRateFor100.Replace("</span>", "").Trim();
                                    TripRateFor100 = TripRateFor100.Replace("$", "").Trim();
                                    TripRateFor100 = TripRateFor100.Replace("tion-cost body-font bold" + '"' + ">", "").Trim();
                                    IWebElement submit2 = driver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div[1]/div[2]/form/div[1]/div/div[1]/div[2]/div/div/a"));
                                    Thread.Sleep(500);
                                    submit2.Click();
                                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(99);
                                    Thread.Sleep(3000);
                                    /////page3 
                                    pagesource = driver.PageSource;
                                    idx = pagesource.IndexOf("cancellation-cover col-xs-12 col-sm-8");
                                    pagesource = pagesource.Substring(idx);
                                    idx = pagesource.IndexOf("-cost");
                                    pagesource = pagesource.Substring(idx + 7);
                                    idx = pagesource.IndexOf('"');
                                    cancel2500for100 = pagesource.Substring(0, idx);
                                    cancel2500for100 = cancel2500for100.Replace("$", "");
                                    idx = pagesource.IndexOf("-cost");
                                    pagesource = pagesource.Substring(idx + 7);
                                    idx = pagesource.IndexOf('"');
                                    cancel5000for100 = pagesource.Substring(0, idx);
                                    cancel5000for100 = cancel5000for100.Replace("$", "");
                                    idx = pagesource.IndexOf("-cost");
                                    pagesource = pagesource.Substring(idx + 7);
                                    idx = pagesource.IndexOf('"');
                                    cancel10000for100 = pagesource.Substring(0, idx);
                                    cancel10000for100 = cancel10000for100.Replace("$", "");
                                    idx = pagesource.IndexOf("-cost");
                                    pagesource = pagesource.Substring(idx + 7);
                                    idx = pagesource.IndexOf('"');
                                    cancelunlimitedfor100 = pagesource.Substring(0, idx);
                                    cancelunlimitedfor100 = cancelunlimitedfor100.Replace("$", "");
                                    /////navigate from page3 to page2
                                    driver.Navigate().Back();
                                    /////Again page2
                                    driver.FindElement(By.ClassName("select2-selection__arrow")).Click();
                                    Thread.Sleep(1000);
                                    pagesource = driver.PageSource;
                                    idx = pagesource.IndexOf("select2-results__options");
                                    pagesource = pagesource.Substring(idx);
                                    idx = pagesource.IndexOf("$100.00");
                                    pagesource = pagesource.Substring(idx);
                                    idx = pagesource.IndexOf("id=");
                                    pagesource = pagesource.Substring(idx + 4);
                                    idx = pagesource.IndexOf('"');
                                    pagesource = pagesource.Substring(0, idx);
                                    driver.FindElement(By.Id(pagesource)).Click();
                                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(99);
                                    Thread.Sleep(2000);
                                    pagesource = driver.PageSource;
                                    idx = pagesource.IndexOf("option-cost body-font bold");
                                    pagesource = pagesource.Substring(idx + 2);
                                    idx = pagesource.IndexOf("</h2>");
                                    TripRateFor200 = pagesource.Substring(0, idx);
                                    TripRateFor200 = TripRateFor200.Replace("<span class=" + '"' + "small" + '"' + ">", "");
                                    TripRateFor200 = TripRateFor200.Replace("</span>", "").Trim();
                                    TripRateFor200 = TripRateFor200.Replace("$", "").Trim();
                                    TripRateFor200 = TripRateFor200.Replace("tion-cost body-font bold" + '"' + ">", "").Trim();
                                    submit2 = driver.FindElement(By.XPath("/html/body/div[1]/div/div[2]/div[1]/div[2]/form/div[1]/div/div[1]/div[2]/div/div/a"));
                                    submit2.Click();
                                    /////Again page3 
                                    pagesource = driver.PageSource;
                                    idx = pagesource.IndexOf("cancellation-cover col-xs-12 col-sm-8");
                                    pagesource = pagesource.Substring(idx);
                                    idx = pagesource.IndexOf("-cost");
                                    pagesource = pagesource.Substring(idx + 7);
                                    idx = pagesource.IndexOf('"');
                                    cancel2500for200 = pagesource.Substring(0, idx);
                                    cancel2500for200 = cancel2500for200.Replace("$", "");
                                    idx = pagesource.IndexOf("-cost");
                                    pagesource = pagesource.Substring(idx + 7);
                                    idx = pagesource.IndexOf('"');
                                    cancel5000for200 = pagesource.Substring(0, idx);
                                    cancel5000for200 = cancel5000for200.Replace("$", "");
                                    idx = pagesource.IndexOf("-cost");
                                    pagesource = pagesource.Substring(idx + 7);
                                    idx = pagesource.IndexOf('"');
                                    cancel10000for200 = pagesource.Substring(0, idx);
                                    cancel10000for200 = cancel10000for200.Replace("$", "");
                                    idx = pagesource.IndexOf("-cost");
                                    pagesource = pagesource.Substring(idx + 7);
                                    idx = pagesource.IndexOf('"');
                                    cancelunlimitedfor200 = pagesource.Substring(0, idx);
                                    cancelunlimitedfor200 = cancelunlimitedfor200.Replace("$", "");
                                    sheet1price = Convert.ToString(Convert.ToDecimal(TripRateFor100) + Convert.ToDecimal(cancel2500for100));
                                    sheet2price = Convert.ToString(Convert.ToDecimal(TripRateFor100) + Convert.ToDecimal(cancel5000for100));
                                    sheet3price = Convert.ToString(Convert.ToDecimal(TripRateFor100) + Convert.ToDecimal(cancel10000for100));
                                    sheet4price = Convert.ToString(Convert.ToDecimal(TripRateFor100) + Convert.ToDecimal(cancelunlimitedfor100));

                                    sheet5price = Convert.ToString(Convert.ToDecimal(TripRateFor200) + Convert.ToDecimal(cancel2500for200));
                                    sheet6price = Convert.ToString(Convert.ToDecimal(TripRateFor200) + Convert.ToDecimal(cancel5000for200));
                                    sheet7price = Convert.ToString(Convert.ToDecimal(TripRateFor200) + Convert.ToDecimal(cancel10000for200));
                                    sheet8price = Convert.ToString(Convert.ToDecimal(TripRateFor200) + Convert.ToDecimal(cancelunlimitedfor200));
                                    range.Cells[row, cCnt] = sheet1price;
                                    range2.Cells[row, cCnt] = sheet2price;
                                    range3.Cells[row, cCnt] = sheet3price;
                                    range4.Cells[row, cCnt] = sheet4price;
                                    range5.Cells[row, cCnt] = sheet5price;
                                    range6.Cells[row, cCnt] = sheet6price;
                                    range7.Cells[row, cCnt] = sheet7price;
                                    range8.Cells[row, cCnt] = sheet8price;
                                    xlWorkBook.Save();
                                    driver.Navigate().Back();
                                    driver.Navigate().Back();
                                    cCnt++;
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    catch (Exception ex)
                    {
                    }

                } 
            }
            catch (Exception ex)
            {

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
        }

    }
}
