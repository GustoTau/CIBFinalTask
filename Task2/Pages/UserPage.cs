using System;
using System.Configuration;
using Task2.Helpers;
using OpenQA.Selenium;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using OpenQA.Selenium.Support.UI;
using System.Data;
using DataTable = System.Data.DataTable;
using System.Linq;

namespace Task2.Pages
{
    public class UserPage : BasePage
    {
        CaptureScreenshot cpr = new CaptureScreenshot();

        //Initialize Page Objects
        #region PageObjects

        private readonly By FIRSTNAME = By.Name("FirstName");
        private readonly By LASTNAME = By.Name("LastName");
        private readonly By USERNAME = By.Name("UserName");
        private readonly By PASSWORD = By.Name("Password");
        private readonly By CUSTOMER_RADIO_BTN = By.Name("optionsRadios");
        private readonly By ROLE_DROPDOWN = By.Name("RoleId");
        private readonly By EMAIL = By.Name("Email");
        private readonly By CELLPHONE = By.Name("Mobilephone");
        private readonly By CLOSE_BTN = By.Name("txtBranchChannelCode");
        private readonly By SAVE_BTN = By.XPath("/html/body/div[3]/div[3]/button[2]");
        private readonly string _url = ConfigurationManager.AppSettings["EnvironmentUrl"];
        private readonly By USER_TABLE_LIST = By.XPath("/html/body/table");
        #endregion

        public override string URL
        {
            get
            {
                return _url;
            }
        }
        public override bool IsValid()
        {
            return BrowserInfo.WaitForElement(USER_TABLE_LIST, 40) != null;
        }
       public void ClickAddButton()
        {
            IList<IWebElement> rows = BrowserInfo.Current.FindElements(By.XPath("/html/body/table/thead/tr"));
            string minXpath = "/html/body/table/thead/tr[";
            string maxXpath = "]/td/button";

            for (int i = 0; i < rows.Count; i++)
            {
                string actualPath = minXpath + i + maxXpath;
                try
                {
                    IWebElement webElement = BrowserInfo.Current.FindElement(By.XPath(actualPath));

                    if (webElement.Text == "Add User")
                    {
                        webElement.Click();
                        break;
                    }
                }
                catch (NoSuchElementException)
                {
                    Console.WriteLine("Retry to find element");
                }
            }
            cpr.TakeScreenshot(BrowserInfo.Current, "Click add button", "TestCase2");
        }
        public void EnterUserDetails()
        {
            try
            {
                bool popupIsDisplayed = BrowserInfo.WaitForElement(FIRSTNAME).Displayed;
                if (popupIsDisplayed.Equals(true))
                {
                    DataTable excelData = DataReader.GetExcelTableData();
                    for (int i = 1; i < excelData.Rows.Count; i++)
                    {
                        foreach (DataRow row in excelData.Rows)
                        {
                            BrowserInfo.Current.FindElement(FIRSTNAME).Clear();
                            BrowserInfo.Current.FindElement(FIRSTNAME).SendKeys(row.ItemArray[0].ToString());
                            BrowserInfo.Current.FindElement(LASTNAME).Clear();
                            BrowserInfo.Current.FindElement(LASTNAME).SendKeys(row.ItemArray[1].ToString());
                            BrowserInfo.Current.FindElement(USERNAME).Clear();
                            BrowserInfo.Current.FindElement(USERNAME).SendKeys(row.ItemArray[2].ToString());
                            BrowserInfo.Current.FindElement(PASSWORD).Clear();
                            BrowserInfo.Current.FindElement(PASSWORD).SendKeys(row.ItemArray[3].ToString());

                            //Radio button
                            IList<IWebElement> elements = BrowserInfo.Current.FindElements(CUSTOMER_RADIO_BTN);
                            if (row.ItemArray[4].ToString() == "Company AAA")
                                elements[0].Click();
                            else if (row.ItemArray[4].ToString() == "Company BBB")
                                elements[1].Click();

                            //Drop Down Selection
                            SelectElement selectRole = new SelectElement(BrowserInfo.Current.FindElement(ROLE_DROPDOWN));
                            selectRole.SelectByText(row.ItemArray[5].ToString());
                            cpr.TakeScreenshot(BrowserInfo.Current, "Add User details", "TestCase2");

                            BrowserInfo.Current.FindElement(EMAIL).Clear();
                            BrowserInfo.Current.FindElement(EMAIL).SendKeys(row.ItemArray[6].ToString());
                            BrowserInfo.Current.FindElement(CELLPHONE).Clear();
                            BrowserInfo.Current.FindElement(CELLPHONE).SendKeys(row.ItemArray[7].ToString());
                            BrowserInfo.Current.FindElement(SAVE_BTN).Click();

                            ClickAddButton();
                        }
                    } 
                }
            }
            catch(Exception e)
            { Console.WriteLine(e.Message); }
        }
        //public void ValidateUserAdded()
        //{
        //    DataTable excelData = DataReader.GetExcelTableData();    
        //    IList<IWebElement> tableRows = BrowserInfo.Current.FindElements(By.XPath("/html/body/table/tbody/tr"));

        //    DataTable webTable = new DataTable();
        //    webTable.Rows.Add(tableRows);

        //    for (int i = 0; i < excelData.Rows.Count; i++)
        //    {
        //        foreach (DataRow item in webTable.Rows)
        //        {
        //            DataRow selectedRow = webTable.Select("First Name =" + item.ItemArray.GetValue(1)).FirstOrDefault();
        //        }
        //    }

        //}

    }
}
