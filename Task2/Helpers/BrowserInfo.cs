using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Chrome;
using TechTalk.SpecFlow;
using OpenQA.Selenium.Support.UI;
using System.Drawing;

namespace Task2.Helpers
{
    public class BrowserInfo
    {
        public static IWebDriver Current
        {
            get
            {
                if (!FeatureContext.Current.ContainsKey("browser"))
                {
                    IWebDriver driver = null;

                    var proxy = new Proxy { IsAutoDetect = true };
                    var capabilities = new DesiredCapabilities();
                    capabilities.SetCapability(CapabilityType.Proxy, proxy);

                    var ChromeOptions = new ChromeOptions();

                    driver = new ChromeDriver(ChromeOptions);

                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
                    FeatureContext.Current["browser"] = driver;
                }
                return (IWebDriver)FeatureContext.Current["browser"];
            }
        }
        public static IWebElement WaitForElement(By by, int timeoutInSeconds = 20)
        {
            var wait = new WebDriverWait(Current, TimeSpan.FromSeconds(timeoutInSeconds));

            var waitUntilExpectedCondition = ExpectedConditions.ElementIsVisible(by);

            wait.Message = "Could not find element " + by.ToString();

            var result = wait.Until(waitUntilExpectedCondition);

            return result;
        }
        public static T NavigateToPage<T>() where T : BasePage, new()
        {
            var page = new T();
            page.OnPageLoad();

            Current.Navigate().GoToUrl(page.URL);

            try
            {
                if (page.IsValid())
                {
                    FeatureContext.Current["currentPage"] = page;
                    return page;
                }
            }
            catch { }
            Console.WriteLine("Attempt 2 at Navigate");

            return page;
        }

        [Obsolete]
        public static T NewCurrentPage<T>() where T : BasePage, new()
        {
            var page = new T();
            page.OnPageLoad();

            FeatureContext.Current["currentPage"] = page;

            return page;
        }

        [Obsolete]
        public static T GetCurrentPage<T>() where T : BasePage
        {
            var page = FeatureContext.Current["currentPage"] as T;
            return page;
        }
        public static void NavigateBack()
        {
            Current.Navigate().Back();
        }

    }
}
