using OpenQA.Selenium;
using System;

namespace Task2.Helpers
{
    public abstract class BasePage : BroswerablePage
    {
        public string Title
        {
            get
            {
                return BrowserInfo.Current.Title;
            }
        }
        public virtual void OnPageLoad()
        {
            BrowserInfo.Current.Manage().Window.Maximize();
            (BrowserInfo.Current as IJavaScriptExecutor).ExecuteScript("if (typeof jQuery !== 'undefined') $.fx.off = true;");
        }
        public abstract bool IsValid();
    }
}
