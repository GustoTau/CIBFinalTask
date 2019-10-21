using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using System;
using TechTalk.SpecFlow;
using System.Reflection;
using BoDi;
using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using AventStack.ExtentReports.Gherkin.Model;
using Task2.Helpers;

namespace AutomationSolution.Reporting
{
    [Binding]
    public class Hooks
    {
        private static ExtentTest _feature; // node for Feature
        private static ExtentTest _scenario; // node for Scenario
        private static ExtentReports _extent; // ExtentReports object to be created

        private static readonly string PathReport = $" { AppDomain.CurrentDomain.BaseDirectory } ExtentReport.html ";

        [BeforeTestRun]
        public static void ConfigureReport()
        {
            // Enter file path
            var reporter = new ExtentHtmlReporter(PathReport);
            reporter.Config.Theme = AventStack.ExtentReports.Reporter.Configuration.Theme.Dark;

            // instantiate the ExtentReports object
            _extent = new ExtentReports();

            // here I attach to ExtentHtmlReporter
            _extent.AttachReporter(reporter);
        }
        [AfterTestRun]
        public static void TearDownReport()
        {
            _extent.Flush();
        }
        [BeforeFeature]
        public static void BeforeFeature()
        {
            _feature = _extent.CreateTest<Feature>(FeatureContext.Current.FeatureInfo.Title);
        }
        [BeforeScenario]
        public static void BeforeScenario()
        {
            _scenario = _feature.CreateNode<Scenario>(ScenarioContext.Current.ScenarioInfo.Title);
        }
        [AfterStep]
        public void InsertReportSteps()
        {
            var stepType = ScenarioStepContext.Current.StepInfo.StepDefinitionType.ToString();
            //Handle failed steps

            if (ScenarioContext.Current.TestError == null)
            {
                if (stepType == "Given")
                    _scenario.CreateNode<Given>(ScenarioStepContext.Current.StepInfo.Text);
                else if (stepType == "When")
                    _scenario.CreateNode<When>(ScenarioStepContext.Current.StepInfo.Text);
                else if (stepType == "Then")
                    _scenario.CreateNode<Then>(ScenarioStepContext.Current.StepInfo.Text);
                else if (stepType == "And")
                    _scenario.CreateNode<And>(ScenarioStepContext.Current.StepInfo.Text);
            }
            else if (ScenarioContext.Current.TestError != null)
            {
                if (stepType == "Given")
                    _scenario.CreateNode<Given>(ScenarioStepContext.Current.StepInfo.Text).Fail(ScenarioContext.Current.TestError.InnerException);
                else if (stepType == "When")
                    _scenario.CreateNode<When>(ScenarioStepContext.Current.StepInfo.Text).Fail(ScenarioContext.Current.TestError.InnerException);
                if (stepType == "Then")
                    _scenario.CreateNode<Then>(ScenarioStepContext.Current.StepInfo.Text).Fail(ScenarioContext.Current.TestError.Message);
            }
        }
        [AfterScenario]
        public void AfterScenario()
        {
            BrowserInfo.Current.Quit();
        }
    }
}
