using System;
using TechTalk.SpecFlow;
using Task2.Helpers;
using Task2.Pages;
using FluentAssertions;
using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Gherkin.Model;

namespace Task2.Step_Definitions
{
    [Binding]
    public class TestCaseTwoSteps
    {
        [Given(@"User has navigated to (.*)")]
        public void GivenUserHasNavigatedTo(string p0)
        {
            BrowserInfo.NavigateToPage<UserPage>();
        }
        
        [Then(@"Validate that user is on User List Table")]
        public void ThenValidateThatUserIsOnUserListTable()
        {
            var validateUserTable = BrowserInfo.NewCurrentPage<UserPage>().IsValid();
            validateUserTable.Should().BeTrue();
        }

        [Then(@"Click add user")]
        public void ThenClickAddUser()
        {
            BrowserInfo.GetCurrentPage<UserPage>().ClickAddButton();
        }
        [Then(@"Add User with details from excel")]
        public void ThenAddUserWithDetailsFromExcel()
        {
            BrowserInfo.GetCurrentPage<UserPage>().EnterUserDetails();
        }

    }
}
