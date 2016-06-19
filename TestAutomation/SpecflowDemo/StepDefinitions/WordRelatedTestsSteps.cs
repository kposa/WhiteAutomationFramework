using System;
using System.Collections.Generic;
using Microsoft.Win32;
using NUnit.Framework;
using TechTalk.SpecFlow;
using TestStack.White.UIItems;
using TestStack.White.UIItems.WindowItems;
using SpecflowDemo.Helpers;
using SpecflowDemo.TestDataSet;
using SpecflowDemo.GuiMaps;

namespace SpecflowDemo.StepDefinitions
{
    [Binding]
    public class WordRelatedTestsSteps
    {
        private readonly StepHelpers StepHelpers = new StepHelpers();

        [Given(@"I have installed Word Add-in")]
        public void GivenIHaveInstalledWordAddIn()
        {
            RegistryKey registryKey = StepHelpers.GetAddInRegistryKey;
            Assert.IsNotNull(registryKey);
        }

        [When(@"I launched new MS Word document")]
        public void WhenILaunchedNewMicrosoftWordDocument()
        {
            Assert.IsNotNull(BaseSteps.CurrentWindowUnderTest);
        }

        [Then(@"I should see Word menu")]
        public void ThenIShouldSeeWordMenu()
        {
            Actions.Exists("Tab_WordMenu_Txt", "Word Menu is not displayed");
        }

        [Given(@"I launched new MS Word document")]
        public void GivenILaunchedNewMicrosoftWordDocument()
        {
            Assert.IsNotNull(BaseSteps.CurrentWindowUnderTest);

            if (Actions.IsExists("Pne_NewDocContent_Txt","New Document is not displayed"))
            {
                Panel PanelElement = (Panel)Actions.GetElement("Pne_NewDocContent_Txt");
                TextBox txt = (TextBox)PanelElement.Items[0];
                Actions.VerifyText(txt.Text, string.Empty,"Text is not displayed");
            }
        }

   
        [Given(@"I Click Word menu in MS Word Document")]
        [When(@"I Click Word menu in MS Word Document")]
        public void WhenIClickWordMenuInMicrosoftWordDocument()
        {
            Actions.ClickElement("Tab_WordMenu_Txt");
        }

     
        [Then(@"I should see Word ribbon with below options")]
        public void ThenIShouldSeeWordRibbonWithBelowOptions(Table table)
        {
            Actions.Exists("Pne_WordRibbon_Txt");
            BaseSteps.LogoutWordAddIn();
            foreach (var row in table.Rows)
            {
                Actions.ExistByText<Button>(row["Menu Options"]);
            }
        }

    
        [When(@"I click Login button")]
        public void WhenIClickLoginButton()
        {
            Actions.ClickElement("Tab_WordMenu_Txt");
            Actions.ClickElement("Btn_Login_Txt");
        }

     
        [When(@"I see login window")]
        public void WhenISeeLoginWindow()
        {
            Window loginWindow;
            Actions.ClickElement("Tab_WordMenu_Txt");
            if (Actions.IsExists("Btn_Login_Txt","Login button is not dispalyed"))
            {
                Actions.ClickElement("Btn_Login_Txt");
                loginWindow = Actions.GetModalWindow("Win_Loing_Txt");
            }
            loginWindow = Actions.GetModalWindow("Win_Loing_Txt");
            Assert.IsNotNull(loginWindow);
        }

     
        [When(@"I click Cancel button in login screen in Word AddIn")]
        public void WhenIClickCancelButtonInLoginScreenInWordAddIn()
        {
            Actions.ClickElement("Btn_Cancel_Txt");
        }

      
        [Then(@"Login Dialog should be disappeared")]
        public void ThenLoginDialogShouldBeDisappeared()
        {
            Actions.IsNotExists("Btn_Cancel_Txt","Cancel Button is not displayed");
        }

      
        [Then(@"I should be see login screen")]
        public void ThenIShouldBeSeeLoginScreen()
        {
            var loginWindow = Actions.GetModalWindow("Win_Loing_Txt");
            Assert.IsNotNull(loginWindow);
            Actions.ClickElement("Btn_Cancel_Txt");
        }

      
        [When(@"I enter ""(.*)"" and ""(.*)"" and click Login button")]
        public void WhenIEnterAndAndClickLoginButton(string userName, string password)
        {
            TestData.UserName = userName;
            TestData.Password = password;
            BaseSteps.LoginIntoWord(TestData.UserName, TestData.Password);
        }

       
        [When(@"I selected customer ""(.*)""")]
        public void WhenISelectedCustomer(string customer)
        {
            BaseSteps.SelectCustomer(customer);
        }
       
        [Then(@"I should be logged in successfully")]
        public void ThenIShouldBeLoggedInSuccessfully()
        {
            Actions.Exists("Btn_Logout_Txt");
        }

       
        [Given(@"I logged in with ""(.*)"" and ""(.*)"" credentials")]
        public void GivenILoggedInWithAndCredentials(string userName, string password)
        {
            Actions.ClickElement("Tab_WordMenu_Txt");
            if (!Actions.IsExists("Btn_Logout_Txt"))
            {
                WhenIClickLoginButton();
                WhenIEnterAndAndClickLoginButton(userName, password);
            }
        }

       
        [When(@"I click logout option and confirm")]
        public void WhenIClickLogoutOptionAndConfirm(Table table)
        {
            Actions.ClickElement("Btn_Logout_Txt");
            var LogoutWindow = Actions.GetModalWindow("Win_Confirmation_Txt");
            if (LogoutWindow!=null)
            {
                Actions.ClickElementByTextInModal<Button>("Yes");
            }
            else
            {
                LogoutWindow = Actions.GetModalWindow("Win_Logout_Txt");
                Assert.IsNotNull(LogoutWindow);
                foreach (var row in table.Rows)
                {
                    var LogoutConfirmText = Actions.GetText("Lbl_LogoutMsg_ID");
                    Actions.VerifyText(LogoutConfirmText, row["Message"], "Message is not displayed as expected");
                }
                Actions.ClickElement("Btn_Logout_ID");
            }
        }

       
        [Then(@"I should be logged out successfully")]
        public void ThenIShouldBeLoggedOutSuccessfully()
        {
            Actions.Exists("Btn_Login_Txt");
        }

       
        [When(@"I click ""(.*)"" button")]
        [Given(@"I click ""(.*)"" button")]
        public void WhenIClickButton(string buttonText)
        {
            Actions.ClickElement("Tab_WordMenu_Txt");
            Actions.ClickButtonByText(buttonText);
        }

        
        [Given(@"I see Compose Form Window")]
        public void GivenISeeComposeFormWindow()
        {
            Actions.ClickElement("Tab_WordMenu_Txt");
            Actions.ClickElement("Btn_ComposeForm_Txt");
            Window ComposeFormWindow = Actions.GetChildWindow("Win_ComposeFrom_Txt");
            Assert.IsNotNull(ComposeFormWindow);
        }

        
        [Given(@"I enter data into Compose Form")]
        public void GivenIEnterDataIntoComposeForm(Table table)
        {
            List<string> ComposeFormValues = new List<string>();
            foreach (var row in table.Rows)
            {
                ComposeFormValues.Add(row["Value"]);
            }
            BaseSteps.UpdateComposeFormWithData(ComposeFormValues.ToArray());
        }

        
        [Then(@"Document with form data should be displayed")]
        public void ThenDocumentWithFormDataShouldBeDisplayed()
        {
            System.Threading.Thread.Sleep(5000);
            Window DocWindow = Actions.GetExistingWordDocument;
            Assert.IsNotNull(DocWindow);
            var ExpectedText = "THIS ENDORSEMENT CHANGES THE POLICY. PLEASE READ IT CAREFULLY.\r\rTITLE OF THE FORM GOES HERE\r\rThis endorsement modifies Insurance provided under the following:\r" + TestData.Coverages[0] + "\v\r\f\f";
            Actions.VerifyText(ExpectedText, BaseSteps.GetDocumnetContent(BaseSteps.GetDocumentTitle()), "Text is not displayed as expected");
        }

        
        [Then(@"Compose form fields flown into View Info")]
        public void ThenComposeFormFieldsFlownIntoViewInfo()
        {
            Actions.ClickElement("Tab_WordMenu_Txt");
            Actions.ClickElement("Btn_ViewInfo_Txt");
            Actions.VerifySelectedComboBoxValue("Cmb_cmbLineOfbusiness_ID", TestData.LOB);
            BaseSteps.VerifySelectedCoveragesInViewInfo(TestData.Coverages);
        }

        
        [When(@"I see Comments section is displayed")]
        public void ThenISeeCommentsSectionIsDisplayed()
        {
            Actions.Exists("Tab_COMMENTS_Txt");
        }

        
        [When(@"I add ""(.*)"" comments and Click Save button")]
        public void WhenIAddComments(string comment)
        {
            Actions.SetText("Txt_CommentsArea_ID", comment);
            Actions.ClickElement("Btn_SaveCommentsButton_ID");
        }

        
        [Then(@"I see Comments are added")]
        public void ThenISeeCommentsAreAdded()
        {
            string DateFormat = "";
            if (DateTime.Now.Month < 10)
            {
                DateFormat = "0" + DateTime.Now.Month + "/" + DateTime.Now.Day + "/" + DateTime.Now.Year;
            }
            else
            {
                DateFormat = DateTime.Now.Month + "/" + DateTime.Now.Day + "/" + DateTime.Now.Year;
            }
            Actions.VerifyText(DateFormat, Actions.GetElement("Lbl_DateTextBlock_ID").Name, "Defect VMV-1841");
            
            Actions.VerifyText(Users.User_Name, Actions.GetElement("Lbl_UserNameTextBlock_ID").Name, "User Name is not dispalyed as expected");

            Actions.VerifyText(Actions.GetControlValue("TestComments"), Actions.GetElement("Lbl_CommentTextBlock_ID").Name, "Comments are not added as expected");

            Actions.Exists("Btn_DeleteCommentButton_ID", "Defect VMV-1841");
        }

        
        [Then(@"I add ""(.*)"" as reply comment")]
        public void ThenIAddAsReplyComment(string comment)
        {
            Actions.ClickElement("Btn_ReplyCommentsButton_ID");
            Actions.GetElements("Txt_CommentsArea_ID")[1].SetValue(comment);
            Actions.ClickElement("Btn_ReplyComments_ID");
        }

        
        [Then(@"I see Reply comments are added")]
        public void ThenISeeReplyCommentsAreAdded()
        {
            string DateFormat = "";
            if (DateTime.Now.Month < 10)
            {
                DateFormat = "0" + DateTime.Now.Month + "/" + DateTime.Now.Day + "/" + DateTime.Now.Year;
            }
            else
            {
                DateFormat = DateTime.Now.Month + "/" + DateTime.Now.Day + "/" + DateTime.Now.Year;
            }
            Actions.VerifyText(DateFormat, Actions.GetElements("Lbl_DateTextBlock_ID")[1].Name, "Defect VMV-1841");

            Actions.VerifyText(Users.User_Name, Actions.GetElements("Lbl_UserNameTextBlock_ID")[1].Name, "User Name is not dispalyed as expected");

            Actions.VerifyText(Actions.GetControlValue("ReplyComment"), Actions.GetElements("Lbl_CommentTextBlock_ID")[1].Name, "Reply Comments are not added as expected");

            Actions.ElementVisible(Actions.GetElements("Btn_DeleteCommentButton_ID")[1], "Defect VMV-1841");
        }

        
        [Then(@"I add max number of characters in comments")]
        public void ThenIAddMaxNumberOfCharactersInComments()
        {
            Actions.SetText("Txt_CommentsArea_ID", Actions.GetControlValue("CommentWithMaxChar"));
            Actions.GetElement("Txt_CommentsArea_ID");
            
            Actions.ClickElement("Btn_SaveCommentsButton_ID");
            var CommentsDisplayed = Actions.GetElements("Lbl_CommentTextBlock_ID")[0].Name;
            Actions.VerifyText(CommentsDisplayed, Actions.GetControlValue("CommentWithMaxChar"), "Defect VMV-1841");
        }

        
        [Given(@"View Info button is enabled")]
        public void GivenViewInfoButtonIsEnabled()
        {
            Actions.ClickElement("Tab_WordMenu_Txt");
            Actions.Exists("Btn_ViewInfo_Txt");
        }

        
        [Given(@"Save to Word button is enabled")]
        public void GivenSaveToWordButtonIsEnabled()
        {
            Actions.ClickElement("Tab_WordMenu_Txt");
            Actions.Exists("Btn_SavetoWord_Txt");
        }


       
    }
}