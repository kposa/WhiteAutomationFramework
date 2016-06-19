namespace SpecflowDemo.Helpers
{
    using System;
    using System.Diagnostics;
    using System.Globalization;
    using System.Reflection;
    using System.Resources;
    using NUnit.Framework;
    using TestStack.White;
    using TestStack.White.Factory;
    using TestStack.White.UIItems;
    using TestStack.White.UIItems.Finders;
    using TestStack.White.UIItems.ListBoxItems;
    using TestStack.White.UIItems.TabItems;
    using TestStack.White.UIItems.TreeItems;
    using TestStack.White.UIItems.WindowItems;
    using System.Collections.Generic;
    using System.Threading;

    /// <summary>
    /// The Actions class
    /// </summary>
    public static class Actions
    {
        public static T GetByText<T>(string text) where T : UIItem
        {
            T uiItem;
            try
            {
                uiItem = BaseSteps.CurrentWindowUnderTest.Get<T>(SearchCriteria.ByText(text));
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Retrieving " + text,
                            "Text of " + text,
                            "PASS");
            }
            catch (AutomationException e)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Retrieving Text of" + text,
                            "Exception " + e,
                            "FAIL");
                return null;
            }
            return uiItem;
        }

        public static void SelectListOption(string listBoxId, int index)
        {
            var listOptions = (ListBox)GetElement(listBoxId);
            if (listOptions != null)
            {
                listOptions.Select(index);
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Selecting "+index+" nd Option of "+listBoxId,
                            index+" nd Option of "+listBoxId+" is selected",
                            "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Selecting "+index+" nd Option of "+listBoxId,
                           index+" nd Option of "+listBoxId+" is not selected",
                           "FAIL");
            }
        }

        public static void SelectListOptionInModal(string listBoxId, int index)
        {
            var listOptions = (ListBox)GetElementInModal(listBoxId);
            if (listOptions != null)
            {
                listOptions.Select(index);
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Selecting " + index + " nd Option of " + listBoxId,
                            index + " nd Option of " + listBoxId + " is selected",
                            "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Selecting " + index + " nd Option of " + listBoxId,
                           index + " nd Option of " + listBoxId + " is not selected",
                           "FAIL");
            }
        }

        public static void SelectComboBoxOptionByIndex(string comboBoxId, int index)
        {
            var comboBoxOptions = (ComboBox)GetElement(comboBoxId);
            if (comboBoxOptions != null)
            {
                comboBoxOptions.Select(index);
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Selecting "+index+" nd Option of "+comboBoxId,
                            index+" nd Option of "+comboBoxId+" is selected",
                            "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Selecting "+index+" nd Option of "+comboBoxId,
                           index+" nd Option of "+comboBoxId+" is not selected",
                           "FAIL");
            }
        }

        public static void SelectComboBoxOptionByText(string optionText, string value)
        {
            var comboBoxOptions = (ComboBox)GetElement(optionText);
            if (comboBoxOptions != null)
            {
                comboBoxOptions.Select(value);
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Selecting "+value+" Option of "+optionText,
                            value+" Option of "+optionText+" is selected",
                            "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Selecting "+value+" Option of "+optionText,
                           value+" Option of "+optionText+" is not selected",
                           "FAIL");
            }
        }

         public static void SelectListBoxOptionByText(string optionText, string val)
        {
            var ListBoxOptions = (ListBox)GetElement(optionText);
            if (ListBoxOptions != null)
            {
                ListBoxOptions.Select(val);
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Selecting " + val + " nd Option of " + optionText,
                            val + " nd Option of " + optionText + " is selected",
                            "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Selecting " + val + " nd Option of " + optionText,
                           val + " nd Option of " + optionText + " is not selected",
                           "FAIL");
            }
        }

        public static void VerifySelectedComboBoxValue(string optionText, string expectedValue)
        {
            try
            {
                var ComboBoxOptions = (ComboBox)GetElement(optionText);
                ComboBoxOptions.DoubleClick();
                if (expectedValue == ComboBoxOptions.SelectedItemText)
                {
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                "Verify selected value in " + optionText,
                                expectedValue + " is selected",
                                "PASS");
                }
                else
                {
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                               "Verify selected value in " + optionText,
                               expectedValue + " is not selected",
                               "FAIL");
                    Assert.Fail("Expected Options are not selected");
                }
            }
            catch (AutomationException e)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                               "Verify selected value in " + optionText,
                                               "Exception:" + e,
                                               "FAIL");
            }
        }

        public static void VerifyComboBoxDefaultSelectedValue(string comboBoxPropety,string expectedValue)
        {
            var ComboBoxOptions = (ComboBox)GetElement(comboBoxPropety);
            try
            {
                if(ComboBoxOptions.EditableText == expectedValue)
                {
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                "Verify selected value in " + comboBoxPropety,
                                expectedValue + " is selected",
                                "PASS");
                }
                else
                {
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                               "Verify selected value in " + comboBoxPropety,
                               expectedValue + " is not selected",
                               "FAIL");
                    Assert.Fail("Expected Option is not selected");
                }
            }
            catch (AutomationException e)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                                "Verify selected value in " + comboBoxPropety,
                                                "Exception:" + e,
                                                "FAIL");
            }

        }

      public static void ClickButtonByText(string buttonName)
        {
            var option = GetByText<Button>(buttonName);
            if (option != null)
            {
                option.Click();
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Clicking " + buttonName,
                            "Successfully clicked " + buttonName,
                            "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Clicking " + buttonName,
                            "Clicking " + buttonName + " Failed",
                            "FAIL");
            }
        }

        public static UIItem GetElement(string controlKey)
        {
            string controlType = controlKey.Substring(0, 3);
            string idType = controlKey.Substring(controlKey.Length - 3, 3);
            SearchCriteria searchCriteria;
            UIItem uiItem=null;
            if (idType == "_ID")
            {
                searchCriteria = SearchCriteria.ByAutomationId(GetControlValue(controlKey));
            }
            else
            {
                searchCriteria = SearchCriteria.ByText(GetControlValue(controlKey));
            }

            try
            {
                switch (controlType)
                {
                    case "Btn":
                        uiItem = BaseSteps.CurrentWindowUnderTest.Get<Button>(searchCriteria);
                        break;
                    case "Chk":
                        uiItem = BaseSteps.CurrentWindowUnderTest.Get<CheckBox>(searchCriteria);
                        break;
                    case "Cmb":
                        uiItem = BaseSteps.CurrentWindowUnderTest.Get<ComboBox>(searchCriteria);
                        break;
                    case "Lbl":
                        uiItem = BaseSteps.CurrentWindowUnderTest.Get<Label>(searchCriteria);
                        break;
                    case "Lnk":
                        uiItem = BaseSteps.CurrentWindowUnderTest.Get<Hyperlink>(searchCriteria);
                        break;
                    case "Lst":
                        uiItem = BaseSteps.CurrentWindowUnderTest.Get<ListBox>(searchCriteria);
                        break;
                    case "Tab":
                       uiItem = BaseSteps.CurrentWindowUnderTest.Get<TabPage>(searchCriteria);
                        break;
                    case "Pne":
                        uiItem = BaseSteps.CurrentWindowUnderTest.Get<Panel>(searchCriteria);
                        break;
                    case "Rdo":
                        uiItem = BaseSteps.CurrentWindowUnderTest.Get<RadioButton>(searchCriteria);
                        break;
                    case "Txt":
                        uiItem = BaseSteps.CurrentWindowUnderTest.Get<TextBox>(searchCriteria);
                        break;
                    case "Win":
                        uiItem = BaseSteps.CurrentWindowUnderTest.Get<Window>(searchCriteria);
                        break;
                    case "Tre":
                        uiItem = BaseSteps.CurrentWindowUnderTest.Get<Tree>(searchCriteria);
                        break;
                    case "Cus":
                    case "Dtp":
                        uiItem = (UIItem)BaseSteps.CurrentWindowUnderTest.Get(searchCriteria);
                        break;
                    case "Lsv":
                        uiItem = BaseSteps.CurrentWindowUnderTest.Get<ListView>(searchCriteria);
                        break;
                }

            }
            catch (AutomationException)
            {
                return null;
            }
            return uiItem;
        }

         public static UIItem GetElementInModal(string controlKey)
        {
            Window modal = BaseSteps.ModalWindowUnderTest;
            string controlType = controlKey.Substring(0, 3);
            string idType = controlKey.Substring(controlKey.Length - 3, 3);
            SearchCriteria searchCriteria;
            UIItem uiItem = null;
            if (idType == "_ID")
            {
                searchCriteria = SearchCriteria.ByAutomationId(GetControlValue(controlKey));
            }
            else
            {
                searchCriteria = SearchCriteria.ByText(GetControlValue(controlKey));
            }

            try
            {
                switch (controlType)
                {
                    case "Btn":
                        uiItem = modal.Get<Button>(searchCriteria);
                        break;
                    case "Chk":
                        uiItem = modal.Get<CheckBox>(searchCriteria);
                        break;
                    case "Cmb":
                        uiItem = modal.Get<ComboBox>(searchCriteria);
                        break;
                    case "Lbl":
                        uiItem = modal.Get<Label>(searchCriteria);
                        break;
                    case "Lnk":
                        uiItem = modal.Get<Hyperlink>(searchCriteria);
                        break;
                    case "Lst":
                        uiItem = modal.Get<ListBox>(searchCriteria);
                        break;
                    case "Tab":
                        uiItem = modal.Get<TabPage>(searchCriteria);
                        break;
                    case "Pne":
                        uiItem = modal.Get<Panel>(searchCriteria);
                        break;
                    case "Rdo":
                        uiItem = modal.Get<RadioButton>(searchCriteria);
                        break;
                    case "Txt":
                        uiItem = modal.Get<TextBox>(searchCriteria);
                        break;
                    case "Win":
                        uiItem = modal.Get<Window>(searchCriteria);
                        break;
                    case "Tre":
                        uiItem = modal.Get<Tree>(searchCriteria);
                        break;
                    case "Cus":
                    case "Dtp":
                        uiItem = (UIItem)modal.Get(searchCriteria);
                        break;
                    case "Lsv":
                        uiItem = modal.Get<ListView>(searchCriteria);
                        break;
                }

            }
            catch (AutomationException)
            {
                return null;
            }
            return uiItem;
        }

       public static IUIItem[] GetElements(string controlKey)
        {
            string idType = controlKey.Substring(controlKey.Length - 3, 3);
            SearchCriteria searchCriteria;
            IUIItem[] uiItem = null;
            if (idType == "_ID")
            {
                searchCriteria = SearchCriteria.ByAutomationId(GetControlValue(controlKey));
            }
            else
            {
                searchCriteria = SearchCriteria.ByText(GetControlValue(controlKey));
            }

            try
            {
               uiItem = BaseSteps.CurrentWindowUnderTest.GetMultiple(searchCriteria);
            }
            catch (AutomationException e)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Returning  " + controlKey,
                            "Returning " + e,
                            "FAIL");
                return null;
            }
            BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Returning  " + controlKey,
                            "Successfully returned" + controlKey + " Elements",
                            "PASS");
            return uiItem;
        }

        public static void ExistByText<T>(string text) where T : UIItem
        {
            T uiItem = null;
            try
            {
                uiItem = BaseSteps.CurrentWindowUnderTest.Get<T>(SearchCriteria.ByText(text));
            }
            catch (AutomationException e)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Exists  " + text,
                            "Exception "+e,
                            "FAIL");
            }
            
            BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Exists  " + text,
                           text + " Element is available",
                          "PASS");
            Assert.IsNotNull(uiItem);
        }

           public static void ClickElementByText<T>(string text) where T : UIItem
        {
            T uiItem = null;
            try
            {
                uiItem = BaseSteps.CurrentWindowUnderTest.Get<T>(SearchCriteria.ByText(text));
                uiItem.Click();
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Exists  " + text,
                           text + " Element is clicked",
                          "PASS");
            }
            catch (AutomationException e)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Click  " + text,
                           "Exception "+e,
                            "FAIL");
            }
        }

        public static void ClickElementByTextInModal<T>(string text) where T : UIItem
        {
            T uiItem = null;
            try
            {
                uiItem = BaseSteps.ModalWindowUnderTest.Get<T>(SearchCriteria.ByText(text));
                uiItem.Click();
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Exists  " + text,
                           text + " Element is clicked",
                          "PASS");
            }
            catch (AutomationException e)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Click  " + text,
                            "Exception "+e,
                            "FAIL");
            }
        }

         public static void SetText(string controlKey, string text)
        {
            var element = GetElement(controlKey);
            if(element!=null)
            {
                element.SetValue(text);
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Set Text to  " + controlKey,
                           " Text "+ text+" is entered",
                          "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Set Text to  " + controlKey,
                            controlKey + " element is not editable",
                            "FAIL");
            }   
        }

        public static string GetControlValue(string controlKey)
        {
            string value = null;
            try
            {
                ResourceManager rm = new ResourceManager("SpecflowDemoProject.GuiMaps.GuiMap",Assembly.GetExecutingAssembly());
                value = rm.GetString(controlKey, CultureInfo.CurrentCulture);
            }
            catch (AutomationException e)
            {
                Assert.Fail("Excetion: "+e);
            }

            return value;
        }

          public static string GetUserCredentials(string userKey)
        {
            string value = null;
            try
            {
                ResourceManager rm = new ResourceManager("SpecflowDemoProject.GuiMaps.Users",
                Assembly.GetExecutingAssembly());
                value = rm.GetString(userKey, CultureInfo.CurrentCulture);
            }
            catch (AutomationException e)
            {

                Assert.Fail("Exception: "+e);
            }

            return value;
        }

        public static void ClickElement(string controlKey)
        {
            var Option = GetElement(controlKey);
            
            if (Option!=null)
            {
                Option.Click();
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Click  " + controlKey,
                           controlKey + " Element is clicked",
                          "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Click  " + controlKey,
                           controlKey + " Element not found",
                          "FAIL");
                Assert.Fail(controlKey + " Element not found");
            }

        }

       public static void ClickElementInModal(string controlKey)
        {
            var Option = GetElementInModal(controlKey);

            if (Option != null)
            {
                Option.Click();
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Click  " + controlKey,
                           controlKey + " Element is clicked",
                          "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Click  " + controlKey,
                           controlKey + " Element not found",
                          "FAIL");
                Assert.Fail(controlKey + " Element not found");
            }

        }

         public static void ClickElementToDismiss(string controlKey)
        {
            var mouseObj = BaseSteps.CurrentWindowUnderTest.Mouse;
            var Option = GetElement(controlKey);
            if (Option != null)
            {
                mouseObj.Location = Option.AutomationElement.GetClickablePoint();
                mouseObj.Click();
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Click  " + controlKey,
                           controlKey + " Element is clicked",
                          "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Click  " + controlKey,
                            controlKey + " Element is not found",
                           "FAIL");
                Assert.Fail(controlKey + " Element not found");
            }
        }

        public static void Exists(string controlKey,string message)
        {
            var element = GetElement(controlKey);
            if (element!=null)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Is Exist  " + controlKey,
                           controlKey + " Element is available",
                          "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Is Exist  " + controlKey,
                            message,
                           "FAIL");
                Assert.Fail(controlKey + " Element not found");
            }
            
        }

         public static void ExistsInModal(string controlKey)
        {
            var element = GetElementInModal(controlKey);
            if (element != null)
            {
                Assert.IsNotNull(element);
                Assert.IsTrue(element.Visible);
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Is Exist  " + controlKey,
                           controlKey + " Element is available",
                          "PASS");
            }
            else
            {
                Assert.IsNotNull(element);
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Is Exist  " + controlKey,
                            controlKey + " Element is not available",
                           "FAIL");
                Assert.Fail(controlKey + " Element not found");
            }

        }

      public static void IsNotExists(string controlKey, string message)
        {
            var element = GetElement(controlKey);
            if (element == null)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Is Exist  " + controlKey,
                           controlKey + " Element is not available",
                          "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Is Exist  " + controlKey,
                            controlKey + " Element is available",
                           "FAIL");
                Assert.Fail(controlKey + " Element still visible");
            }
            Assert.IsNull(element, message);
        }

        public static void Exists(string controlKey)
        {
            var element = GetElement(controlKey);
            if (element!=null)
            {
                Assert.IsNotNull(element);
                Assert.IsTrue(element.Visible);
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Is Exist  " + controlKey,
                           controlKey + " Element is available",
                          "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Is Exist  " + controlKey,
                            controlKey + " Element is not available",
                           "FAIL");
                Assert.Fail(controlKey + " Element not found");
            }
            
        }

       public static bool IsExists(string controlKey, string message)
        {
            var element = GetElement(controlKey);
            if (element!=null)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Is Exist  " + controlKey,
                           controlKey + " Element is available",
                          "PASS");
                Assert.IsTrue(element.Visible, message);
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Is Exist  " + controlKey,
                            controlKey + " Element is not available",
                           "FAIL");
            }

            return (element != null);
        }

       public static bool IsExists(string controlKey)
        {
            var element = GetElement(controlKey);
            if (element!=null)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                          "Is Exist  " + controlKey,
                                           controlKey + " Element is available",
                                          "PASS");
                Assert.IsTrue(element.Visible);
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Is Exist  " + controlKey,
                            controlKey + " Element is not available",
                           "FAIL");
            }
            return (element != null);
        }

        public static void ElementVisible(IUIItem control,string message)
        {
            if (control != null)
            {
                if (control.Visible)
                {
                    Assert.IsTrue(control.Visible, message);
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                          "Is Visible  " + control,
                                           control + " Element is Visible",
                                          "PASS");
                }
                else
                {
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Is Visible  " + control,
                            control + " Element is not Visible",
                           "FAIL");
                }
            }
            
        }
        public static void ElementVisible(IUIItem control)
        {
            if (control != null)
            {
                if (control.Visible)
                {
                    Assert.IsTrue(control.Visible);
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                          "Is Visible  " + control,
                                           control + " Element is Visible",
                                          "PASS");
                }
                else
                {
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Is Visible  " + control,
                            control + " Element is not Visible",
                           "FAIL");
                }
            }
        }

        public static bool ElementsExists(string controlKey,int count)
        {
            IUIItem[] element;
            try
            {
                element = GetElements(controlKey);
                if (element.Length == count)
                {
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                              "Elements count  " + controlKey,
                                               controlKey + " Elements count is" + count,
                                              "PASS");
                }
                else
                {
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                                "Elements count  " + controlKey,
                                                controlKey + " Elements count is" + count,
                           "FAIL");
                    Assert.Fail(controlKey + " elements are not visible");
                }
            }
            catch (AutomationException e)
            {
                Assert.Fail("Exception: "+e);
                return false;
            }
            
            return (element.Length == count);
        }
        
         public static Window GetChildWindow(string docTitle)
        {
            string DocumentTitle = GetControlValue(docTitle);
            BaseSteps.ModalWindowUnderTest = BaseSteps.AppUnderTest.GetWindow(SearchCriteria.ByText(DocumentTitle), InitializeOption.NoCache);
            return BaseSteps.ModalWindowUnderTest;
        }

        public static Window GetChildWindowByText(string docTitle)
        {
            BaseSteps.ModalWindowUnderTest = BaseSteps.AppUnderTest.GetWindow(SearchCriteria.ByText(docTitle), InitializeOption.NoCache);
            return BaseSteps.ModalWindowUnderTest;
        }

        public static string GetText(string controlKey)
        {
            var element = GetElement(controlKey);
            return element.Name;
        }

         public static Window GetModalWindow(string windowTitle)
        {
            BaseSteps.ModalWindowUnderTest = BaseSteps.CurrentWindowUnderTest.ModalWindow(GetControlValue(windowTitle));
            return BaseSteps.ModalWindowUnderTest;
        }

       public static void VerifyText(string expectedText, string actualText, string message)
        {
            if (string.Compare(actualText,expectedText,StringComparison.CurrentCulture) ==0)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                          "Verify Text",
                                           "ActualText is displayed as ExpectedText",
                                          "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                          "Verify Text",
                                           "ActualText is not displayed as ExpectedText",
                                          "FAIL");
            }
            Assert.AreEqual(expectedText, actualText, message);
        }

       public static Window GetExistingWordDocument
        {
            get
            {
                Process p = null;
                int counter = 0;
                do
                {
                    foreach (var process in Process.GetProcessesByName("WINWORD"))
                    {
                        p = process;
                    }
                    if (p == null)
                    {
                        System.Threading.Thread.Sleep(2000);
                        counter++;
                    }
                } while ((p == null) && (counter < 100));

                BaseSteps.AppUnderTest = Application.Attach(p);
                BaseSteps.CurrentWindowUnderTest = BaseSteps.AppUnderTest.GetWindow((SearchCriteria.ByClassName("OpusApp")), InitializeOption.NoCache);
                return BaseSteps.CurrentWindowUnderTest;
            }
            
        }

       public static Window GetLoginPopupOnBrowser
        {
            get
            {
                Process p = null;
                int counter = 0;
                do
                {
                    foreach (var process in Process.GetProcessesByName("SpecflowDemoProjectLauncher"))
                    {
                        p = process;
                    }
                    if (p == null)
                    {
                        System.Threading.Thread.Sleep(2000);
                        counter++;
                    }
                } while ((p == null) && (counter < 100));

                BaseSteps.AppUnderTest = Application.Attach(p);
                BaseSteps.ModalWindowUnderTest = BaseSteps.AppUnderTest.GetWindow((SearchCriteria.ByAutomationId("LoginForm")), InitializeOption.NoCache);
                return BaseSteps.ModalWindowUnderTest;
            }

        }

        
          public static List<Window> GetExistingWordDocuments
        {
            get
            {
                List<Window> ExistingWindows;
                Process p = null;
                int counter = 0;
                do
                {
                    foreach (var process in Process.GetProcessesByName("WINWORD"))
                    {
                        p = process;
                    }
                    if (p == null)
                    {
                        System.Threading.Thread.Sleep(2000);
                        counter++;
                    }
                    BaseSteps.AppUnderTest = Application.Attach(p);
                    ExistingWindows = BaseSteps.AppUnderTest.GetWindows();
                } while ((p == null) && (counter < 100)&& ExistingWindows.Count<2);
                return ExistingWindows;
            }
        }

         public static void Enabled(string controlKey)
        {
            
            try
            {
                var element = GetElement(controlKey);
                if (element != null && element.Enabled)
                {
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                              "Element  " + controlKey + " Enabled",
                                               controlKey + " Element is enabled",
                                              "PASS");
                }
                else
                {
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                                "Element  " + controlKey + " Enabled",
                                               controlKey + " Element is not enabled",
                                              "FAIL");
                }
            }
            catch (AutomationException e)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                                "Element  " + controlKey + " Enabled",
                                               "Exceptions: "+e,
                                              "FAIL");
                Assert.Fail(controlKey + "Control is not visible");
            }
            }

        public static void IsNotEnabled(string controlKey)
        {
            try
            {
                var element = GetElement(controlKey);
                if (element != null && !element.Enabled)
                {
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                              "Element  " + controlKey + " Not Enabled",
                                               controlKey + " Element is not enabled",
                                              "PASS");
                }
                else
                {
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                                "Element  " + controlKey + " Not Enabled",
                                               controlKey + " Element is enabled",
                                              "FAIL");
                    Assert.Fail(controlKey + " element is not visible");
                }
            }
            catch (AutomationException e)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                                 "Element  " + controlKey + " Not Enabled",
                                                "Exception: "+e,
                                               "FAIL");
                Assert.Fail(controlKey + " element is not visible");
            }
        }

          public static void VerifyComboBoxVauels(string controlKey, string[] values)
        {
            int valuesCount = 0;
            try
            {
                var element = GetElement(controlKey) as ComboBox;
                foreach (var item in element.Items)
                {
                    foreach (var value in values)
                    {
                        if (item.Text == value)
                        {
                            valuesCount++;
                        }
                    }
                }
                if (valuesCount == values.Length)
                {
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                              "Verify" + controlKey + " CombBox options",
                                                    "Options are available in " + controlKey,
                                              "PASS");
                }
                else
                {
                    BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                                     "Verify"+controlKey+" CombBox options",
                                                    "Options are not available in " + controlKey,
                                                   "FAIL");
                    Assert.Fail("Options are not available");
                }
                element.Click();
            }
            catch (AutomationException e)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                                                 "Element  " + controlKey + " Not Enabled",
                                                "Exception: " + e,
                                               "FAIL");
                Assert.Fail(controlKey + " element is not visible");
            }
        }

         public static void SelectListBoxOptionByTextInModal(string optionText, string val)
        {
            var ListBoxOptions = (ListBox)GetElementInModal(optionText);
            if (ListBoxOptions != null)
            {
                ListBoxOptions.Select(val);
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Selecting " + val + " nd Option of " + optionText,
                            val + " nd Option of " + optionText + " is selected",
                            "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Selecting " + val + " nd Option of " + optionText,
                           val + " nd Option of " + optionText + " is not selected",
                           "FAIL");
            }
        }

       public static void SelectComboBoxOptionByTextInModal(string optionText, string value)
        {
            var comboBoxOptions = (ComboBox)GetElementInModal(optionText);
            if (comboBoxOptions != null)
            {
                comboBoxOptions.Select(value);
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Selecting " + value + " nd Option of " + optionText,
                            value + " nd Option of " + optionText + " is selected",
                            "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                           "Selecting " + value + " nd Option of " + optionText,
                           value + " nd Option of " + optionText + " is not selected",
                           "FAIL");
            }
        }

          public static void SetTextInModal(string controlKey, string text)
        {
            var element = GetElementInModal(controlKey);
            if (element != null)
            {
                element.SetValue(text);
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Set Text to  " + controlKey,
                           " Text " + text + " is entered",
                          "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Set Text to  " + controlKey,
                            controlKey + " element is not editable",
                            "FAIL");
            }
        }

        public static void WaitForElementEnable(string controlKey)
        {
            var element = GetElement(controlKey);
            int counter = 0;
            do
            {
                element = GetElement(controlKey);
                Thread.Sleep(1000);
                counter++;

            } while (!element.Enabled && counter<60);

            if (element.Enabled)
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                          "Element Enabled  " + controlKey,
                           " Element " + controlKey + " is enabled",
                          "PASS");
            }
            else
            {
                BaseSteps.WebDriverApi.ReportStepDetailsWithScreenshot(
                            "Element Enabled  " + controlKey,
                           " Element " + controlKey + " is not enabled",
                            "FAIL");
                Assert.Fail(controlKey + " element is not enabled");
            }
        }
    }
}
