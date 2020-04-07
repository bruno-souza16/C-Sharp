using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Threading;

namespace VariousFileFunctions
{
    class BrowserFunctions
    {
        static IWebDriver driver;
        static ChromeOptions chromeOptions;

        //Aux Functions
        #region Auxiliar Functions
        // --------------------------------


        // Wait page's or action's load.
        // driver --> chrome selenium driver
        static void WaitForLoad(IWebDriver driver, int timeoutSec = 15)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, timeoutSec));
            wait.Until(wd => js.ExecuteScript("return document.readyState").ToString() == "complete");
        }
        
        // Load driver, set download's folder and chrome options
        // downloadFolder --> Fullpath of download folder
        public static bool LoadDriver(string downloadFolder)
        {
            try
            {
                chromeOptions = new ChromeOptions();
                chromeOptions.AddUserProfilePreference("download.default_directory", downloadFolder);
                chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
                driver = new ChromeDriver(chromeOptions);
                return true;

            }
            catch
            {
                return false;
            }
        }

        // --------------------------------
        #endregion

        //Navigate Functions
        #region Navigate Functions
        // --------------------------------


        // Loads a url
        // url --> Site's path
        public static bool OpenURL(string url)
        {
            try
            {
                driver.Navigate().GoToUrl(url);
                driver.Manage().Window.Maximize();
                return true;
            }
            catch
            {
                CloseChrome();
                return false;
            }
        }

        public static bool CloseChrome()
        {
            try
            {
                driver.Close();
                driver.Dispose();
                var processes1 = Process.GetProcessesByName("Chrome");
                foreach (var p in processes1)
                    p.Kill();
                return true;
            }
            catch
            {
                return false;
            }
           
        }

        public static bool ClickObject(string obj)
        {
            try
            {
                IWebElement element = driver.FindElement(By.ClassName(obj));
                element.Click();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool WriteObjectField(string obj, string text)
        {
            try
            {
                IWebElement element = driver.FindElement(By.ClassName(obj));
                element.SendKeys("");
                Thread.Sleep(2000);
                element.SendKeys(text);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static bool ConfirmAlert()
        {
            driver.FindElement(By.Id("submitButton")).Click();
            try
            {
                Alert alert = driver.switchTo().alert();                
                alert.accept();
                return true;
            }
            catch (NoAlertPresentException ex)
            {
                ex.printStackTrace();
                return false;
            }
        }    

        // --------------------------------
        #endregion
    }
}
