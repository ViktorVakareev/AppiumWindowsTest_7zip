using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.IO;
using System.Threading;

namespace AppiumWindowsTest_7zip
{
    public class AppiumWinwows_7zipTest
    {
        private WindowsDriver<WindowsElement> driver;
        private WindowsDriver<WindowsElement> tempDriver;

        private const string Open7zipPath = @"C:\Program Files\7-Zip\7zFM.exe";   // fields
        private const string AppPath = @"C:\Program Files\7-Zip\";
        private const string AppiumServerUrl = "http://[::1]:4723/wd/hub";
        private string tempDir;

        [OneTimeSetUp]
        public void Setup()
        {

            var appiumOptions = new AppiumOptions();
            appiumOptions.AddAdditionalCapability("platformName", "Windows");
            appiumOptions.AddAdditionalCapability("app", Open7zipPath);

            driver = new WindowsDriver<WindowsElement>
                (new Uri(AppiumServerUrl), appiumOptions);

            var appiumOptionsTemp = new AppiumOptions();
            appiumOptionsTemp.AddAdditionalCapability("platformName", "Windows");
            appiumOptionsTemp.AddAdditionalCapability("app", "Root");

            tempDriver = new WindowsDriver<WindowsElement>
               (new Uri(AppiumServerUrl), appiumOptionsTemp);

            tempDir = Directory.GetCurrentDirectory() + @"\tempDir";  // create new dir in 7zip directory

            if (Directory.Exists(tempDir))                //if dir already exist-delete it!
            {
                Directory.Delete(tempDir, true);
                Directory.CreateDirectory(tempDir);
            }
        }
        [Test]
        public void Test7zip()
        {
            Thread.Sleep(500);
            var fileLocationTextBox = driver.FindElement(By.XPath("/Window/Pane/Pane/ComboBox/Edit"));
            fileLocationTextBox.Click();
            fileLocationTextBox.Clear();
            fileLocationTextBox.SendKeys(AppPath);
            fileLocationTextBox.SendKeys(Keys.Enter);
            Thread.Sleep(1000);

            var directoryListOfFiles = driver.FindElementByXPath("/Window/Pane/List");
            //directoryListOfFiles.Click();
            directoryListOfFiles.SendKeys(Keys.Control + "a");

            var addToArchiveButton = driver.FindElementByName("Добавяне");
            addToArchiveButton.Click();
            Thread.Sleep(500);

            // We are working in "Add to Archive(Добавяне към архив)" window, which is operated by tempDriver!!!
            var newWindowAddToArchive = tempDriver.FindElementByName("Добавяне към архив");

            Thread.Sleep(500);
            var fileSaveLocationBox = newWindowAddToArchive.FindElementByXPath
                ("/Window/ComboBox/Edit[@ClassName=\"Edit\"][@Name=\"Архив:\"]");
            fileSaveLocationBox.Click();
            fileSaveLocationBox.SendKeys(Keys.Control + "a");
            fileSaveLocationBox.Clear();
            var archiveName = tempDir + "\\" + DateTime.Now.Ticks + ".7z";
            string folderName = tempDir + @"\" + DateTime.Now.Ticks;
            fileSaveLocationBox.SendKeys(archiveName);

            var formatArchiveBox = newWindowAddToArchive.FindElementByXPath
                ("/Window/ComboBox[@Name='Формат на архива:']");
            formatArchiveBox.Click();
            formatArchiveBox.SendKeys(Keys.Home + Keys.Enter);

            var compressionLevelBox = newWindowAddToArchive.FindElementByXPath
                ("/Window/ComboBox[@Name='Ниво на компресия:']");
            compressionLevelBox.Click();
            compressionLevelBox.SendKeys(Keys.End + Keys.Enter);

            var compressionMethodBox = newWindowAddToArchive.FindElementByXPath
                ("/Window/ComboBox[@Name='Метод за компресия:']");
            compressionMethodBox.Click();
            compressionMethodBox.SendKeys(Keys.Home + Keys.Enter);

            var vocabularySizeBox = newWindowAddToArchive.FindElementByXPath
                ("/Window/ComboBox[@Name=\"Размер на речника:\"]");
            vocabularySizeBox.Click();
            vocabularySizeBox.SendKeys(Keys.End + Keys.Enter);

            var wordSizeBox = newWindowAddToArchive.FindElementByXPath
                ("/Window/ComboBox[@Name=\"Размер на думата:\"]");
            wordSizeBox.Click();
            wordSizeBox.SendKeys(Keys.End + Keys.Enter);

            var buttonOK = newWindowAddToArchive.FindElementByName("OK");
            buttonOK.Click();

            Thread.Sleep(1000);

            fileLocationTextBox.Click();
            fileLocationTextBox.SendKeys(archiveName);
            fileLocationTextBox.SendKeys(Keys.Enter);

            var buttonExtract = driver.FindElementByName("Извличане");
            buttonExtract.Click();

            var buttonExtraxtOK = driver.FindElementByName("OK");
            buttonExtraxtOK.Click();
            Thread.Sleep(1000);

            // Assertions

            string extractedZip = tempDir + @"\7zFM.exe";
            FileAssert.AreEqual(Open7zipPath, extractedZip);

        }

        [OneTimeTearDown]
        public void Shutdown()
        {
            //driver.Quit();
        }

    }
}




