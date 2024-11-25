
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using System;
using static System.Collections.Specialized.BitVector32;
using static System.Net.Mime.MediaTypeNames;
using System.Drawing;
using System.Drawing.Design;
using System.Drawing.Imaging;

internal class Program
{
    private static void Main(string[] args)
    {
        IWebDriver driver = new FirefoxDriver();

        driver.Navigate().GoToUrl("https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=165&ct=1732498175&rver=7.5.2211.0&wp=MBI_SSL&wreply=https%3a%2f%2foutlook.live.com%2fowa%2f0%2f%3fstate%3d1%26redirectTo%3daHR0cHM6Ly9vdXRsb29rLmxpdmUuY29tL21haWwvMC8%26RpsCsrfState%3d7f3e52aa-03ed-dca7-e44a-8a51a2dc6e61&id=292841&aadredir=1&whr=outlook.com&CBCXT=out&lw=1&fl=dob%2cflname%2cwld&cobrandid=90015");
        driver.Manage().Window.Maximize();
        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);


        //Iniciar sesion
        string attr = driver.SwitchTo().ActiveElement().GetAttribute("id");
        IWebElement input = driver.FindElement(By.Id("i0116"));
        input.SendKeys("20230920@itla.edu.do");
        input.Submit();

        ITakesScreenshot screenshotdriver = (ITakesScreenshot)driver;
        Screenshot screenshot = screenshotdriver.GetScreenshot();
        screenshot.SaveAsFile("abc.png");
        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
        IWebElement password = driver.FindElement(By.Id("i0118"));
        screenshot = screenshotdriver.GetScreenshot();
        password.SendKeys("Kjimenez0920!");       
        password.Submit();
        screenshot.SaveAsFile("def.png");

        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
        IWebElement btnsearch = driver.FindElement(By.Id("idBtn_Back"));
        btnsearch.Click();
         screenshot = screenshotdriver.GetScreenshot();
        screenshot.SaveAsFile("ghi.png");

        //Buscar por nombre de correo
        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
        IWebElement btnbuscar = driver.FindElement(By.Id("topSearchInput"));
        btnbuscar.Click();
        btnbuscar.SendKeys("ITLA");
        driver.FindElement(By.Id("topSearchInput")).SendKeys(Keys.Enter);
        screenshot = screenshotdriver.GetScreenshot();
        screenshot.SaveAsFile("jkl.png");

        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
        IWebElement btnhelp2 = driver.FindElement(By.Id("id__819"));
        btnhelp2.Click();
        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
        IWebElement btnnuevocorreo = driver.FindElement(By.Id("Ribbon-588"));
        btnnuevocorreo.Click();




        //Cerrar sesion
        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
        IWebElement btnperfil = driver.FindElement(By.Id("O365_MainLink_Me"));
        btnperfil.Click();
        btnperfil.SendKeys(Keys.Enter);

        IWebElement btncerrar = driver.FindElement(By.Id("mectrl_body_signOut"));
        btncerrar.Click();


        // Organizar correos
        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
        IWebElement btnsort = driver.FindElement(By.Id("mailListSortMenu"));
        btnsort.Click();
        IWebElement btnsort2 = driver.FindElement(By.ClassName("w4FE0i8yyV"));
        btnsort2.Click();







    }
}