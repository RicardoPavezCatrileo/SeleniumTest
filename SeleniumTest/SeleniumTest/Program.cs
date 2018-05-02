using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;


namespace SeleniumTest
{
    class Program
    {
        static void Main(string[] args)
        {
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://dutyfree.custhelp.com/");
            driver.Manage().Window.Maximize();

            IWebElement FPregunta = driver.FindElement(By.Id("rn_NavigationTab_5"));
            FPregunta.Click();

            IWebElement DirecccionEmail = driver.FindElement(By.Id("rn_TextInput_28_Contact.Emails.PRIMARY.Address"));
            DirecccionEmail.SendKeys("ricardoalfredo.pavez@gmail.com");
            //rn_TextInput_28_Contact.Emails.PRIMARY.Address

            IWebElement Asunto = driver.FindElement(By.Id("rn_TextInput_30_Incident.Subject"));
            Asunto.SendKeys("PRUEBA DE QA");
            //rn_TextInput_30_Incident.Subject

            IWebElement Pregunta = driver.FindElement(By.Id("rn_TextInput_32_Incident.Threads"));
            Pregunta.SendKeys("Como puedo hacer una devolución");
            //rn_TextInput_32_Incident.Threads

            IWebElement EnvieSupregunta = driver.FindElement(By.Id("rn_FormSubmit_36_Button"));
            EnvieSupregunta.Click();
            //rn_FormSubmit_36_Button

            Thread.Sleep(5000);

            driver.Close();

        }
    }
}
