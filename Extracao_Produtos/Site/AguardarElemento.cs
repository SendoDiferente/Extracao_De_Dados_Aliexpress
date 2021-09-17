using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Extracao_Produtos.Site
{
    public class AguardarElemento
    {
        public bool AguardarElementos(string elemento, IWebDriver driver, int i)
        {
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(i));
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(elemento)));
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(elemento)));
                Console.WriteLine("Elemento Ok");
                driver.FindElement(By.XPath(elemento)).Click();
                return true;
            }
            catch
            {
                Console.WriteLine("Elemento Ok");
                return false;
            }
        }
        public bool ElementoPresente(string elemento, IWebDriver driver, int i)
        {
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(i));
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(elemento)));
                wait.Until(ExpectedConditions.ElementExists(By.XPath(elemento)));
                return true;
            }
            catch
            {
                Console.WriteLine("Elemento Ok");
                return false;
            }
        }
    }
}
