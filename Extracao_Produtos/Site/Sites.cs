using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Extracao_Produtos.Site
{
    public class Sites
    {
        public static ChromeDriver driver;
        AguardarElemento aguardar = new AguardarElemento();
        public void Aliexpress()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://pt.aliexpress.com/");
            aguardar.AguardarElementos("/html/body/div[3]/div[5]/div[1]/div/div/div[2]/div/div/div[2]/dl[4]/dt/span/a[2]", driver, 4);
            IWebElement ContagemdeDivs = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/div[2]/div/div[2]"));
            IList<IWebElement> contador = driver.FindElements(By.XPath("div"));
            int i = 0;
            foreach(IWebElement cont in contador)
            {
                string Nome_do_Produto = driver.FindElement(By.XPath("/html/body/div[3]/div/div/div[2]/div[2]/div/div[2]/div[1]/div/div[1]/a/span")).Text;
                string Numero_de_Vendidos =  driver.
            }
        }
    }
}
