using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;

namespace Extracao_Produtos.Site
{
    public class Sites
    {
        public static string nome = Environment.UserName;
        public static ChromeDriver driver;
        public static AguardarElemento aguardar = new AguardarElemento();
        public static XLWorkbook planilha = new XLWorkbook(@$"C:\Users\{nome}\Desktop\C#\EXTRACAO DE DADOS\GC Importados\Dados\PlanilhaAli.xlsx");
        public static string Nome_do_Produto = null;
        public static string Numero_de_Vendidos;
        public static string Valor;
        public static string Nota_Produto;
        public static string Frete;
        public static string Imagem;
        public static string Link;
        public static string Area;

        public void Aliexpress(string resposta, string respostas1)
        {
            IniciarChrome();
            if (resposta == "1")
            {
                AtualizarDados();
            }
            if(respostas1 == "1")
            {
                Extracao_De_Dados();
            }
            planilha.Save();
        }

        private static void AtualizarDados()
        {
            var tabela = planilha.Worksheet(1);
            int final = tabela.LastRowUsed().RowNumber();
            for (int cont = 3; cont <= final; cont++)
            {
                string site = tabela.Cell(cont, 20).GetString();
                driver.Navigate().GoToUrl(site);
                Numero_de_Vendidos = aguardar.ExtrairDados("/html/body/div[7]/div/div[2]/div/div[2]/div[2]/span[2]", "/html/body/div[7]/div/div[2]/div/div[2]/div[2]/span[2]", "/html/body/div[7]/div/div[2]/div/div[2]/div[2]/span[2]", driver);
                Valor = aguardar.ExtrairDados($"/html/body/div[7]/div/div[2]/div/div[2]/div[4]/div[1]/span", "/html/body/div[7]/div/div[2]/div/div[2]/div[3]/div[2]/div[1]/span[1]", "/html/body/div[7]/div/div[2]/div/div[2]/div[4]/div[1]/span", driver);
                Nota_Produto = aguardar.ExtrairDados("/html/body/div[7]/div/div[2]/div/div[2]/div[2]/div/span", "/html/body/div[7]/div/div[2]/div/div[2]/div[2]/div/span", "/html/body/div[7]/div/div[2]/div/div[2]/div[2]/div/span", driver);
                string Valor_Produto1 = tabela.Cell(cont, 8).GetString();
                if (Valor_Produto1 == null)
                {
                    tabela.Cell(cont, 11).SetValue(Numero_de_Vendidos);
                    tabela.Cell(cont, 8).SetValue(DateTime.Now.ToString("dd-MM-yyyy"));
                    tabela.Cell(cont, 14).SetValue(Valor);
                    tabela.Cell(cont, 17).SetValue(Nota_Produto);
                }
                else
                {
                    tabela.Cell(cont, 12).SetValue(Numero_de_Vendidos);
                    tabela.Cell(cont, 9).SetValue(DateTime.Now.ToString("dd-MM-yyyy"));
                    tabela.Cell(cont, 15).SetValue(Valor);
                    tabela.Cell(cont, 18).SetValue(Nota_Produto);
                }
            }
        }

        public static void InserirDados(int valor)
        {
            var tabela = planilha.Worksheet(1);
            int final = tabela.LastRowUsed().RowNumber();
            bool Existente = false;
            for (int cont = 1; cont <= final; cont++)
            {
                string NomeDoProduto1 = tabela.Cell(valor, 1).GetString();
                string Valor_Produto = tabela.Cell(valor, 3).GetString();
                string Valor_Produto1 = tabela.Cell(valor, 11).GetString();
                if (Nome_do_Produto == NomeDoProduto1)
                {
                    if (Valor_Produto == null)
                    {
                        tabela.Cell(valor, 2).SetValue(Numero_de_Vendidos);
                        tabela.Cell(valor, 3).SetValue(Valor);
                        tabela.Cell(valor, 4).SetValue(Nota_Produto);
                        tabela.Cell(valor, 6).SetValue(Imagem);
                        tabela.Cell(valor, 5).SetValue(Frete);
                    }
                    else if (Valor_Produto1 == null)
                    {
                        tabela.Cell(valor, 11).SetValue(Numero_de_Vendidos);
                        tabela.Cell(valor, 8).SetValue(DateTime.Now.ToString("dd-MM-yyyy"));
                        tabela.Cell(valor, 14).SetValue(Valor);
                        tabela.Cell(valor, 17).SetValue(Nota_Produto);
                    }
                    else
                    {
                        tabela.Cell(valor, 12).SetValue(Numero_de_Vendidos);
                        tabela.Cell(valor, 9).SetValue(DateTime.Now.ToString("dd-MM-yyyy"));
                        tabela.Cell(valor, 15).SetValue(Valor);
                        tabela.Cell(valor, 18).SetValue(Nota_Produto);
                    }
                    Existente = true;
                }
            }
            if (Existente == false)
            {
                tabela.Cell(final + 1, 1).SetValue(Nome_do_Produto);
                tabela.Cell(final + 1, 2).SetValue(Numero_de_Vendidos);
                tabela.Cell(final + 1, 3).SetValue(Valor);
                tabela.Cell(final + 1, 4).SetValue(Nota_Produto);
                tabela.Cell(final + 1, 6).SetValue(Imagem);
                tabela.Cell(final + 1, 7).SetValue(DateTime.Now.ToString("dd-MM-yyyy"));
                tabela.Cell(final + 1, 5).SetValue(Frete);
                tabela.Cell(final + 1, 20).SetValue(Link);
                tabela.Cell(final + 1, 21).SetValue(Area);
            }
        }

        public static void IniciarChrome()
        {
            try
            {
                driver = new ChromeDriver();

            driver.Manage().Window.Maximize();

            }
            catch
            {

            }
        }

        public static void Extracao_De_Dados()
        {
            driver.Navigate().GoToUrl("https://pt.aliexpress.com/");
            aguardar.AguardarElementos("/html/body/div[3]/div[5]/div[1]/div/div/div[2]/div/div/div[2]/dl[4]/dt/span/a[2]", driver, 4);
            aguardar.AguardarElementos("/html/body/div[2]/div[3]/div/div/div[5]/div[3]/span/a/span", driver, 3);
            Area = "Escritorio";
            IList<IWebElement> contador = driver.FindElements(By.XPath("/html/body/div[3]/div/div/div[2]/div[2]/div/div[2]/div"));
            int i = 1;
            int j = 8;
            int k = 3;
            var rolagem_Para_Baixo = 1000;
            foreach (IWebElement cont in contador)
            {
                Nome_do_Produto = aguardar.ExtrairDados($"/html/body/div[3]/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[1]/a/span",$"/html/body/div[3]/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[4]/div/a/span", $"/html/body/div[4]/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[1]/a/span", driver);
                Numero_de_Vendidos = aguardar.ExtrairDados($"/html/body/div[3]/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[3]/a/span", $"//*[@id='root']/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[4]/a/span", $"/html/body/div[4]/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[3]/a/span", driver);
                Valor = aguardar.ExtrairDados($"//*[@id='root']/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[2]/div", $"//*[@id='root']/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[3]/div", $"/html/body/div[4]/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[2]/div[2]", driver);
                Nota_Produto = aguardar.ExtrairDados($"/html/body/div[3]/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[3]/div/a", $"//*[@id='root']/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[4]/div/a", $"/html/body/div[4]/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[3]/div/a", driver);
                Frete = aguardar.ExtrairDados($"/html/body/div[3]/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[4]/span", $"//*[@id='root']/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[5]/span[1]", $"/html/body/div[4]/div/div/div[2]/div[2]/div/div[2]/div[{i}]/div/div[4]/span", driver);
                try
                {
                    Imagem = driver.FindElement(By.XPath($"//*[@id='root']/div/div/div[2]/div[2]/div/div[2]/div[{i}]/a/img")).GetAttribute("src");
                }
                catch
                {

                }
                try
                {
                    Link = driver.FindElement(By.XPath($"//*[@id='root']/div/div/div[2]/div[2]/div/div[2]/div[{i}]/a")).GetAttribute("href");
                }
                catch
                {

                }
                InserirDados(i);
                if (i == 5)
                {
                        k = 4;
                }
                if (j == i)
                {
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript($"window.scrollBy(0,{rolagem_Para_Baixo})", "");
                    j += 8;
                    rolagem_Para_Baixo += 750;
                    planilha.Save();
                }
                Nome_do_Produto = null;
                Numero_de_Vendidos = null;
                Valor = null;
                Nota_Produto = null;
                Frete = null;
                Imagem = null;
                i++;
            }
        }
    }
}