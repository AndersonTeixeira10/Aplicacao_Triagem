using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager.Helpers;

namespace Verificador_de_Arquivo
{
    public class AutomationWeb
    {
        public IWebDriver Driver { get; private set; }

        public AutomationWeb()
        {
            new DriverManager().SetUpDriver(new ChromeConfig(), VersionResolveStrategy.MatchingBrowser);
            var service = ChromeDriverService.CreateDefaultService();
            var options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            Driver = new ChromeDriver(service, options, TimeSpan.FromMinutes(10));
        }

        public bool GoToSEI()
        {
            Driver.Navigate().GoToUrl("https://portal.antt.gov.br/sei");
            return WaitForPageLoadWithTimeout(TimeSpan.FromMinutes(3));
        }

        public void ClickUsuarioInterno()
        {
            var wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
            var link = wait.Until(ExpectedConditions.ElementToBeClickable(
                By.CssSelector("a[title='Acesso para Acessar SEI (Usuário Interno)']")));
            link.Click();
        }

        public void SwitchToNewWindow()
        {
            var wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
            wait.Until(d => d.WindowHandles.Count >= 2);
            Driver.SwitchTo().Window(Driver.WindowHandles.Last());
        }

        public void Login(string usuario, string senha)
        {
            var wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
            var campoUsuario = wait.Until(d => d.FindElement(By.Id("txtUsuario")));
            campoUsuario.Clear();
            campoUsuario.SendKeys(usuario);

            var campoSenha = wait.Until(d => d.FindElement(By.Id("pwdSenha")));
            campoSenha.Clear();
            campoSenha.SendKeys(senha);

            var botao = wait.Until(d => d.FindElement(By.CssSelector("button[type='submit']")));
            botao.Click();
        }

        public void Quit() => Driver.Quit();

        private bool WaitForPageLoadWithTimeout(TimeSpan timeout)
        {
            var js = (IJavaScriptExecutor)Driver;
            var sw = System.Diagnostics.Stopwatch.StartNew();

            while (sw.Elapsed < timeout)
            {
                try
                {
                    if (js.ExecuteScript("return document.readyState")?.ToString() == "complete")
                        return true;
                }
                catch
                {
                }
                Thread.Sleep(200);
            }
            return false;
        }

        /// <summary>
        /// Extrai todos os nomes dos documentos do processo atual.
        /// Pré-requisito: já estar no contexto do iframe da árvore!
        /// </summary>
        public List<string> ExtrairNomesDocumentos()
        {
            var nomes = new List<string>();

            // Já estamos no iframe!
            var links = Driver.FindElements(By.CssSelector("a.infraArvoreNo"));

            foreach (var link in links)
            {
                var span = link.FindElement(By.TagName("span"));
                nomes.Add(span.Text.Trim());
            }

            return nomes;
        }

        /// <summary>
        /// Verifica se o Auto de Infração do processo está assinado.
        /// Pré-requisito: já estar no contexto do iframe da árvore!
        /// </summary>
        public string VerificarAutoAssinado()
        {
            // Já estamos no iframe!
            Thread.Sleep(1000);

            // Busca todos os links de Auto de Infração
            var autoLinks = Driver.FindElements(By.XPath("//a[span[contains(text(), 'Auto de Infração')]]"));

            if (autoLinks.Count == 0)
            {
                return "VERIFICAR MANUALMENTE"; // Não encontrou auto de infração
            }

            // Verifica se algum Auto de Infração NÃO está assinado (tem style no span)
            foreach (var link in autoLinks)
            {
                try
                {
                    var span = link.FindElement(By.TagName("span"));
                    string style = span.GetAttribute("style");
                    if (!string.IsNullOrEmpty(style))
                    {
                        return "NÃO APTO - AUTO SEM ASSINATURA";
                    }
                }
                catch
                {
                    // Se não achar o span, ignora
                }
            }

            return "SEGUIR";
        }
    }
}
