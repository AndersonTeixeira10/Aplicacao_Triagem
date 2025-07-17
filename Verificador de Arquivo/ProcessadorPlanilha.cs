using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using OpenQA.Selenium;

namespace Verificador_de_Arquivo
{
    public class ProcessadorPlanilha
    {
        private readonly string _arquivoSelecionado;
        private readonly string _pastaDados;

        public ProcessadorPlanilha(string arquivoSelecionado, string pastaDados)
        {
            _arquivoSelecionado = arquivoSelecionado;
            _pastaDados = pastaDados;
        }

        public string Processar(string usuario, string senha)
        {
            var resultados = new List<string[]>();
            resultados.Add(new[] { "Numero_Processo", "Resultado_Fase1", "Documento" });

            var numerosProcesso = new List<string>();
            using (var wb = new XLWorkbook(_arquivoSelecionado))
            {
                var ws = wb.Worksheets.First();
                int colNum = ws.Row(1)
                               .CellsUsed()
                               .First(c => c.GetString().Trim() == "Número_Processo")
                               .Address.ColumnNumber;
                int ultimaLinha = ws.LastRowUsed().RowNumber();

                for (int i = 2; i <= ultimaLinha; i++)
                {
                    var numProc = ws.Cell(i, colNum).GetString().Trim();
                    if (!string.IsNullOrWhiteSpace(numProc))
                        numerosProcesso.Add(numProc);
                }
            }

            var automation = new AutomationWeb();
            var driver = automation.Driver;

            try
            {
                if (!automation.GoToSEI())
                {
                    automation.Quit();
                    throw new Exception("SEI instável, tente quando normalizar!!");
                }

                automation.ClickUsuarioInterno();
                automation.SwitchToNewWindow();
                automation.Login(usuario, senha);

                // Alert de login inválido
                try
                {
                    var wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    var alert = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.AlertIsPresent());
                    if (alert.Text.Contains("Usuário ou Senha Inválida", StringComparison.OrdinalIgnoreCase))
                    {
                        alert.Accept();
                        automation.Quit();
                        throw new Exception("Senha ou usuário inválido.");
                    }
                    alert.Accept();
                }
                catch (OpenQA.Selenium.WebDriverTimeoutException) { }

                // Fecha tela inicial e volta para a janela principal
                var originalHandle = driver.WindowHandles[0];
                driver.Close();
                driver.SwitchTo().Window(originalHandle);

                automation.ClickUsuarioInterno();
                automation.SwitchToNewWindow();

                foreach (var numeroProcesso in numerosProcesso)
                {
                    string resultado = "";

                    try
                    {
                        var wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(10));
                        var campo = driver.FindElement(By.Id("txtPesquisaRapida"));
                        campo.Clear();
                        campo.SendKeys(numeroProcesso + OpenQA.Selenium.Keys.Enter);

                        wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.FrameToBeAvailableAndSwitchToIt("ifrArvore"));

                        try
                        {
                            var elementos = driver.FindElements(By.CssSelector("#joinPASTA1"));
                            if (elementos.Count > 0)
                            {
                                elementos[0].Click();
                                var waitDoc = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(15));
                                waitDoc.Until(d =>
                                {
                                    var aguarde = d.FindElements(By.XPath("//*[contains(text(),'Aguarde')]"));
                                    return aguarde.Count == 0;
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Erro ao tentar expandir pasta: " + ex.Message);
                        }

                        var nomesArquivos = automation.ExtrairNomesDocumentos();
                        resultado = automation.VerificarAutoAssinado();
                        driver.SwitchTo().DefaultContent();

                        if (nomesArquivos.Count == 0)
                        {
                            resultados.Add(new[] { numeroProcesso, resultado, "Nenhum documento encontrado" });
                        }
                        else
                        {
                            bool primeiraLinha = true;
                            foreach (var nome in nomesArquivos)
                            {
                                resultados.Add(new[]
                                {
                                    numeroProcesso,
                                    primeiraLinha ? resultado : "",
                                    nome
                                });
                                primeiraLinha = false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        resultados.Add(new[] { numeroProcesso, "ERRO: " + ex.Message, "" });
                    }
                }

            }
            finally
            {
                driver.Quit();
            }

            // --- Gera Excel Final na subpasta TEMP ---
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string pastaResultados = Path.Combine(desktopPath, "Resultados_SEI");
            if (!Directory.Exists(pastaResultados))
                Directory.CreateDirectory(pastaResultados);

            // Cria a subpasta TEMP dentro de Resultados_SEI
            string pastaTemp = Path.Combine(pastaResultados, "TEMP");
            if (!Directory.Exists(pastaTemp))
                Directory.CreateDirectory(pastaTemp);

            string dataHora = DateTime.Now.ToString("yyyy-MM-dd_HH-mm");
            string nomeArquivoExcel = Path.Combine(pastaTemp, $"resultado_sei_temp_{dataHora}.xlsx");

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Resultados");

                // Cabeçalhos
                worksheet.Cell(1, 1).Value = "Numero_Processo";
                worksheet.Cell(1, 2).Value = "Resultado_Fase1";
                worksheet.Cell(1, 3).Value = "Documento";

                // Dados
                for (int i = 1; i < resultados.Count; i++)
                {
                    var linha = resultados[i];
                    worksheet.Cell(i + 1, 1).Value = linha[0];
                    worksheet.Cell(i + 1, 2).Value = linha[1];
                    worksheet.Cell(i + 1, 3).Value = linha[2];
                }

                worksheet.Columns().AdjustToContents();
                workbook.SaveAs(nomeArquivoExcel);
            }

            // Retorna o caminho do arquivo Excel gerado na subpasta TEMP
            return nomeArquivoExcel;
        }
    }
}
