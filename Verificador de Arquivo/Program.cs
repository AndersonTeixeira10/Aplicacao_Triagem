using System;
using System.IO;
using System.Windows.Forms;

namespace Verificador_de_Arquivo
{
    internal class Program
    {
        private static string _pastaDados;

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            CriarPastaDados();

            Console.WriteLine($"📁 Pasta de dados: {_pastaDados}");
            Console.WriteLine($"📁 Pasta existe: {Directory.Exists(_pastaDados)}");

            using var loginForm = new LoginForm();
            if (loginForm.ShowDialog() != DialogResult.OK)
                return;

            var usuario = loginForm.Usuario;
            var senha = loginForm.Senha;

            string arquivoSelecionado;

            using (var dlg = new OpenFileDialog
            {
                Title = "Selecione a planilha com a coluna Número_Processo",
                Filter = "Excel (*.xlsx)|*.xlsx|Todos (*.*)|*.*"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                arquivoSelecionado = dlg.FileName;
            }

            try
            {
                var processador = new ProcessadorPlanilha(arquivoSelecionado, _pastaDados);

                // 1. Executa e pega o caminho do arquivo TEMP
                string caminhoExcelTemp = processador.Processar(usuario, senha);

                // 2. Caminho para a pasta principal de resultados
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string pastaResultados = Path.Combine(desktopPath, "Resultados_SEI");
                if (!Directory.Exists(pastaResultados))
                    Directory.CreateDirectory(pastaResultados);

                // 3. Nome do arquivo final
                string dataHora = DateTime.Now.ToString("yyyy-MM-dd_HH-mm");
                string nomeArquivoFinal = Path.Combine(
                    pastaResultados,
                    $"ResultadoSEI_Triagem_{dataHora}.xlsx"
                );

                // 4. Gera o arquivo final
                var processadorBanco = new ProcessadorBanco(caminhoExcelTemp);
                processadorBanco.CriarBancoEDefinirStatus(nomeArquivoFinal);

                // 5. Mostra mensagem de sucesso
                MessageBox.Show(
                    $"✅ Verificação concluída!\n\n📁 Resultado salvo em:\n{nomeArquivoFinal}",
                    "Sucesso",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                // 6. Tenta excluir o arquivo Excel temporário (mantém a pasta TEMP e o banco .db)
                try
                {
                    if (File.Exists(caminhoExcelTemp))
                    {
                        File.Delete(caminhoExcelTemp);
                        Console.WriteLine("🗑️ Arquivo Excel TEMP excluído com sucesso.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"⚠️ Erro ao excluir o arquivo Excel temporário:\n{ex.Message}", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro inesperado: {ex.Message}", "Automação", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void CriarPastaDados()
        {
            string areaDeTrabalho = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            _pastaDados = Path.Combine(areaDeTrabalho, "Dados_Imagens");

            try
            {
                if (!Directory.Exists(_pastaDados))
                {
                    Directory.CreateDirectory(_pastaDados);
                    Console.WriteLine($"Pasta criada em: {_pastaDados}");
                }
                else
                {
                    Console.WriteLine($"Pasta já existe em: {_pastaDados}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao criar pasta: {ex.Message}", "Erro",
                               MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
