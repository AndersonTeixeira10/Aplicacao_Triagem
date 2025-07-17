# 🔍 Verificador de Arquivos SEI - Triagem Automatizada

[![.NET](https://img.shields.io/badge/.NET-8.0-blue.svg)](https://dotnet.microsoft.com/download/dotnet/8.0)
[![Selenium](https://img.shields.io/badge/Selenium-4.32.0-green.svg)](https://selenium.dev/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## 📋 Descrição

O **Verificador de Arquivos SEI** é uma ferramenta de automação desenvolvida em C# que realiza a triagem automática de processos no Sistema Eletrônico de Informações (SEI) da ANTT. O sistema analisa documentos de processos administrativos, verifica a assinatura de autos de infração e gera relatórios detalhados para facilitar o trabalho de triagem.

## 🚀 Funcionalidades Principais

### ✨ Automação Web Completa
- **Login automático** no SEI da ANTT
- **Navegação inteligente** pelos processos
- **Extração automática** de documentos
- **Verificação de assinaturas** em autos de infração

### 📊 Análise Inteligente de Documentos
- **Classificação automática** de tipos de documentos
- **Detecção de autos de infração** (CTB e 5083)
- **Verificação de assinaturas** digitais
- **Identificação de documentos** de parcelamento, recursos, etc.

### 📈 Geração de Relatórios
- **Planilhas Excel** estruturadas e organizadas
- **Banco de dados SQLite** para consultas avançadas
- **Relatórios detalhados** com status de triagem
- **Exportação automática** para desktop

## 🛠️ Tecnologias Utilizadas

| Tecnologia | Versão | Propósito |
|------------|--------|-----------|
| **.NET 8.0** | 8.0 | Framework principal |
| **Selenium WebDriver** | 4.32.0 | Automação web |
| **ChromeDriver** | 136.0+ | Navegador automatizado |
| **ClosedXML** | 0.105.0 | Manipulação de Excel |
| **SQLite** | 9.0.7 | Banco de dados local |

## 📦 Pré-requisitos

### Sistema Operacional
- ✅ Windows 10/11 (64-bit)
- ✅ .NET 8.0 Runtime instalado

### Software Necessário
- 🌐 Google Chrome (versão 136+)
- 📊 Microsoft Excel (para visualização dos resultados)

### Acesso
- 🔑 Credenciais válidas do SEI da ANTT
- 🌐 Conexão com internet estável

## 🚀 Como Usar

### 1. Preparação da Planilha
Crie uma planilha Excel com a seguinte estrutura:
```
| Número_Processo |
|-----------------|
| 12345/2024      |
| 67890/2024      |
| ...             |
```

### 2. Execução do Programa
1. **Execute** o arquivo `Verificador de Arquivo.exe`
2. **Insira** suas credenciais do SEI
3. **Selecione** a planilha com os números dos processos
4. **Aguarde** a automação processar todos os processos

### 3. Resultados
Os resultados serão salvos automaticamente em:
```
📁 Desktop/Resultados_SEI/
├── 📄 ResultadoSEI_Triagem_YYYY-MM-DD_HH-mm.xlsx
└── 📁 TEMP/
    ├── 📄 resultado_sei_temp_YYYY-MM-DD_HH-mm.xlsx
    └── 🗄️ resultado_sei_temp_YYYY-MM-DD_HH-mm.db
```

## 📊 Estrutura dos Resultados

### Planilha Final
A planilha de resultados contém as seguintes colunas:

| Coluna | Descrição |
|--------|-----------|
| **Numero_Processo** | Número do processo analisado |
| **Tipo do Auto** | CTB ou 5083 (classificação automática) |
| **1 - Auto Infração** | Presença de auto de infração |
| **2 - Decisão de Cancelamento** | Documentos de cancelamento |
| **3 - Termo de Arquivamento** | Termos de arquivamento |
| **4 - Invalidação de Auto** | Solicitações de invalidação |
| **5.1 - Parcelamento** | Requerimentos de parcelamento |
| **5.2 - Rescisão de Parcelamento** | Rescisões de parcelamento |
| **6 - Notificação de Autuação** | Notificações de autuação |
| **7.1 - AR de Autuação** | Avisos de recebimento |
| **7.2 - Publicação DOU Autuação** | Publicações no DOU |
| **8.1 - Termo Preclusão de Defesa** | Prazos de defesa |
| **8.2 - Decisão de Análise de Defesa** | Análises de defesa |
| **8.3 - Análise de Defesa** | Documentos de defesa |
| **9 - Notificação de Penalidade/Multa** | Notificações de multa |
| **10.1 - AR de Penalidade/Multa** | ARs de penalidade |
| **10.2 - Publicação DOU Multa** | Publicações de multa |
| **11.1 - Termo Preclusão de Recurso** | Prazos de recurso |
| **11.2 - Decisão de Análise de Recurso** | Análises de recurso |
| **11.3 - Análise de Recurso** | Documentos de recurso |
| **12 - 2ª Notificação Penalidade/Final Multa** | Notificações finais |
| **13 - AR 2ª Penalidade/Final Multa** | ARs finais |
| **14.1 - DOC. Incrição SERASA** | Inscrições no SERASA |
| **14.2 - DOC. Eclusão/Erro SERASA** | Exclusões/erros SERASA |
| **15.1 - DOC. Incrição DÍVIDA** | Inscrições na Dívida Ativa |
| **15.2 - DOC. Eclusão/Erro DÍVIDA** | Exclusões/erros Dívida Ativa |
| **16.1 - DOC. Incrição CADIN** | Inscrições no CADIN |
| **16.2 - DOC. Eclusão/Erro CADIN** | Exclusões/erros CADIN |
| **17 - Outros** | Outros documentos |
| **Resultado_Fase1** | Status da verificação inicial |

## 🔧 Configuração Avançada

### Personalização de Timeouts
```csharp
// Em AutomationWeb.cs
var wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
```

### Configuração do Chrome
```csharp
// Opções do navegador
options.AddArgument("--start-maximized");
options.AddArgument("--disable-gpu");
options.AddArgument("--no-sandbox");
```

## 📁 Estrutura do Projeto

```
Verificador de Arquivo/
├── 📄 Program.cs                 # Ponto de entrada da aplicação
├── 📄 AutomationWeb.cs           # Automação web com Selenium
├── 📄 ProcessadorPlanilha.cs     # Processamento de planilhas Excel
├── 📄 ProcessadorBanco.cs        # Geração de banco SQLite
├── 📄 LoginForm.cs               # Interface de login
├── 📁 Properties/                # Recursos da aplicação
├── 📁 tessdata/                  # Dados do OCR (Tesseract)
└── 📁 bin/                       # Arquivos compilados
```

## 🐛 Solução de Problemas

### Erro de Login
```
❌ "Senha ou usuário inválido"
```
**Solução:** Verifique suas credenciais do SEI

### SEI Instável
```
❌ "SEI instável, tente quando normalizar!!"
```
**Solução:** Aguarde alguns minutos e tente novamente

### ChromeDriver Desatualizado
```
❌ "ChromeDriver version mismatch"
```
**Solução:** O WebDriverManager atualiza automaticamente

### Processo Não Encontrado
```
❌ "Nenhum documento encontrado"
```
**Solução:** Verifique se o número do processo está correto

## 🔒 Segurança

- ✅ **Credenciais não são salvas** localmente
- ✅ **Conexões HTTPS** para o SEI
- ✅ **Dados temporários** são limpos automaticamente
- ✅ **Logs não contêm** informações sensíveis

## 📈 Performance

### Tempo de Processamento
- **~2-3 segundos** por processo
- **Depende da** velocidade da internet
- **Otimizado para** processamento em lote

### Recursos Utilizados
- **RAM:** ~200-300 MB
- **CPU:** Baixo uso
- **Rede:** Conexão estável necessária

## 🤝 Contribuição

Para contribuir com o projeto:

1. **Fork** o repositório
2. **Crie** uma branch para sua feature
3. **Commit** suas mudanças
4. **Push** para a branch
5. **Abra** um Pull Request

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.

## 👥 Autores

- **Desenvolvido para** ANTT
- **Tecnologia** .NET 8.0 + Selenium
- **Versão atual** 3.0


<div align="center">

**⚡ Desenvolvido com ❤️ para otimizar a triagem de processos na ANTT**

[![ANTT](https://img.shields.io/badge/ANTT-Oficial-red.svg)](https://www.gov.br/antt/)

</div> 
