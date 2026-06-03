# 🤖 RPA FIAPE - Automação de Envio de Planejamento

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Processing-150458.svg)
![Playwright](https://img.shields.io/badge/Playwright-Web%20Automation-2EAD33.svg)
![Tkinter](https://img.shields.io/badge/Tkinter-GUI-lightgrey.svg)

Este projeto é uma ferramenta de **Automação de Processos Robóticos (RPA)** desenvolvida em Python. O objetivo principal é automatizar o processamento de planilhas de planejamento de coletas (operações DHL / Stellantis) e enviar e-mails formatados automaticamente através do **Outlook Web**.

## ✨ Funcionalidades

- 📊 **Processamento de Dados**: Lê planilhas de Planejamento e mapeia automaticamente com a base de e-mails (Fornecedores, Transportadoras e equipe Stellantis).
- 📧 **Geração Automática de E-mails**: Agrupa dados por `SAP` + `Fornecedor` e cria tabelas HTML formatadas para o corpo do e-mail.
- 🌐 **Automação Web (Playwright)**: Controla o navegador (Edge ou Chrome) para fazer login no Outlook Web, preencher destinatários (TO e CC), assunto, inserir o corpo do e-mail e enviar.
- 🖥️ **Interface Gráfica (GUI)**: Interface moderna construída com Tkinter, contendo barra de progresso, logs em tempo real e seletor de modo de execução.
- 🧪 **Modo de Teste / Produção**: 
  - **Modo Teste**: Gera os e-mails com dados reais, mas os envia apenas para um e-mail de teste seguro definido nas configurações.
  - **Modo Produção**: Envia os e-mails para os contatos reais mapeados nas planilhas.
- 🔐 **Gestão de Sessão**: Utiliza o perfil local do navegador (Edge/Chrome) para manter as credenciais do usuário e pular a etapa de login (quando possível).

---

## 📁 Estrutura de Diretórios Esperada

Para que o script funcione corretamente, o projeto exige uma estrutura de pastas e arquivos base. Certifique-se de que a estrutura abaixo exista na mesma pasta do executável ou script:

```text
RPA FIAPE/
│
├── Base/
│   ├── Dados_email/
│   │   └── (Arquivo .xlsx contendo abas: GERAL, TRANSPORTADORAS, STELLANTIS, FORNECEDORES)
│   └── Planejamento/
│       └── (Arquivo .xlsx de planejamento contendo a aba 'Planejado' e indicando a semana no nome, ex: w42)
│
├── App.py               # Arquivo principal (Interface Gráfica e Automação Playwright)
├── Tasks.py             # Lógica de processamento de dados (Pandas) e HTML
├── .env                 # Arquivo de configuração de variáveis de ambiente
├── requirements.txt     # Dependências do projeto
└── README.md
```

---

## ⚙️ Configurações Iniciais (`.env`)

Crie um arquivo chamado `.env` na raiz do projeto contendo as seguintes variáveis:

```env
# Define o e-mail que receberá os envios quando o "Modo de Teste" estiver ativado.
# Pode ser separado por vírgulas para múltiplos e-mails.
EMAIL_TESTE=seu.email@exemplo.com

# Navegador a ser utilizado pela automação (edge ou chrome)
Nav=edge

# URL do Outlook Web
OUTLOOK_URL=https://outlook.office.com/mail/
```

---

## 🚀 Instalação e Execução (Para Desenvolvedores)

### 1. Pré-requisitos
- **Python 3.8** ou superior.
- Navegador Microsoft Edge ou Google Chrome instalado na máquina.

### 2. Instalar Dependências
Abra o terminal na pasta do projeto e instale as bibliotecas necessárias:

```bash
pip install -r requirements.txt
```

### 3. Instalar os Navegadores do Playwright
O Playwright precisa baixar seus binários de automação para funcionar corretamente:

```bash
playwright install chromium
```

### 4. Executar o Projeto
Execute o arquivo principal:

```bash
python App.py
```

---

## 📦 Compilando para Executável (.exe)

Este projeto já foi adaptado (usando `sys._MEIPASS` e `sys.frozen`) para ser compilado com **PyInstaller**.

Para gerar o executável sem abrir o terminal em background e incluindo o Playwright corretamente, recomenda-se um comando similar a este:

```bash
pyinstaller --noconsole  --name="RPA - FIAPE Envio de Panejamento" --icon="icon.ico"   --hidden-import playwright --hidden-import playwright.sync_api --hidden-import dotenv App.py


```
*(Atenção: A compilação com Playwright pode requerer a inclusão manual da pasta `ms-playwright` usando o argumento `--add-data` do PyInstaller, dependendo do ambiente).*

---

## ⚠️ Avisos Importantes

1. **Fechamento do Navegador**: O robô tenta usar a sessão do seu Edge/Chrome pessoal para não exigir login recorrente. Para isso funcionar, o robô fechará processos em segundo plano do navegador selecionado antes de iniciar.
2. **Formato das Planilhas**: O sistema é *case-insensitive* para o cabeçalho das planilhas, mas exige a existência de colunas específicas (`SAP`, `Fornecedor`, `Veículos`, `Destino`, `Semana`, `Dia`, `Data Planejada de Coleta`, `Hora Planejada de Coleta`).
3. **Tempos de Espera (Timeouts)**: A automação possui tempos de espera (`page.wait_for_timeout`) programados para simular o comportamento humano e evitar bloqueios pelo sistema do Outlook Web. Não minimize abruptamente a janela enquanto ele clica e digita.

---
**Desenvolvido por:** Vincent Pernarh