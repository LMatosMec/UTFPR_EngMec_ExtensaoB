# Automação de Extração e Mapeamento de Custeio de Máquinas Industriais (SAP ERP + Python)

Projeto de Extensão Tecnológica desenvolvido no curso de **Engenharia Mecânica da Universidade Tecnológica Federal do Paraná (UTFPR – Campus Curitiba / DAMEC)**, no âmbito das Disciplinas Extensionistas (`MEC7EX1-S41` / `MEC7EX1-S43`).

---

## Sobre o Projeto

No ambiente industrial, o fechamento contábil e a auditoria de custeio de máquinas e centros de trabalho em sistemas ERP (como o SAP) frequentemente dependem de extrações manuais repetitivas e demoradas, gerando gargalos operacionais e maior propensão a erros no tratamento de grandes volumes de dados.

Esta ferramenta de código aberto aplica conceitos de **Engenharia de Processos**, **Análise de Custos Industriais**, **Lean Office** e **Programação em Python** para automatizar o ciclo completo de **ETL (Extração, Transformação e Carga)** diretamente integrado ao **SAP GUI**:

* **Extração Paralela em Lotes (Multithreading):** Abre e gerencia um *pool* de sessões simultâneas do SAP GUI via API COM (`pywin32`), respeitando o limite de janelas abertas e encerrando apenas as janelas criadas pelo robô ao final da execução.
* **Arquitetura Idempotente (Resiliência a Falhas):**
  * Remove automaticamente arquivos residuais de dias anteriores, mantendo os arquivos já baixados na data atual.
  * Filtra a fila de extração para pular itens já atualizados no disco.
  * Consolida as planilhas de forma incremental, validando a consistência dos cabeçalhos e adicionando rastreabilidade pela coluna `arquivo_origem`.
* **Transações SAP Integradas:**
  * **`CK51N` / `CK53N`:** Extração em lote de estruturas de custeio de ordens/máquinas (M-SD).
  * **`SE16N`:** Motor autônomo para extração de tabelas com aplicação dinâmica de filtros (incluindo seleção múltipla via área de transferência), rolagem automática da grade e seleção de layout.
* **Pós-processamento e Formatação Automática:** Fecha janelas residuais do Microsoft Excel abertas durante a exportação do SAP (`win32gui`) e entrega a planilha consolidada já formatada com `openpyxl` (cabeçalho em negrito, autofiltro, centralização, autoajuste de largura das colunas e ocultação de linhas de grade).

---

## Estrutura do Repositório

```text
UTFPR_EngMec_ExtensaoB/
├── run.py               # Orquestrador principal do fluxo ETL (Setup -> Extract -> Transform -> Load)
├── config.py            # Parâmetros centralizados (diretório, layout SAP, nome do arquivo e lista de ordens)
├── conexao.py           # Gerenciamento de sessões SAP GUI via COM (abertura segura, detecção de deadlock e fechamento)
├── motor.py             # Executor paralelo em lotes (ThreadPoolExecutor + pythoncom) com reaproveitamento de janelas
├── transacoes_sap.py    # Rotinas de navegação e exportação nas transações CK51N/CK53N e SE16N
├── funcoes.py           # Utilitários de idempotência, controle de Excel, consolidação (Pandas) e formatação (OpenPyXL)
├── requirements.txt     # Dependências do projeto com versões fixadas
├── ruff.toml            # Configurações de padronização e lint de código
└── .gitignore           # Exclusão de ambiente virtual, cache e planilhas locais
```

---

## Pré-requisitos

Antes de iniciar a instalação, verifique se o computador atende aos requisitos abaixo:

1. **Sistema Operacional Windows:** Obrigatório, pois a automação utiliza a interface COM nativa do Windows (`pywin32` / `win32gui`) para comunicação com o SAP GUI e o Microsoft Excel.
2. **SAP Logon (SAP GUI for Windows):** Instalado, autenticado e com o **SAP GUI Scripting habilitado** (ver seção de configuração abaixo).
3. **Microsoft Excel:** Instalado na máquina (utilizado na exportação padrão de planilhas XXL do SAP).
4. **Python 3.10 ou superior:** Disponível em [python.org/downloads](https://www.python.org/downloads/).
5. **Visual Studio Code (VS Code):** Disponível em [code.visualstudio.com](https://code.visualstudio.com/).
6. **Git (Opcional):** Necessário apenas caso deseje clonar o repositório via terminal ([git-scm.com](https://git-scm.com/)).

---

## Guia de Instalação e Configuração (Python + VS Code)

Este passo a passo foi estruturado para permitir que qualquer usuário configure o ambiente do zero sem conflitos de bibliotecas.

### 1. Instalar o Python e o VS Code
1. Ao executar o instalador do **Python**, marque obrigatoriamente a caixa **"Add python.exe to PATH"** na tela inicial antes de clicar em *Install Now*.
2. Instale o **VS Code**. Ao abri-lo, vá à aba de **Extensões** (`Ctrl + Shift + X`), pesquise por **Python** (publicada pela Microsoft) e clique em **Install**.

### 2. Baixar o Projeto
Você pode obter o código de duas formas:

* **Via Git (recomendado):** Abra o terminal na pasta desejada e execute:
  ```bash
  git clone [https://github.com/LMatosMec/UTFPR_EngMec_ExtensaoB.git](https://github.com/LMatosMec/UTFPR_EngMec_ExtensaoB.git)
  ```
* **Download direto:** Clique no botão verde **Code** no topo deste repositório > **Download ZIP** e extraia o conteúdo em uma pasta local (ex: `C:\Projetos\Extensao_B`).

### 3. Abrir o Projeto e Criar o Ambiente Virtual (`venv`)
1. No VS Code, vá em **File > Open Folder** e selecione a pasta do projeto.
2. Abra o terminal integrado pressionando `Ctrl + '` (ou no menu superior: **Terminal > New Terminal**).
3. Crie o ambiente virtual isolado executando:
   ```bash
   python -m venv venv
   ```
4. Ative o ambiente virtual:
   ```bash
   .\venv\Scripts\activate
   ```
   > **Nota (Windows PowerShell):** Caso o PowerShell bloqueie a ativação com um erro de política de execução de scripts, rode o comando abaixo uma única vez e tente ativar novamente:
   > ```powershell
   > Set-ExecutionPolicy Unrestricted -Scope CurrentUser
   > ```
   Quando ativado corretamente, o prefixo **`(venv)`** aparecerá no início da linha do terminal.

### 4. Vincular o Interpretador Python no VS Code
1. Pressione `Ctrl + Shift + P` para abrir a paleta de comandos.
2. Digite **`Python: Select Interpreter`** e pressione Enter.
3. Selecione a opção que aponta para a pasta `venv` dentro do projeto (geralmente indicada como `('venv': venv)`).

### 5. Instalar as Bibliotecas
Com o prefixo `(venv)` ativo no terminal do VS Code, atualize o `pip` e instale todas as dependências listadas no `requirements.txt`:

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Principais bibliotecas instaladas:
* `pywin32`: Comunicação COM com o SAP GUI e controle de janelas do Windows.
* `pandas` e `numpy`: Leitura, limpeza e consolidação vetorizada das bases extraídas.
* `openpyxl`: Formatação visual automatizada da planilha consolidada final.
* `pyperclip`: Manipulação da área de transferência para preenchimento em massa de filtros no SAP.

---

## Como Configurar e Executar

### 1. Habilitar o Scripting no SAP GUI
Para que o Python consiga interagir com as telas do SAP:
1. Abra o **SAP Logon** e faça login normalmente.
2. Clique no ícone de **Personalizar Layout Local** (`Alt + F12`) no canto superior direito e acesse **Opções...**.
3. Navegue até **Acessibilidade & Scripting > Scripting**.
4. Marque a opção **Ativar scripting** (*Enable scripting*).
5. Desmarque as opções **"Notificar quando um script se conectar ao SAP GUI"** e **"Notificar quando um script abrir uma conexão"** para evitar pop-ups que pausam a automação.
6. Clique em **OK**. *(Nota: O parâmetro `sapgui/user_scripting` também precisa estar habilitado pelo time de Basis/TI no servidor SAP).*

### 2. Ajustar os Parâmetros em `config.py`
Abra o arquivo `config.py` no VS Code e substitua os valores de exemplo pelos dados do seu ambiente:

```python
class MSD:
    # Caminho da pasta local onde os arquivos extraídos e o consolidado serão salvos
    PATH = r"C:\Caminho\Para\Sua\Pasta\De\Saida"

    # Nome do layout salvo no SAP (quando aplicável nas extrações da SE16N)
    LAYOUT_MSD = "NOME_DO_SEU_LAYOUT"

    # Prefixo do arquivo Excel consolidado final
    NOME_MSD = "Estrutura_Maquinas_Consolidada"

    # Lista de ordens/documentos (M-SD) a serem extraídos
    LISTA_MSD = [
        "1234567891",
        "2345678912",
    ]
```

> **Dica de Segurança:** Nunca versione no Git arquivos com números de ordens reais, planilhas ou dados corporativos sensíveis. O arquivo `.gitignore` já bloqueia automaticamente arquivos `.xlsx`, `.csv` e `.env`.

### 3. Executar a Automação
1. Certifique-se de que há **pelo menos uma janela do SAP GUI aberta, logada e livre** (sem transações em processamento).
2. No terminal do VS Code com o `(venv)` ativo, execute:
   ```bash
   python run.py
   ```
3. Ao término da execução, o terminal exibirá o resumo de linhas consolidadas e o tempo total. O relatório final formatado será salvo na pasta definida em `MSD.PATH` com o padrão:
   ```text
   Estrutura_Maquinas_Consolidada_DD_MM_AAAA.xlsx
   ```

---

## Solução de Problemas Comuns

| Sintoma / Erro | Causa Provável | Solução |
| :--- | :--- | :--- |
| Sublinhado amarelo em `import pythoncom` no VS Code | O VS Code não está apontando para o interpretador do `venv` ou o `pywin32` não foi instalado. | Execute `pip install pywin32` no `(venv)`, selecione o interpretador via `Ctrl + Shift + P` > `Python: Select Interpreter` e recarregue a janela (`Developer: Reload Window`). |
| `O SAP já atingiu o limite de janelas` | Muitas sessões manuais abertas simultaneamente no SAP GUI. | Feche sessões ociosas do SAP ou reduza o parâmetro `tamanho_lote_desejado` em `run.py`. |
| `PermissionError` ao consolidar ou apagar arquivo | A planilha de saída ou um dos arquivos extraídos está aberto no Microsoft Excel. | Feche as planilhas abertas no Excel e rode `python run.py` novamente (a idempotência retomará de onde parou). |
| `Todas as sessões do SAP estão ocupadas` | A sessão mestre do SAP está executando outra tarefa no momento da chamada. | Aguarde a conclusão da tela atual no SAP ou abra uma sessão livre na tela inicial (`/n`). |

---

## Vídeo Tutorial de Implementação

> *[Inserir aqui o link do vídeo demonstrativo de instalação e execução do projeto]*

---

## Autores e Créditos Institucionais

* **Desenvolvedores:**
  * Luis Gustavo Godoy Matos
  * Lucas Augusto Peppes do Valle Torres
* **Instituição:** Universidade Tecnológica Federal do Paraná (UTFPR – Campus Curitiba) — Departamento Acadêmico de Mecânica (DAMEC)
* **Disciplinas:** Extensão B (`MEC7EX1-S41` / `MEC7EX1-S43`) — 2026
* **Responsável Institucional:** Prof. Dr. Eng. Tiago Rodrigues Weller