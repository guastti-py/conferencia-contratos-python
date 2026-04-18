# ConferenciApp 📊

![Tela do sistema](img_conferencia_app.png)

> Aplicação desktop desenvolvida em Python para automatização de conciliação financeira entre dados do sistema interno (Excel/CSV) e arquivos dos parceiros (PDF/CSV).

---

## 💡 Contexto

Em operações financeiras com múltiplos parceiros de crédito, é comum receber arquivos de diferentes formatos (Excel, CSV, PDF) que precisam ser comparados com os dados do sistema interno — processo chamado de **conciliação bancária**.

Este sistema automatiza esse fluxo, eliminando comparações manuais, reduzindo erros e acelerando o envio de relatórios para a tesouraria.

---

## ✨ Funcionalidades

- **Identificação de parceiro por código** — interface simples, sem necessidade de navegar em menus
- **Geração automática de resumo financeiro** — totais de Valor Bruto, Valor Líquido, IOF, e outros campos
- **Conferência Excel × CSV** — compara operações do sistema com arquivo do parceiro, identifica divergências
- **Conferência Excel × PDF** — extrai contratos de PDFs com `pdfplumber` e cruza com dados do sistema
- **Detecção de contratos faltando** — gera aba automática com contratos presentes no PDF mas ausentes no sistema
- **Exportação formatada** — arquivos `.xlsx` prontos para envio à tesouraria, com formatação condicional automática
- **Barra de progresso** — feedback visual durante o processamento

---

## 🛠️ Tecnologias utilizadas

| Tecnologia | Uso |
|---|---|
| `Python 3.x` | Linguagem principal |
| `CustomTkinter` | Interface gráfica moderna |
| `pandas` | Leitura e tratamento de dados (CSV/Excel) |
| `openpyxl` | Criação e formatação de planilhas Excel |
| `pdfplumber` | Extração de dados de arquivos PDF |
| `Pillow` | Manipulação de imagens na interface |

---

## 📁 Estrutura do projeto

```
ConferenciApp/
│
├── main.py                    # Ponto de entrada da aplicação
│
├── config/
│   └── parceiros.py           # Tabela de códigos e regras dos parceiros
│
├── interface/
│   ├── janela.py              # Interface gráfica principal (CustomTkinter)
│   └── acoes.py               # Lógica dos botões e fluxo geral
│
├── parceiros/
│   ├── geral.py               # Processamento base (resumo, exportação)
│   ├── parceiro_a.py          # Módulo: conferência Excel × CSV
│   ├── parceiro_b.py          # Módulo: conferência Excel × CSV + PDF (Endosso)
│   ├── parceiro_c.py          # Módulo: conferência multi-PDF com Endosso
│   ├── parceiro_d.py          # Módulo: conferência Excel × PDF com cálculo de ágio
│   ├── parceiro_e.py          # Módulo: conferência Excel × PDF (4 colunas)
│   └── ...
│
└── utils/
    ├── arquivos.py            # Leitura de arquivos, validação de datas, conversão
    └── excel.py               # Formatação de planilhas, largura de colunas, formatação condicional
```

---

## 🚀 Como executar

### Pré-requisitos

```bash
pip install customtkinter pandas openpyxl pdfplumber pillow
```

### Executar

```bash
python main.py
```

---

## 🖥️ Como funciona

1. O usuário digita o **código do parceiro** na interface
2. O sistema identifica o parceiro e exibe o nome correspondente
3. O usuário informa a **data da operação**
4. Clica em **Anexar e Gerar** para selecionar o arquivo do sistema (Excel/CSV)
5. O sistema processa, gera o resumo financeiro e salva o arquivo formatado na mesma pasta
6. Opcionalmente, pode rodar a **Conferência** para cruzar com o arquivo do parceiro e detectar divergências

---

## 📌 Destaques técnicos

- Leitura inteligente de CSV com detecção automática de separador e encoding (`utf-8` / `latin-1`)
- Extração de tabelas de PDFs com múltiplas páginas usando `pdfplumber`
- Normalização automática de CPF/CNPJ para comparação entre fontes diferentes
- Formatação condicional no Excel: células verdes para "CONFERE", vermelhas para "NÃO CONFERE"
- Tratamento de erros com mensagens amigáveis ao usuário

---

## 👤 Autor: Gabriel Guastti de Almeida

Projeto desenvolvido com foco em automação de processos financeiros, redução de erros operacionais e ganho de eficiência.  
Portfólio: (https://github.com/guastti-py)  
LinkedIn: (https://www.linkedin.com/in/gabriel-guastti-de-almeida-analista-de-dados/)
