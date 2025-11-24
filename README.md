# README.md (Português + English)

## 🇧🇷 README – Consolidador de Planilhas

### 🧩 **Descrição do Projeto**

Este repositório contém o código-fonte oficial do programa desktop desenvolvido em **Python + Tkinter** para **consolidar automaticamente planilhas financeiras da Elco Engenharia**, produzindo arquivos padronizados para uso direto no **Power BI**.

O software corrige formatações, organiza hierarquias, trata mesclas, converte datas e consolida tanto o **Orçado** quanto o **Previsto**, gerando relatórios finais limpos e consistentes.
Código-fonte: `app_visual7.py` 

---

## 🚀 **Funcionalidades Principais**

### ✔️ 1. Interface Gráfica Completa (Tkinter)

* Navegação por abas: **Orçado**, **Previsto** e **Ajuda**
* Seleção de múltiplos arquivos `.xlsx`
* Inserção de *Código Externo da Obra (CC)*
* Processamento em *thread* separada para evitar travamentos

### ✔️ 2. Correção Automática das Planilhas

* Renomeia aba ativa para “Aba1”
* Remove coluna B quando necessário
* Desfaz mesclas incorretas e reposiciona cabeçalhos
* Converte datas `mmm/aa` de PT-BR → EN para processamento

### ✔️ 3. Consolidação Inteligente – Orçado

* Reconstrução hierárquica: **Classe2**, **Classe3**, **ClasseComp**
* Filtra apenas linhas válidas com **exatamente 1 mês preenchido**
* Regras especiais para código **1030303**
* Gera arquivo final: `RESULTADO_CONSOLIDADO.xlsx`

### ✔️ 4. Consolidação Inteligente – Previsto

* Extrai verbas previstas por ClasseComp
* Gera arquivo final: `RESULTADO_PREVISTO_CONSOLIDADO.xlsx`

---

## 📦 **Instalação e Uso**

### 🔧 **1. Instalar dependências pelo requirements.txt**

Para instalar as dependências listadas no arquivo requirements.txt, execute:

```bash
pip install -r requirements.txt
```

---

### ▶️ **2. Executar o programa**

```bash
python app_visual7.py
```

---

## 🖥️ **Gerando o Executável (.exe)**

O programa pode ser transformado em um executável Windows usando o **PyInstaller**.

### **1. Instalar o PyInstaller**

```bash
pip install pyinstaller
```

### **2. Gerar o .exe com um comando simples**

Rodar no terminal, dentro da pasta do projeto:

```bash
pyinstaller --onefile --windowed app_visual7.py
```

**Explicação dos parâmetros:**

* `--onefile` → gera apenas um único .exe
* `--windowed` → remove o console preto (ideal para apps Tkinter)

O executável será criado na pasta:

```
dist/app_visual7.exe
```

Se quiser incluir um ícone:

```bash
pyinstaller --onefile --windowed --icon=icone.ico app_visual7.py
```

---

## 📂 **Saídas Geradas**

* `RESULTADO_CONSOLIDADO.xlsx`
* `RESULTADO_PREVISTO_CONSOLIDADO.xlsx`

Prontos para uso no **Power BI**.

---

## 🛠️ **Tecnologias Utilizadas**

* Python 3
* Tkinter
* Pandas
* OpenPyXL
* Threading
* Pathlib

---

## 📧 **Suporte**

Desenvolvedor: **Eric Rosa**

* [ericorosa27@gmail.com](mailto:ericorosa27@gmail.com)
* [eric.rosa@elco.com.br](mailto:eric.rosa@elco.com.br)

---

## 🏷️ **Versão**

**V7.0.2 — Novembro/2025**

---

# 🇺🇸 README – Spreadsheet Compiler

### 🧩 **Project Description**

This repository contains the official source code of a Python + Tkinter desktop application designed to automatically **clean, fix, and consolidate financial spreadsheets** used by Elco Engenharia, generating standardized reports ready for **Power BI**.

Processing includes structural correction, hierarchy rebuilding, date parsing, merged-cell handling, and consolidation of both **Budgeted** and **Forecast** spreadsheets.

---

## 🚀 **Main Features**

### ✔️ 1. Full Graphical Interface (Tkinter)

* Tab navigation: **Budgeted**, **Forecast**, and **Help**
* Multi-file selection
* External Project Code (CC) insertion
* Thread-based processing to avoid UI freezing

### ✔️ 2. Automatic Spreadsheet Correction

* Renames active sheet to “Aba1”
* Removes column B (if applicable)
* Unmerges problematic header cells
* Converts PT-BR dates `mmm/aa` → EN for processing

### ✔️ 3. Smart Consolidation – Budgeted

* Rebuilds hierarchy: **Classe2**, **Classe3**, **ClasseComp**
* Keeps only rows with **exactly one valid month value**
* Special rule handling for code **1030303**
* Output file: `RESULTADO_CONSOLIDADO.xlsx`

### ✔️ 4. Smart Consolidation – Forecast

* Extracts forecast budgets per ClasseComp
* Output file: `RESULTADO_PREVISTO_CONSOLIDADO.xlsx`

---

## 🛠️ **Technologies Used**

* **Python 3.x**
* **Tkinter**
* **Pandas**
* **OpenPyXL**
* **Threading**
* **Pathlib**

---

## 📦 **Installation**

### **1. Install dependencies from requirements.txt**

```bash
pip install -r requirements.txt
```

---

### **2. Run the application**

```bash
python app_visual7.py
```

---

## 💾 **Building the Windows Executable (.exe)**

You can generate a standalone executable using **PyInstaller**.

### **1. Install PyInstaller**

```bash
pip install pyinstaller
```

### **2. Create the .exe**

Run inside the project directory:

```bash
pyinstaller --onefile --windowed app_visual7.py
```

The executable will be generated in:

```
dist/app_visual7.exe
```

Optional with custom icon:

```bash
pyinstaller --onefile --windowed --icon=icon.ico app_visual7.py
```

---

## 📧 **Support**

Developer: **Eric Rosa**

* [ericorosa27@gmail.com](mailto:ericorosa27@gmail.com)
* [eric.rosa@elco.com.br](mailto:eric.rosa@elco.com.br)

---

## 🏷️ **Version**

**V7.0.2 — November/2025**
