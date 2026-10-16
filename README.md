# ✦ Desktop Archaeology Toolkit ✦

> **Two tiny Python scripts for turning a chaotic folder into something that looks like you have your life together.**
>
> 🗂️ inventory first → 🧠 classify → 🧹 organize → ✨ pretend it was always this way

This repository contains two complementary utilities for **exploring, understanding, and organizing large folders** — especially a Windows Desktop that has slowly evolved into an ecosystem.

They are intentionally simple, local, and boring in the best possible way.

No cloud uploads.  
No mysterious telemetry.  
No "AI-powered productivity ecosystem".  
Just Python, files, and a suspicious amount of `Path()`.

---

## 🌸 What's inside?

### `file_inventory_maker.py`

The **reconnaissance girlie**.

It recursively scans the folder where the script lives and creates:

```text
folder_inventory.txt
```

The inventory contains:

- 🌳 the complete recursive folder tree
- 📄 every discovered file
- 📏 file sizes
- 🧩 file extensions
- 🏷️ MIME types when detectable
- 👀 a preview of the first **5 lines** of readable content
- 📚 text extraction from several document formats
- 🔍 OCR for supported images
- 📑 OCR fallback for scanned PDFs
- ⚠️ notes about unsupported or unreadable files
- 🔗 detection of symbolic links without following them

It supports text files, PDFs, DOC/DOCX, spreadsheets, presentations, OpenDocument formats, images, archives, and various other file types.

For example, it can turn:

```text
Desktop/
├── coisas/
│   ├── trabalho_final.pdf
│   ├── planilha.xlsx
│   └── imagem.png
├── Downloads/
│   └── definitivamente_importante.docx
└── ?????/
    └── arquivo_que_eu_nao_lembro_o_que_e.txt
```

into a much more useful inventory containing metadata **and tiny previews of what those files actually contain**.

The generated report is deliberately limited to the first few lines of each file so your computer doesn't spend three business days attempting to philosophically understand a 400 MB text dump.

The script also skips files larger than **50 MB** and never follows symbolic links. 

---

### `organizar_area_trabalho.py`

The **organizational fairy with a clipboard**.

This script takes a folder — by default, your Windows Desktop — and classifies files based on their **names, paths, extensions, and keywords**, moving them into sensible category folders.

Some of its categories include:

```text
Escola - Secretaria
Escola - Frequência e Presença
Escola - Atas e Resultados
Escola - Censo Escolar
Escola - Alunos e Responsáveis
Escola - Atestados e Declarações
Escola - Inclusão e Atendimento
Escola - Passeios e Autorizações
Escola - Biblioteca e Livros
Escola - Financeiro e Prestação de Contas

Estudos - Faculdade e Matemática
Programação e Projetos

Programas - Instaladores e Sistemas
Imagens e Fotografias
Áudio e Vídeo
Arquivos Compactados
Documentos Gerais
Documentos - PDFs
Documentos - Word e Texto
Planilhas e Dados
Dados e Páginas Web
Temporários e Arquivos Auxiliares
Ferramentas de Organização
```

And, importantly, it **doesn't just start throwing your files around immediately**.

It first creates a proposed organization plan and shows you examples:

```text
arquivo_importante.pdf
    -> Documentos - PDFs/arquivo_importante.pdf

planilha_faltas.xlsx
    -> Escola - Frequência e Presença/planilha_faltas.xlsx
```

Then it asks you to type:

```text
ORGANIZAR
```

If you type literally anything else:

```text
✨ nothing happens ✨
```

The script cancels without moving anything.

---

# 🛡️ Safety First

The organizer was designed with the assumption that **your files are more important than the script's desire to be clever**.

### It does NOT overwrite files

If the destination already contains a file with the same name:

```text
relatorio.pdf
relatorio (2).pdf
relatorio (3).pdf
...
```

are created instead.

No silent replacement. No `oopsies`.

### Protected folders stay protected

Anything containing terms associated with:

```text
lagora
censo
presença
presenca
```

is protected from organization.

The protection applies to paths inside the target folder, including their contents. 

### `shutdown.bat` is sacred

Any file named:

```text
shutdown.bat
```

is never moved.

Its necessary parent folders are preserved as well.

Because sometimes there is a tiny batch file somewhere that you absolutely, definitely, remember the purpose of.

Probably.

Maybe.

Anyway: **it stays.**

### It doesn't delete files

The organizer only removes directories that have become completely empty after organization.

Files themselves are never deleted by the cleanup process.

### It creates an audit trail

After organization, a CSV named:

```text
registro_organizacao.csv
```

records:

- date/time
- original location
- destination
- status
- error details, if applicable

So you can actually inspect what happened instead of staring at your newly organized Desktop wondering what the hell just occurred.

---

# 💻 Installation

## Python

You need **Python 3**.

Check:

```bash
python --version
```

or on some systems:

```bash
python3 --version
```

---

## 📦 Install all Python dependencies at once

The organizer itself uses only Python's standard library, so **it requires no third-party packages**.

The inventory generator has optional format/OCR support.

If you want the **full feature set**, install everything in one command:

```bash
pip install pymupdf python-docx openpyxl xlrd python-pptx odfpy pytesseract Pillow
```

That's the entire Python dependency collection in one ✨ command ✨.

The packages correspond roughly to:

| Package | Used for |
|---|---|
| `pymupdf` | PDF reading + PDF rendering for OCR |
| `python-docx` | `.docx` files |
| `openpyxl` | `.xlsx` / `.xlsm` files |
| `xlrd` | legacy `.xls` files |
| `python-pptx` | `.pptx` presentations |
| `odfpy` | `.odt` / `.odp` files |
| `pytesseract` | OCR |
| `Pillow` | image handling / OCR |

The inventory script itself documents these dependencies and separates the OCR requirements from the other Python packages.

---

# 👁️ OCR Setup

Python packages aren't enough for OCR.

`pytesseract` is basically a Python translator for the actual **Tesseract OCR engine**, so Tesseract itself needs to be installed too.

You also want the Portuguese language data if you're dealing with Brazilian documents.

### Windows

Install Tesseract from:

https://github.com/UB-Mannheim/tesseract/wiki

Make sure the Tesseract executable is available through your `PATH`.

### Linux

```bash
sudo apt install tesseract-ocr tesseract-ocr-por
```

### macOS

```bash
brew install tesseract tesseract-lang
```

The inventory script attempts OCR using:

```text
por+eng
por
eng
```

in that order, so Portuguese documents get the VIP treatment when the language pack is available.

---

# 🚀 Usage

## 1. Make an inventory

Put:

```text
file_inventory_maker.py
```

inside the folder you want to inspect.

Then run:

```bash
python file_inventory_maker.py
```

It will generate:

```text
folder_inventory.txt
```

in the same directory.

The script automatically treats its own directory as the root folder. 

---

## 2. Organize the folder

Run:

```bash
python organizar_area_trabalho.py
```

You can also explicitly provide a folder:

```bash
python organizar_area_trabalho.py "C:/Users/you/Desktop"
```

If you don't provide a path, it attempts to find the Windows Desktop automatically, including common OneDrive locations.

The recommended workflow is:

```text
                    ┌────────────────────┐
                    │   CHAOS DIRECTORY  │
                    │   847 files ???    │
                    └─────────┬──────────┘
                              │
                              ▼
                    file_inventory_maker
                              │
                              ▼
                    folder_inventory.txt
                              │
                    👀 inspect reality
                              │
                              ▼
                  organizar_area_trabalho
                              │
                              ▼
                    safety simulation
                              │
                              ▼
                    type ORGANIZAR
                              │
                              ▼
                    ✨ organized files ✨
```

---

# 🧠 Why use both?

Because blindly organizing a giant folder is a little like performing surgery after looking at the patient for approximately 0.4 seconds.

The inventory tool gives you **situational awareness first**.

The organizer then performs a rule-based cleanup.

That means you can do:

```text
1. Scan
2. Inspect
3. Decide
4. Organize
5. Check the CSV
```

instead of:

```text
1. RUN SCRIPT
2. 😨
3. "WHERE DID MY FILES GO"
```

---

# 🎀 Design Philosophy

This project follows a very specific engineering philosophy:

### ✦ Local

Your files stay on your machine.

### ✦ Deterministic

Classification is based on explicit rules, keywords, paths, and extensions.

No mysterious machine-learning oracle is deciding that `document_final_FINAL2.pdf` is "probably spiritually related to mathematics."

### ✦ Reversible-ish

Files are moved rather than deleted, collisions are handled safely, and an operation log is produced.

### ✦ Paranoid

Protected paths are deliberately excluded.

`shutdown.bat` gets diplomatic immunity.

### ✦ Practical

The scripts are intended for real folders with real accumulated junk, not immaculate example directories containing:

```text
hello.txt
world.py
nice_folder/
```

---

# 🧪 Supported Content

The inventory generator understands a surprisingly large amount of file-world nonsense.

### Text

```text
.txt .md .rst .log .ini .cfg .conf
.toml .yaml .yml .xml .html .css
.js .ts .py .java .c .cpp .cs
.go .rs .php .rb .sh .bat .cmd
.sql .tex .env .gitignore ...
```

### Documents

```text
.pdf
.doc
.docx
.odt
```

### Spreadsheets

```text
.xlsx
.xlsm
.xls
.ods
.csv
```

### Presentations

```text
.pptx
.odp
```

### Images

```text
.png
.jpg
.jpeg
.gif
.bmp
.webp
.tif
.tiff
.ico
.heic
.svg
```

Some image formats can be passed through OCR.

### Archives

```text
.zip
.rar
.7z
.tar
.gz
.bz2
.xz
```

ZIP files have their entries listed, while other archive formats are detected but not extracted.

Audio and video are detected but **not transcribed**. 

---

# 📁 Output Files

After running the tools, you may see:

```text
folder_inventory.txt
```

A human-readable recursive inventory.

and:

```text
registro_organizacao.csv
```

A machine-readable record of the organizer's movements.

These are intentionally boring names.

Boring filenames are good filenames.

---

# 💅 TL;DR

Install the full inventory stack:

```bash
pip install pymupdf python-docx openpyxl xlrd python-pptx odfpy pytesseract Pillow
```

Then:

```bash
python file_inventory_maker.py
```

to **understand the chaos**.

And:

```bash
python organizar_area_trabalho.py
```

to **make the chaos develop a filing system**.

---

## ✦ Final status

```text
Files:        everywhere
Organization: questionable
Python:       available
Fear:         unnecessary
Backup:       still recommended
Aesthetic:    immaculate
```

**May your folders be recursive, your paths relative, and your `shutdown.bat` forever undisturbed.** 🌷🐍💾