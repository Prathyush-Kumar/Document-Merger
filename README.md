# 📄 PDF Merger

A simple Windows desktop utility to merge multiple **PDF** and **DOCX** files into a single PDF — with a clean GUI powered by Tkinter.

---

## ✨ Features

- 📂 Select multiple `.pdf` and `.docx` files via a file picker dialog
- 🔄 Automatically converts `.docx` files to PDF before merging
- 🧹 Cleans up temporary files after merging
- 💾 Saves the merged PDF to a configurable output directory
- ✅ Success popup confirmation when done

---

## 🖥️ Requirements

| Requirement | Details |
|---|---|
| OS | Windows only |
| Python | 3.8 or higher |
| Microsoft Word | Required for `.docx` → PDF conversion |

> ⚠️ **Microsoft Word must be installed** on your machine. The `docx2pdf` library uses Word in the background to convert `.docx` files.

---

## 🚀 Installation

**1. Clone the repository**
```bash
git clone https://github.com/yourusername/pdf-merger.git
cd pdf-merger
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

---

## ▶️ Usage

**1. Set your save location**

Open `pdf_merger.py` and update the `SAVE_PATH` variable at the top:
```python
SAVE_PATH = r"C:\Users\YourName\Documents"
```

**2. Run the script**
```bash
python pdf_merger.py
```

**3. Follow the prompts**
- A file picker will open — select any combination of `.pdf` and `.docx` files
- Enter a name for your merged output file (without `.pdf`)
- The merged file will be saved to `SAVE_PATH`

---

---

## 🔧 Configuration

| Variable | Location | Description |
|---|---|---|
| `SAVE_PATH` | `pdf_merger.py` line 12 | Directory where the merged PDF is saved |

---

## 📦 Dependencies

| Package | Purpose |
|---|---|
| [`pypdf`](https://pypdf.readthedocs.io/) | PDF merging |
| [`docx2pdf`](https://github.com/AlexeevAl/docx2pdf) | DOCX to PDF conversion via Microsoft Word |

> `tkinter` is included with Python's standard library — no installation needed.

---

## ⚠️ Known Limitations

- Windows only (due to `docx2pdf` using the Word COM interface)
- Microsoft Word must be installed for DOCX support
- Very large files may take longer to convert

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
