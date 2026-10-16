# IndicNLP: English-to-Tamil Neural Machine Translation & Engineering Toolkit

[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![HuggingFace Transformers](https://img.shields.io/badge/%F0%9F%A4%97-Transformers-orange.svg)](https://huggingface.co/)
[![PyTorch](https://img.shields.io/badge/PyTorch-EE4C2C?logo=pytorch&logoColor=white)](https://pytorch.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

An end-to-end Natural Language Processing (NLP) pipeline and machine translation workspace developed by **Bharath Suresh**. This repository provides neural sequence-to-sequence translation from **English to Tamil** using fine-tuned transformer architectures, dataset ingestion pipelines, and specialized engineering document generation tools.

---

## ⚡ Features

- **Neural Machine Translation (Seq2Seq)**: Interactive translation engine powered by `suriya7/English-to-Tamil` transformer model.
- **HuggingFace Dataset Ingestion**: Automated loader for bilingual corpora (`gopi30/english-tamil`) with batch conversion to clean pandas DataFrames and CSV formats.
- **Model Checkpoint Management**: Local caching and weight management for fast offline inference.
- **Electrical Engineering MCQ Generator (`pdf.py`)**: Automated script compiling **215+ Electric Circuit Analysis questions** with answers and explanations into formatted `.docx`/Word examination banks for core engineering placement preparation.

---

## 🏗️ Architecture & Pipeline

```
[ English Input Text ]
          │
          ▼
┌─────────────────────────────────┐
│ AutoTokenizer                   │  (suriya7/English-to-Tamil)
│ Subword Tokenization & Tensors  │
└────────────────┬────────────────┘
                 │
                 ▼
┌─────────────────────────────────┐
│ Seq2Seq Neural Architecture     │  (AutoModelForSeq2SeqLM)
│ Encoder-Decoder Attention       │
└────────────────┬────────────────┘
                 │
                 ▼
┌─────────────────────────────────┐
│ Beam Search Decoding            │  (max_length: 128)
└────────────────┬────────────────┘
                 │
                 ▼
[ Tamil Translated Output ]
```

---

## 📁 Repository Structure

```
├── run_model.py          # Interactive CLI for English-to-Tamil translation
├── download_dataset.py   # HuggingFace dataset downloader (gopi30/english-tamil)
├── CsvCon.py             # Converts HuggingFace dataset to structured CSV
├── hug.py                # Pre-downloads & caches transformer model weights
├── log.py                # Secure Hugging Face Hub authentication
├── pdf.py                # Electric Circuit Analysis 215+ MCQ generator
├── requirements.txt      # Python dependencies
├── .env.example          # Environment variable template
└── .gitignore            # Ignores large datasets and credentials
```

---

## 🚀 Quick Start

### 1. Clone the Repository

```bash
git clone https://github.com/bharath-eee/english-tamil-neural-translator.git
cd english-tamil-neural-translator
```

### 2. Set Up Virtual Environment

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1   # On Windows
# source venv/bin/activate    # On Linux/macOS
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. (Optional) Hugging Face Login

If accessing gated datasets or private models:
```bash
cp .env.example .env
# Add your HUGGINGFACE_TOKEN to .env
python log.py
```

---

## 💻 Usage

### 1. Run Interactive Translator

```bash
python run_model.py
```

**Interactive Menu:**
```
--- Translator ---
1. English -> Tamil
2. Tamil -> English
3. Exit
Choose option: 1
Enter English text: What is your name?
Tamil: உங்கள் பெயர் என்ன?
```

### 2. Download & Export Bilingual Dataset

```bash
python download_dataset.py
python CsvCon.py
```
This loads the `gopi30/english-tamil` dataset and exports `english_tamil_dataset.csv`.

### 3. Generate Circuit Analysis Question Bank

```bash
python pdf.py
```
Compiles 215+ comprehensive multiple-choice questions on Units, Components, Kirchhoff's Laws, Theorems, AC Circuits, and Transients into an organized Word/DOCX document.

---

## 🛠️ Tech Stack

- **Model**: `suriya7/English-to-Tamil` (Seq2Seq Transformer)
- **Dataset**: `gopi30/english-tamil` (Hugging Face Datasets)
- **Frameworks**: PyTorch, HuggingFace Transformers, HuggingFace Hub, Pandas, Python-Docx

---

## 👨‍💻 Author

**Bharath Suresh**  
Electrical & Electronics Engineering Student  
- GitHub: [@bharath-eee](https://github.com/bharath-eee)  
- LinkedIn: [Bharath Suresh](https://www.linkedin.com/in/bharath-s-b22842329/)

---

## 📄 License

This repository is licensed under the [MIT License](LICENSE).
