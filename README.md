# DOCX Document Translator
A multilingual DOCX document translation tool built with Streamlit, supporting mutual translation between Chinese, English, French, German, Spanish, Arabic, Japanese, Korean and other languages.

## Core Features
- 📤 Upload .docx format documents and automatically extract text from paragraphs/tables
- ⚡ Multi-threaded batch translation (1-8 threads adjustable) for faster processing of large documents
- 📝 Fully preserve the original document format and download the translated .docx file directly
- 🌐 Support mutual translation between 8 commonly used languages, powered by Google Translator for accurate results

## Usage Guide (with ILovePDF)
Since this tool only supports .docx format, it is recommended to use it with the ILovePDF website (https://www.ilovepdf.com/):
1. For PDF files to be translated: First convert PDF to DOC/DOCX format via ILovePDF
2. Upload the converted DOCX file to this tool to complete translation
3. For PDF format results (if needed): Convert the translated DOCX file back to PDF via ILovePDF

## Technical Notes (Why Format Conversion is Not Integrated)
PDF ↔ DOCX conversion functionality is not directly integrated into Streamlit for the following key reasons:
1. Performance testing: Embedding format conversion code in the app resulted in extremely slow overall processing speed (especially for large files)
2. API integration attempt: I previously tried to call ILovePDF's API for automatic conversion, but the API consistently returned a 500 error (internal server error), making it impossible to use stably

Therefore, a "tool splitting" solution was adopted: use professional ILovePDF for format conversion and this tool for translation, balancing efficiency and stability.

## Deployment & Running
### Install Dependencies
```bash
pip install -r requirements.txt
