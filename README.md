# ScribeEase

Turn photos and scans of handwritten or printed pages into an editable Word document.

## Features

* Upload images (JPG, PNG, HEIC, WebP) or multi-page PDFs; combine several files into one document
* Handwriting is reflowed into natural paragraphs instead of copying the page's line breaks
* Detects tables and exports them as real Word tables
* Supports maths and chemistry notation: superscripts, subscripts, arrows and Greek letters
* Review and correct the text in the browser before downloading the `.docx`
* Files are processed in memory and are never stored on the server

## How it works

Each page is sent to the Google Gemini API for transcription. You proofread the result in an editor,
then export it as a formatted `.docx` file.

## Tech stack

Flask, Google Gemini API, python-docx, PyMuPDF, Pillow

## Getting started

Requires Python 3.10+ and a [Gemini API key](https://aistudio.google.com/apikey).

```powershell
git clone https://github.com/jemmadev/ScribeEase.git
cd ScribeEase
copy .env.example .env      # then add your GEMINI\\\_API\\\_KEY to .env
.\\\\start.ps1
```

Open http://127.0.0.1:5000 in your browser.

## Note

Pages are sent to Google's Gemini API for processing. Review Google's current terms before uploading
sensitive documents.

