# DocuSpark

DocuSpark is a batch document transformation tool. It ingests folders of PDFs, DOCX, PPTX, text, HTML, and RTF documents, extracting, cleaning, and converting them into easy-to-read Markdown files. Extracted images are saved with OCR-generated captions for accessibility, making your document archives universally useful and portable.

## Features

- **Batch Folder Conversion**: Processes entire directories of supported documents to clean Markdown.
- **Supported Types**: PDF, DOCX, PPTX, TXT, HTML, RTF
- **Image Extraction & OCR**: Images are exported as PNG and automatically described with OCR (Tesseract).
- **Preserves Structure**: Mirrors the source folder structure in the output directory for organized results.
- **Pandoc Integration**: Uses Pandoc for conversion of generic document types.
- **Accessible Output**: Images get text descriptions for improved readability and accessibility.

## Requirements

- Python 3.8+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) (must be installed and available on your PATH)
- [Pandoc](https://pandoc.org/) (for TXT/HTML/RTF conversion)
- See `requirements.txt` for Python packages: `pip install -r requirements.txt`

## Usage

1. Place documents you wish to convert inside the `data/` subdirectory.
2. Run the program:

    ```bash
    python docuspark.py
    ```

3. Processed Markdown files (with extracted images and OCR descriptions) will appear in the `clean_md/` folder, which mirrors the original folder structure.

## Example Output Structure

```
clean_md/
├── MyDoc.md
├── images/
│   └── img_1.png
└── Subfolder/
    ├── AnotherDoc.md
    └── images/
        └── img_1.png
```

## Extending DocuSpark
- Add additional file converters by editing `docuspark.py`.
- Improve image captioning by upgrading the OCR or swapping in an AI captioning model.
- Supports recursive folder ingestion by design—just drop your files into `data/`.

## License

MIT 
