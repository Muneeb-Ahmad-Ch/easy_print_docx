# Easy Print DOCX Files Without MS Word, LibreOffice, or Any Office Suite

## Overview

This Python script allows you to **Print DOCX Files Without MS Word, LibreOffice, or Any Office Suite**. It achieves this by:

1. Converting DOCX to PDF using [`docto.exe`](https://github.com/tobya/DocTo)
2. Printing the PDF directly using [`SumatraPDF.exe`](https://www.sumatrapdfreader.org/)

This approach ensures an **efficient and lightweight** method for printing DOCX files on Windows.

## Features

- ✅ No need for MS Word or LibreOffice
- ✅ Uses `docto.exe` for DOCX-to-PDF conversion
- ✅ Uses `SumatraPDF.exe` for direct printing
- ✅ Lightweight and fast execution
- ✅ Fully automated process

## Requirements

- **Windows OS** (since `docto.exe` and `SumatraPDF.exe` are Windows-based)
- Python 3.x installed
- `docto.exe` and `SumatraPDF.exe` available in the `dependencies` folder

## Usage

```python
from easy_print_docx import print_docx

print_docx("path/to/your/document.docx")
```

### How It Works

1. The script **converts** the DOCX file to a temporary PDF.
2. It **sends** the PDF to the default printer.
3. The **temporary file is deleted** after printing.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Pull requests are welcome! Feel free to open an issue if you encounter any problems.

## Credits

- [`docto.exe`](https://github.com/tobya/DocTo) (DOCX to PDF conversion)
- [`SumatraPDF.exe`](https://www.sumatrapdfreader.org/) (PDF printing)
