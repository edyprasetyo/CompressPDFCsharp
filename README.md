# Compress PDF

This script compresses a PDF file using Ghostscript.

## Dependencies

- Python 3.x
- Ghostscript

## Usage

1. Install Ghostscript on your machine.
2. Clone this repository.
3. Open a terminal or command prompt in the directory where the script is located.
4. Run the script with the following command:

   ```
   python compress_pdf.py input_file_path output_file_path quality
   ```

   Replace `input_file_path` with the path to the PDF file you want to compress, `output_file_path` with the desired path for the compressed PDF file, and `quality` with one of the following values:

   - `screen`: low-resolution output (72 dpi)
   - `ebook`: medium-resolution output (150 dpi)
   - `printer`: high-resolution output (300 dpi)

   Example usage:

   ```
   python compress_pdf.py D:\MyProject\CompressPDF\input.pdf D:\MyProject\CompressPDF\output.pdf ebook
   ```

5. The compressed PDF file will be saved to the specified output path. The script will also print the size of the input and output files in KB and open the output file.

## Script Code

```python
import subprocess
import os


def compress_pdf(input_path, output_path, quality='ebook'):
    """
    :param quality: The quality of the output PDF file. The value can be one of the following:
        - 'screen': low-resolution output (72 dpi)
        - 'ebook': medium-resolution output (150 dpi)
        - 'printer': high-resolution output (300 dpi)
    """
    gs_command = [
        r'C:\Program Files\gs\gs10.00.0\bin\gswin64.exe',
        '-sDEVICE=pdfwrite', f'-dPDFSETTINGS=/{quality}',
        '-dCompatibilityLevel=1.4',
        '-dNOPAUSE',
        '-dQUIET',
        '-dBATCH',
        '-dDetectDuplicateImages=true',
        '-dCompressFonts=true',
        '-dDownsampleColorImages=true',
        '-dColorImageResolution=120',
        '-dMonoImageResolution=120',
        '-sOutputFile=' + os.path.abspath(output_path),
        os.path.abspath(input_path)
    ]
    subprocess.run(gs_command)


input_file_path = input('Enter path to input PDF file: ')
output_file_path = input('Enter path to output PDF file: ')
quality = input('Enter compression quality (screen, ebook, or printer): ')

compress_pdf(input_file_path, output_file_path, quality)

# print output.pdf size in KB
print(str(int(os.path.getsize(output_file_path) / 1024)) + ' KB')

# open output.pdf
subprocess.Popen(output_file_path, shell=True)
```
