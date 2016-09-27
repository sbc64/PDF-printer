#/bin/bash
rm -r __pycache__
rm -r build
rm -r dist
pyinstaller --noconsole --paths='C:\\Python35\\Lib\\site-packages' pdfprinter.spec pdfprinter.py
echo "Moving file to T:\\.."
#mv dist/pdfprinter.exe "T:\RELEASED_FILES\PDFprinterV0.1.1.exe"
echo "Done"
