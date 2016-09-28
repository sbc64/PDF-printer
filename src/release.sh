#/bin/bash
rm -r __pycache__
rm -r build
rm -r dist
#echo "Copying binarie sinto src dir"
#cp "C:\Program Files\gs\gs9.19\bin\gsdll64.dll" .
#cp "C:\Program Files\gs\gs9.19\bin\gsdll64.lib" .
#cp "C:\Program Files\gs\gs9.19\bin\gswin64.exe" .
#cp "C:\Program Files\gs\gs9.19\bin\gswin64c.exe" .
pyinstaller --noconsole --onefile pdfprinter.spec pdfprinter.py
echo "Moving file to T:\\.."
mv dist/pdfprinter.exe "T:\RELEASED_FILES\PDFprinterV0.1.3.exe"

#rm gswin64.exe
#rm gswin64c.exe
#rm gsdll64.lib
#rm gsdll64.dll

#echo "removing binaries from src dir..."

echo "Done"
