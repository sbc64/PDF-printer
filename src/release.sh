#/bin/bash
rm -r __pycache__
rm -r build
rm -r dist
#echo "Copying binarie sinto src dir"
#cp "C:\Program Files\gs\gs9.19\bin\gsdll64.dll" .
#cp "C:\Program Files\gs\gs9.19\bin\gsdll64.lib" .
#cp "C:\Program Files\gs\gs9.19\bin\gswin64.exe" .
#cp "C:\Program Files\gs\gs9.19\bin\gswin64c.exe" .

gitStatusPorcelain=$(git status --porcelain)

if [[ -z $gitStatusPorcelain ]]
	then
	pyinstaller --noconsole --onefile pdfprinter.spec pdfprinter.py
	echo "Moving file to T:\\.."
	commit=$(git log --pretty=format:%H -n1| cut -c1-6)
	path=$(echo "T:\RELEASED_FILES\PDFprinterV"$commit".exe")
	echo $path
	cp dist/pdfprinter.exe $path
	echo "Done"
else
	echo "Requires commit"
fi
#rm gswin64.exe
#rm gswin64c.exe
#rm gsdll64.lib
#rm gsdll64.dll

#echo "removing binaries from src dir..."
