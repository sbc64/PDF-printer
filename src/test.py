import win32print

print (win32print.EnumPrinters(2))
printer = win32print.OpenPrinter('KONICA MINOLTA 423')
print (win32print.GetPrinter(printer))
