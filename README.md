## Description

Unprotect your xlsx file (Excel) with this simple python script.

## Usage

`excel-unprotector input.xlsm output.xlsm`

## Caution

It only works if your file is NOT protected for reading and ONLY for writing / editing. If your file is protected for reading, the content of the file is encrypted.

If you have some macros, you will get a warning about corrupted file after unprotect. Just save and reload your file.

If you have an xls file, save it to xlsx / xlsm file format.

## Macros that protect the file again

A macro can protect the workbook or a sheet again each time it runs (`ActiveWorkbook.Protect`, `ActiveSheet.Protect`, `Sheets(...).Protect`...). After the conversion, the script scans the VBA macros and prints a warning listing every line that calls `Protect`. Remove these lines in the VBA editor. The password is often written in them (`Password:="..."`).
