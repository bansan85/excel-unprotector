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

## Name errors after an xls to xlsx conversion

An xls file has at most 256 columns (last one `IV`), an xlsx file has 16384 (last one `XFD`). A defined name such as `TOT2010` was not a cell address in xls, but it is one in xlsx.

When saving to xlsx, Excel renames these conflicting names with a leading underscore (`TOT2010` becomes `_TOT2010`) and updates the formulas, but NOT the VBA macros. A macro using `Range("TOT2010")` then silently targets the cell TOT2010 instead of the named range.

Fix the macros by replacing each renamed name with its new name (`Range("_TOT2010")`). The list of renamed names is in the Name Manager (Formulas > Name Manager).
