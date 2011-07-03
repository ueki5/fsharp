
all:rep1 rep2 rep3
rep1:PAReport1.exe
	fsc ExcelAutomation.fsx PAReport1.fsx --standalone -r:"Microsoft.Office.Interop.Excel.dll"
rep2:PAReport2.exe
	fsc ExcelAutomation.fsx PAReport2.fsx --standalone -r:"Microsoft.Office.Interop.Excel.dll"
rep3:PAReport3.exe
	fsc ExcelAutomation.fsx PAReport3.fsx --standalone -r:"Microsoft.Office.Interop.Excel.dll"
run:run1 run2 run3
run1:rep1
	PAReport1.bat
run2:rep2
	PAReport2.bat
run3:rep3
	PAReport3.bat
