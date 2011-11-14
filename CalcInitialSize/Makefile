all: insert deploy
insert:
	fsc ExcelAutomation.fsx QuiQPro.fsx FileReader.fsx ListDomain.fsx ListConditionItem.fsx ListCondition.fsx ListItem.fsx ListEntityItem.fsx ListEntity.fsx ListEntityIndex.fsx ListIndex.fsx CalcInitialSize.fsx --standalone -r:"Microsoft.Office.Interop.Excel.dll"
run:
	CalcInitialSize.exe D:\\data\\QuiQpro_CL_8200
clean:
	rm *.exe
deploy:
	cp *.exe D:\\data\\QuiQpro_CL_8200