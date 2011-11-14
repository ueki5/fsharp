all: update insert deploy
insert:
	fsc ExcelAutomation.fsx QuiQPro.fsx FileReader.fsx ListDomain.fsx ListConditionItem.fsx ListCondition.fsx ListItem.fsx ListEntityItem.fsx ListEntity.fsx ListEntityIndex.fsx EntrySheetGen.fsx --standalone -r:"Microsoft.Office.Interop.Excel.dll"
run:
	EntrySheetGen.exe D:\\data\\QuiQpro_CL_8200
update:
	fsc ExcelAutomation.fsx QuiQPro.fsx FileReader.fsx ListDomain.fsx ListConditionItem.fsx ListCondition.fsx ListItem.fsx ListEntityItem.fsx ListEntity.fsx ListEntityIndex.fsx UpdateSheetGen.fsx --standalone -r:"Microsoft.Office.Interop.Excel.dll"
runupdate:
	UpdateSheetGen.exe D:\\data\\QuiQpro_CL_8200
clean:
	rm *.exe
deploy:
	cp *.exe D:\\data\\QuiQpro_CL_8200