#if INTERACTIVE
#r "Microsoft.Office.Interop.Excel.dll"
#load "ExcelAutomation.fsx"
#load "FileReader.fsx"
#load "ListDomain.fsx"
#load "ListCondition.fsx"
// #load "ListConditionItem.fsx"
#load "ListItem.fsx"
#load "ListEntity.fsx"
// #load "ListEntityItem.fsx"
#load "ListEntityIndex.fsx"
#endif
#if COMPILED
module EntrySheetGen
#endif
open System
open System.IO
open System.Collections.Generic
open System.Runtime.InteropServices
open Microsoft.Office.Interop.Excel
open ExcelAutomation
open FileReader
open ListDomain
open ListCondition
open ListConditionItem
open ListItem
open ListEntity
open ListEntityItem
open ListEntityIndex

let currDir = Directory.GetCurrentDirectory()
let csvDir = currDir + "\\" + "CSV"
let tmplDir = currDir + "\\" + "TEMPLATE"
let outDir = currDir + "\\" + "OUT"
let dicDomain =
    FileToArray(csvDir + "\\" + "List_Domain.CSV")
    |> MakeListDomain
let dicCondition =
    FileToArray(csvDir + "\\" + "List_Condition.CSV")
    |> MakeListCondition
let dicConditionItem =
    FileToArray(csvDir + "\\" + "List_Condition_Item.CSV")
    |> MakeListConditionItem
let dicItem =
    FileToArray(csvDir + "\\" + "List_Item.CSV")
    |> MakeListItem
let dicEntity =
    FileToArray(csvDir + "\\" + "List_Entity.CSV")
    |> MakeListEntity
let dicEntityItem =
    FileToArray(csvDir + "\\" + "List_Entity_Item.CSV")
    |> MakeListEntityItem
let dicEntityIndex =
    FileToArray(csvDir + "\\" + "List_Entity_Index.CSV")
    |> MakeListEntityIndex
let ProcessConditionItem =
    for conditem in dicConditionItem do
        if dicCondition.ContainsKey(conditem.Value.ConditionId)
        then
            let cond = dicCondition.[conditem.Value.ConditionId]
            let _ = cond.ConditionItems.Add(conditem.Key, conditem.Value)
            ()
        else ()
let ProcessEntityItem =
    for entitem in dicEntityItem do
        if dicEntity.ContainsKey(entitem.Value.EntPhysicalName)
        then
            let ent = dicEntity.[entitem.Value.EntPhysicalName]
            let _ = ent.EntityItems.Add(entitem.Key, entitem.Value)
            ()
        else ()
let CreateExcel (ent:Entity) = 
    let app = new ApplicationClass(Visible = false) 
    app.DisplayAlerts <- false
    let workbooks = app.Workbooks
    let workbook = workbooks.Open(tmplDir + "\\" + "Document.XLT")
    let sheets = workbook.Worksheets
    let worksheet = (sheets.[box 1] :?> _Worksheet)


    for item in ent.EntityItems.Values do
        match item.PhysicalName with
        | "X_INS_DATETIME"
        | "X_INS_USER_ID"
        | "X_INS_CLIENT_IP"
        | "X_INS_APSERVER_IP"
        | "X_INS_PG_ID"
        | "D_UPD_DATETIME"
        | "D_UPD_USER_ID"
        | "D_UPD_CLIENT_IP"
        | "D_UPD_APSERVER_IP"
        | "D_UPD_PG_ID"
        | "B_UPD_DATETIME"
        | "B_UPD_USER_ID"
        | "B_UPD_CLIENT_IP"
        | "B_UPD_APSERVER_IP"
        | "B_UPD_PG_ID" -> ()
        | _ -> 
            worksheet.Range("CELL_ITEM_LOGICAL_NAME").Value2 <- item.LogicalName
            worksheet.Range("CELL_ITEM_PHYSICAL_NAME").Value2 <- item.PhysicalName
            worksheet.Range("CELL_ITEM_REMARKS").Value2 <- item.Remarks
            ignore <| worksheet.Range("COL_COLUMN").Copy()
            ignore <| worksheet.Range("COL_INSERTAT").Insert(XlDirection.xlToRight)
    ignore <| worksheet.Range("COL_COLUMN").Delete()
    ignore <| worksheet.Range("COL_INSERTAT").Delete()
    ignore <| worksheet.Range(Cell (1,1)).Value2 <- ent.PhysicalName
    ignore <| worksheet.Range(Cell (1,2)).Value2 <- ent.LogicalName
    workbook.SaveAs(outDir + "\\" + ent.PhysicalName + "("+ ent.LogicalName + ")" + ".xls")
    app.UserControl <- false
    app.Quit()
let MakeDirectory dirpath =
    if Directory.Exists(dirpath)
    then Directory.Delete(dirpath,true)
    else ()
    Directory.CreateDirectory(dirpath)
[<EntryPoint>]
let main (_) =
    ignore <| MakeDirectory(outDir)
    ignore <| ProcessConditionItem
    ignore <| ProcessEntityItem
    for ent in dicEntity.Values do
        CreateExcel ent
    0
// #if COMPILED
// System.Threading.Thread.Sleep(1000)
// //  If user interacted with Excel it will not close when the app object is destroyed, so we close it explicitly 
// let _ = 
//   try 
//     workbook.Saved <- true
//     app.UserControl <- false
//     app.Quit()
//   with e -> Console.WriteLine ("User closed Excel manually, so we don't have to do that")

// let _ = Console.WriteLine ("Sample successfully finished!")
// #endif
//fsc ExcelAutomation.fsx QuiQPro.fsx FileReader.fsx ListDomain.fsx ListConditionItem.fsx ListCondition.fsx ListItem.fsx ListEntityItem.fsx ListEntity.fsx ListEntityIndex.fsx EntrySheetGen.fsx --standalone -r:"Microsoft.Office.Interop.Excel.dll"
