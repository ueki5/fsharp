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
        if dicCondition.ContainsKey(conditem.Value.ConditionPhysicalName)
        then
            let cond = dicCondition.[conditem.Value.ConditionPhysicalName]
            let _ = cond.ConditionItems.Add(conditem.Key, conditem.Value)
            ()
        else ()
let ProcessItem =
    for item in dicItem do
        if dicCondition.ContainsKey(item.Value.ConditionPhysicalName)
        then
            let cond = dicCondition.[item.Value.ConditionPhysicalName]
            ignore(item.Value.ConditionRef <- Some(cond))
            ()
        else ()
let ProcessEntityItem =
    for entitem in dicEntityItem do
        if dicEntityIndex.ContainsKey(entitem.Key)
        then
            let indexitem = dicEntityIndex.[entitem.Key]
            ignore(entitem.Value.PkeyIndex <- Some(indexitem.ItemIndex))
            ()
        else ()
        if dicItem.ContainsKey(entitem.Value.PhysicalName)
        then
            let item = dicItem.[entitem.Value.PhysicalName]
            ignore(entitem.Value.ItemRef <- Some(item))
            match item.ConditionRef with
            | None -> ()
            | Some cond -> ignore(entitem.Value.ConditionRef <- Some(cond))
            ()
        else ()
        match entitem.Value.ConditionRef with
        | None ->
            if dicCondition.ContainsKey(entitem.Value.ConditionPhysicalName)
            then
                let cond = dicCondition.[entitem.Value.ConditionPhysicalName]
                ignore(entitem.Value.ConditionRef <- Some(cond))
                ()
            else ()
        | Some _ -> ()
        if dicEntity.ContainsKey(entitem.Value.EntityPhysicalName)
        then
            let ent = dicEntity.[entitem.Value.EntityPhysicalName]
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

    worksheet.Name <- ent.PhysicalName
    for entitem in ent.EntityItems.Values do
        match IsCommonItem(entitem.PhysicalName) with
        | true     -> ()
        | false    -> 
            let ColorIndexBk = worksheet.Range("CELL_ITEM_VALUE").Interior.ColorIndex
            worksheet.Range("CELL_ITEM_LOGICAL_NAME").Value2 <- entitem.LogicalName
            worksheet.Range("CELL_ITEM_PHYSICAL_NAME").Value2 <- entitem.PhysicalName
            worksheet.Range("CELL_ITEM_REMARKS").Value2 <- entitem.Remarks
            match entitem.ItemRef with
            | None -> ()
            | Some itemref ->
                worksheet.Range("CELL_ITEM_DATA_TYPE").Value2 <- itemref.DataType
                worksheet.Range("CELL_ITEM_DATA_LENGTH").Value2 <- itemref.DataLengthDsp
                worksheet.Range("CELL_ITEM_VALUE").NumberFormatLocal <- itemref.NumberFormat
                // printfn "type=%s,len1=%d,len2=%d,format=%s" itemref.DataType itemref.DataLength1 itemref.DataLength2 itemref.NumberFormat
                worksheet.Range("COL_COLUMN").ColumnWidth <- itemref.ColumnWidth
                // worksheet.Range("COL_COLUMN").EntireColumn.AutoFit()
                // if itemref.ColumnWidth > worksheet.Range("COL_COLUMN").ColumnWidth
                // then worksheet.Range("COL_COLUMN").ColumnWidth <- itemref.ColumnWidth
                // else ()

                let validation = worksheet.Range("CELL_ITEM_VALUE").Validation
                ignore(validation.Delete())
                ignore <| match (itemref.DataType, itemref.ConditionRef) with
                            | (_ , Some(cond)) ->
                                validation.Add(
                                    XlDVType.xlValidateList
                                    , XlDVAlertStyle.xlValidAlertStop
                                    , XlFormatConditionOperator.xlBetween
                                    , (GetDropDownList cond.ConditionItems))
                                validation.IMEMode <- int XlIMEMode.xlIMEModeOff
                            | ("NUMBER", None) ->
                                validation.Add(
                                    XlDVType.xlValidateDecimal
                                    , XlDVAlertStyle.xlValidAlertStop
                                    , XlFormatConditionOperator.xlBetween
                                    , itemref.NumberFormatMin
                                    , itemref.NumberFormatMax)
                                validation.IMEMode <- int XlIMEMode.xlIMEModeOff
                            | ("CHAR", None) -> 
                                validation.Add(
                                    XlDVType.xlValidateCustom
                                    , XlDVAlertStyle.xlValidAlertStop
                                    , XlFormatConditionOperator.xlBetween
                                    , "=LENB(A9)" + itemref.ValidationCusmomOperator + "A$7")
                                validation.IMEMode <- int XlIMEMode.xlIMEModeOff
                            | ("VARCHAR2", None) ->
                                validation.Add(
                                    XlDVType.xlValidateCustom
                                    , XlDVAlertStyle.xlValidAlertStop
                                    , XlFormatConditionOperator.xlBetween
                                    , "=LENB(A9)" + itemref.ValidationCusmomOperator + "A$7")
                                validation.IMEMode <- int XlIMEMode.xlIMEModeOn
                            | _ -> ()
                validation.IgnoreBlank <- true
                validation.InCellDropdown <- true
                validation.InputTitle <- ""
                validation.ErrorTitle <- ""
                validation.InputMessage <- ""
                validation.ErrorMessage <- ""
                validation.ShowInput <- true
                validation.ShowError <- true
            match entitem.PkeyIndex with
            | None -> ()
            | Some pkeyindex -> 
                worksheet.Range("CELL_ITEM_VALUE").Interior.ColorIndex <- 38
            ignore <| worksheet.Range("COL_COLUMN").Copy()
            ignore <| worksheet.Range("COL_INSERTAT").Insert(XlDirection.xlToRight)
            ignore <| worksheet.Range("CELL_ITEM_VALUE").Interior.ColorIndex <- ColorIndexBk
    ignore <| worksheet.Range("COL_COLUMN").Delete()
    ignore <| worksheet.Range("COL_INSERTAT").Delete()
    // printfn "%s" (GetInsertSql1 ent)
    // printfn "%s" (GetInsertSql2 ent)
    worksheet.Range(Cell (1,1)).Value2 <- ent.PhysicalName
    worksheet.Range(Cell (1,2)).Value2 <- ent.LogicalName
    worksheet.Range(Cell (1,2)).Value2 <- ent.LogicalName
    worksheet.Range(Cell (2,1)).Value2 <- "日時"
    worksheet.Range(Cell (3,1)).Value2 <- "ユーザＩＤ"
    worksheet.Range(Cell (4,1)).Value2 <- "クライアントＩＰ"
    worksheet.Range(Cell (5,1)).Value2 <- "アプリサーバＩＰ"
    worksheet.Range(Cell (6,1)).Value2 <- "プログラムＩＤ"
    worksheet.Range(Cell (2,2)).Value2 <- "ZPK_DBINFO.GET_DATETIME"
    worksheet.Range(Cell (3,2)).Value2 <- "SETUP"
    worksheet.Range(Cell (4,2)).Value2 <- "ZPK_DBINFO.GET_IPADDR"
    worksheet.Range(Cell (5,2)).Value2 <- "ZPK_DBINFO.GET_IPADDR"
    worksheet.Range(Cell (6,2)).Value2 <- "SQL"
    worksheet.Range(Cell (GetSqlPos ent 0)).Value2 <- (GetInsertSql1 ent)
    worksheet.Range(Cell (GetSqlPos ent 1)).Value2 <- (GetInsertSql2 ent)
    for entitem in ent.EntityItems.Values do
        // printfn "%s" ("=" + (GetSqlValue entitem) + " & \",\" & " + Cell (GetSqlPos ent (2 + entitem.ItemIndex)))
        match entitem.ItemIndex < ent.EntityItems.Count with
        | true -> worksheet.Range(Cell (GetSqlPos ent (1 + entitem.ItemIndex))).Value2 <- "=" + (GetSqlValue entitem) + " & \",\" & " + Cell (GetSqlPos ent (2 + entitem.ItemIndex))
        | false -> worksheet.Range(Cell (GetSqlPos ent (1 + entitem.ItemIndex))).Value2 <- "=" + (GetSqlValue entitem)
    workbook.SaveAs(outDir + "\\" + ent.PhysicalName + "("+ ent.LogicalName + ")" + ".xls")
    app.UserControl <- false
    app.Quit()
let MakeDirectory dirpath =
    if Directory.Exists(dirpath)
    then Directory.Delete(dirpath,true)
    else ()
    Directory.CreateDirectory(dirpath)
#if COMPILED
[<EntryPoint>]
#endif
let main (_) =
    ignore <| MakeDirectory(outDir)
    ignore <| ProcessConditionItem
    ignore <| ProcessEntityItem
    for ent in dicEntity.Values do
        match (IsTarget ent) with
        | true ->
            printfn "エンティティ[%s]を処理中…" ent.PhysicalName
            CreateExcel ent
        | false ->
            printfn "エンティティ[%s]をスキップしました。" ent.PhysicalName
    0
#if INTERACTIVE
main
#endif
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
