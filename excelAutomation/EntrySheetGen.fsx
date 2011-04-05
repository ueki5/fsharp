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

let GetDomain csvDir =
    FileToArray(csvDir + "\\" + "List_Domain.CSV")
    |> MakeListDomain
let GetCondition csvDir =
    FileToArray(csvDir + "\\" + "List_Condition.CSV")
    |> MakeListCondition
let GetConditionItem csvDir =
    FileToArray(csvDir + "\\" + "List_Condition_Item.CSV")
    |> MakeListConditionItem
let GetItem csvDir =
    FileToArray(csvDir + "\\" + "List_Item.CSV")
    |> MakeListItem
let GetEntity csvDir =
    FileToArray(csvDir + "\\" + "List_Entity.CSV")
    |> MakeListEntity
let GetEntityItem csvDir =
    FileToArray(csvDir + "\\" + "List_Entity_Item.CSV")
    |> MakeListEntityItem
let GetEntityIndex csvDir =
    FileToArray(csvDir + "\\" + "List_Entity_Index.CSV")
    |> MakeListEntityIndex
let ProcessConditionItem (dicConditionItem:ConditionItemDictionary) (dicCondition:ConditionDictionary) =
    for conditem in dicConditionItem do
        if dicCondition.ContainsKey(conditem.Value.ConditionPhysicalName)
        then
            let cond = dicCondition.[conditem.Value.ConditionPhysicalName]
            let _ = cond.ConditionItems.Add(conditem.Key, conditem.Value)
            ()
        else ()
let ProcessItem (dicItem:ItemDictionary) (dicCondition:ConditionDictionary) =
    for item in dicItem do
        if dicCondition.ContainsKey(item.Value.ConditionPhysicalName)
        then
            let cond = dicCondition.[item.Value.ConditionPhysicalName]
            ignore(item.Value.ConditionRef <- Some(cond))
            ()
        else ()
let ProcessEntityItem
    (dicEntityItem:EntityItemDictionary)
    (dicEntityIndex:EntityIndexDictionary)
    (dicItem:ItemDictionary)
    (dicCondition:ConditionDictionary)
    (dicEntity:EntityDictionary) =
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
let ReNumberEntityItem (dicEntity:EntityDictionary) =
    // 一般の項目が前に、共通項目が後になるよう、番号を振り直す。
    for ent in dicEntity.Values do
        let mutable idx = 0
        for entitem in ent.EntityItems.Values do
            match IsCommonItem entitem.PhysicalName with
            | false ->
                idx <- idx + 1
                entitem.ItemIndex <- idx
            | true -> ()
        for entitem in ent.EntityItems.Values do
            match IsCommonItem entitem.PhysicalName with
            | false -> ()
            | true ->
                idx <- idx + 1
                entitem.ItemIndex <- idx
    ()

let SetPattern (range:Range) =
    range.Borders.[XlBordersIndex.xlDiagonalDown].LineStyle <- Constants.xlNone
    range.Borders.[XlBordersIndex.xlDiagonalDown].LineStyle <- Constants.xlNone
    range.Borders.[XlBordersIndex.xlDiagonalUp].LineStyle <- Constants.xlNone
    range.Borders.[XlBordersIndex.xlEdgeLeft].LineStyle <- XlLineStyle.xlContinuous
    range.Borders.[XlBordersIndex.xlEdgeLeft].Weight <- XlBorderWeight.xlThin
    range.Borders.[XlBordersIndex.xlEdgeLeft].ColorIndex <- Constants.xlAutomatic
    range.Borders.[XlBordersIndex.xlEdgeTop].LineStyle <- XlLineStyle.xlContinuous
    range.Borders.[XlBordersIndex.xlEdgeTop].Weight <- XlBorderWeight.xlThin
    range.Borders.[XlBordersIndex.xlEdgeTop].ColorIndex <- Constants.xlAutomatic
    range.Borders.[XlBordersIndex.xlEdgeBottom].LineStyle <- XlLineStyle.xlContinuous
    range.Borders.[XlBordersIndex.xlEdgeBottom].Weight <- XlBorderWeight.xlThin
    range.Borders.[XlBordersIndex.xlEdgeBottom].ColorIndex <- Constants.xlAutomatic
    range.Borders.[XlBordersIndex.xlEdgeRight].LineStyle <- XlLineStyle.xlContinuous
    range.Borders.[XlBordersIndex.xlEdgeRight].Weight <- XlBorderWeight.xlThin
    range.Borders.[XlBordersIndex.xlEdgeRight].ColorIndex <- Constants.xlAutomatic

let CreateExcel (ent:Entity) (outDir:string) = 
    let app = new ApplicationClass(Visible = false) 
    app.DisplayAlerts <- false
    let workbooks = app.Workbooks
    let workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet) 
    let sheets = workbook.Worksheets 
    let worksheet = (sheets.[box 1] :?> _Worksheet) 
    // let app = new ApplicationClass(Visible = false) 
    // app.DisplayAlerts <- false
    // let workbooks = app.Workbooks
    // // let workbook = workbooks.Open(tmplDir + "\\" + "Document.XLT")
    // let sheets = workbook.Worksheets
    // let worksheet = (sheets.[box 1] :?> _Worksheet)

    worksheet.Name <- ent.PhysicalName
    for entitem in ent.EntityItems.Values do
        match IsCommonItem(entitem.PhysicalName) with
        | true     -> ()
        | false    ->
            for idx in [4 .. InputRow] do
                SetPattern (worksheet.Range (Cell (entitem.ItemIndex, idx)))
            worksheet.Range(Cell (entitem.ItemIndex, 4)).Value2 <- entitem.LogicalName
            worksheet.Range(Cell (entitem.ItemIndex, 4)).Interior.ColorIndex <- TitleBackColor
            worksheet.Range(Cell (entitem.ItemIndex, 4)).Font.ColorIndex <- TitleFontColor
            worksheet.Range(Cell (entitem.ItemIndex, 4)).HorizontalAlignment <- Constants.xlLeft
            worksheet.Range(Cell (entitem.ItemIndex, 5)).Value2 <- entitem.PhysicalName
            worksheet.Range(Cell (entitem.ItemIndex, 5)).Interior.ColorIndex <- TitleBackColor
            worksheet.Range(Cell (entitem.ItemIndex, 5)).Font.ColorIndex <- TitleFontColor
            worksheet.Range(Cell (entitem.ItemIndex, 5)).HorizontalAlignment <- Constants.xlLeft
            worksheet.Range(Cell (entitem.ItemIndex, 8)).Value2 <- entitem.Remarks
            worksheet.Range(Cell (entitem.ItemIndex, 8)).Interior.ColorIndex <- RemarksColor
            worksheet.Range(Cell (entitem.ItemIndex, 8)).WrapText <- true
            worksheet.Range(Cell (entitem.ItemIndex, 8)).HorizontalAlignment <- Constants.xlLeft

            match entitem.ItemRef with
            | None -> ()
            | Some itemref ->
                worksheet.Range(Cell (entitem.ItemIndex, 6)).Value2 <- itemref.DataType
                worksheet.Range(Cell (entitem.ItemIndex, 6)).Interior.ColorIndex <- DataAttrColor
                worksheet.Range(Cell (entitem.ItemIndex, 6)).HorizontalAlignment <- Constants.xlLeft
                worksheet.Range(Cell (entitem.ItemIndex, 7)).Value2 <- itemref.DataLengthDsp
                worksheet.Range(Cell (entitem.ItemIndex, 7)).Interior.ColorIndex <- DataAttrColor
                worksheet.Range(Cell (entitem.ItemIndex, 7)).HorizontalAlignment <- Constants.xlLeft
                worksheet.Range(Cell (entitem.ItemIndex, InputRow)).NumberFormatLocal <- itemref.NumberFormat
                ignore <| worksheet.Range(Column entitem.ItemIndex).EntireColumn.AutoFit()
                let width = unbox<float> (worksheet.Range(Column entitem.ItemIndex).ColumnWidth)
                if width < itemref.ColumnWidth
                then worksheet.Range(Column entitem.ItemIndex).ColumnWidth <- itemref.ColumnWidth
                else ()
                let validation = worksheet.Range(Cell (entitem.ItemIndex, InputRow)).Validation
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
                                    , "=LENB(" + (Cell (entitem.ItemIndex, InputRow)) + ")"
                                      + itemref.ValidationCusmomOperator
                                      + (CellRA (entitem.ItemIndex, DataLengthRow)))
                                validation.IMEMode <- int XlIMEMode.xlIMEModeOff
                            | ("VARCHAR2", None) ->
                                validation.Add(
                                    XlDVType.xlValidateCustom
                                    , XlDVAlertStyle.xlValidAlertStop
                                    , XlFormatConditionOperator.xlBetween
                                    , "=LENB(" + (Cell (entitem.ItemIndex, InputRow)) + ")"
                                      + itemref.ValidationCusmomOperator
                                      + (CellRA (entitem.ItemIndex, DataLengthRow)))
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
                // printfn "%s,%s,%s" ent.PhysicalName entitem.PhysicalName entitem.NotNull
                match entitem.PkeyIndex with
                | None ->
                    match entitem.NotNull with
                    | "true"
                    | "TRUE" ->
                        worksheet.Range(Cell (entitem.ItemIndex, InputRow)).Interior.ColorIndex <- NotNullColor
                        validation.InputMessage <- "必須項目です。"
                    | _ -> ()
                | Some pkeyindex ->
                    worksheet.Range(Cell (entitem.ItemIndex, InputRow)).Interior.ColorIndex <- PkeyColor
                    validation.InputMessage <- "キー項目です。"
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
    // VALUE句をセットする。
    let offset = 1
    let mutable idx = offset
    for entitem in ent.EntityItems.Values do
        idx <- idx + 1
        match idx < ent.EntityItems.Count + offset with
        | true -> worksheet.Range(Cell (GetSqlPos ent idx)).Value2 <- "=" + (GetSqlValue entitem) + " & \",\" & " + Cell (GetSqlPos ent (idx + 1))
        | false -> worksheet.Range(Cell (GetSqlPos ent idx)).Value2 <- "=" + (GetSqlValue entitem)
    // 固定枠を設定
    ignore <| worksheet.Range(Cell (GetFreezePanesPos ent)).Select()
    app.ActiveWindow.FreezePanes <- true
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
let main (args:string[]) =
    let currDir = Directory.GetCurrentDirectory()
    let csvDir = 
        match args.Length with
        | 1 -> args.[0]
        | _ -> currDir + "\\" + "一覧出力"
    let outDir = currDir + "\\" + "OUT"
    let dicDomain = GetDomain csvDir
    let dicCondition = GetCondition csvDir
    let dicConditionItem = GetConditionItem csvDir
    let dicItem = GetItem csvDir
    let dicEntity = GetEntity csvDir
    let dicEntityItem = GetEntityItem csvDir
    let dicEntityIndex = GetEntityIndex csvDir
    ProcessConditionItem dicConditionItem dicCondition
    ProcessItem dicItem dicCondition
    ProcessEntityItem dicEntityItem dicEntityIndex dicItem dicCondition dicEntity
    ReNumberEntityItem dicEntity
    ignore <| MakeDirectory outDir
    for ent in dicEntity.Values do
        match (IsTarget ent) with
        | true ->
            printfn "エンティティ[%s]を処理中…" ent.PhysicalName
            CreateExcel ent outDir
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
