#if INTERACTIVE
#r "Microsoft.Office.Interop.Excel.dll"
#load "ExcelAutomation.fsx"
#load "FileReader.fsx"
// #load "ListDomain.fsx"
#load "ListCondition.fsx"
// #load "ListConditionItem.fsx"
#load "ListItem.fsx"
#load "ListEntity.fsx"
// #load "ListEntityItem.fsx"
#load "ListEntityIndex.fsx"
#load "ListIndex.fsx"
#endif
#if COMPILED
module CalcInitialSize
#endif
open System
open System.IO
open System.Collections.Generic
open System.Runtime.InteropServices
open Microsoft.Office.Interop.Excel
open ExcelAutomation
open FileReader
// open ListDomain
open ListCondition
open ListConditionItem
open ListItem
open ListEntity
open ListEntityItem
open ListEntityIndex
open ListIndex

let CsvSubDir = "一覧出力"
let OutSubDir = "初期サイズ計算"
// ディレクトリを作成する
let MakeDirectory dirpath =
    if Directory.Exists(dirpath)
    then Directory.Delete(dirpath,true)
    else ()
    Directory.CreateDirectory(dirpath)

// コンディション一覧出力のCSVファイルから、コンディションの情報を読み込みます。
let GetCondition csvDir =
    let filename = csvDir + "\\" + "List_Condition.CSV"
    if File.Exists filename
    then Some(FileToArray filename |> MakeListCondition)
    else printfn "入力ファイル[%s]が見つかりません" filename
         None

// コンディション項目一覧出力のCSVファイルから、コンディション項目の情報を読み込みます。
let GetConditionItem csvDir =
    let filename = csvDir + "\\" + "List_Condition_Item.CSV"
    if File.Exists filename
    then Some(FileToArray filename |> MakeListConditionItem)
    else printfn "入力ファイル[%s]が見つかりません" filename
         None

// 項目一覧出力のCSVファイルから、項目の情報を読み込みます。
let GetItem csvDir =
    let filename = csvDir + "\\" + "List_Item.CSV"
    if File.Exists filename
    then Some(FileToArray filename |> MakeListItem)
    else printfn "入力ファイル[%s]が見つかりません" filename
         None

// エンティティ一覧出力のCSVファイルから、エンティティの情報を読み込みます。
let GetEntity csvDir =
    let filename1 = csvDir + "\\" + "Custom_List_Entity.CSV"
    let filename2 = csvDir + "\\" + "List_Entity.CSV"
    if File.Exists filename1
    then printfn "エンティティ一覧は、入力ファイル[%s]で処理します。" filename1
         Some(FileToArray filename1 |> MakeListEntity)
    elif File.Exists filename2
    then Some(FileToArray filename2 |> MakeListEntity)
    else printfn "入力ファイル[%s]が見つかりません" filename2
         None

// エンティティ項目一覧出力のCSVファイルから、エンティティ項目の情報を読み込みます。
let GetEntityItem csvDir =
    let filename = csvDir + "\\" + "List_Entity_Item.CSV"
    if File.Exists filename
    then Some(FileToArray filename |> MakeListEntityItem)
    else printfn "入力ファイル[%s]が見つかりません" filename
         None

// エンティティインデックス一覧出力のCSVファイルから、PKEYの情報を読み込みます。
let GetEntityIndex csvDir =
    let filename = csvDir + "\\" + "List_Entity_Index.CSV"
    if File.Exists filename
    then Some(FileToArray filename |> MakeListEntityIndex)
    else printfn "入力ファイル[%s]が見つかりません" filename
         None

// エンティティインデックス一覧出力のCSVファイルから、インデックスの情報を読み込みます。
let GetIndex csvDir =
    let filename = csvDir + "\\" + "List_Entity_Index.CSV"
    if File.Exists filename
    then Some(FileToArray filename |> MakeListIndex)
    else printfn "入力ファイル[%s]が見つかりません" filename
         None

// コンディション項目をコンディションに登録します。
let ProcessConditionItem (dicConditionItem:ConditionItemDictionary) (dicCondition:ConditionDictionary) =
    for conditem in dicConditionItem do
        if dicCondition.ContainsKey(conditem.Value.ConditionPhysicalName)
        then
            let cond = dicCondition.[conditem.Value.ConditionPhysicalName]
            let _ = cond.ConditionItems.Add(conditem.Key, conditem.Value)
            ()
        else ()

// 項目にコンディションが定義されている場合、項目情報にコンディション情報を付加します。
let ProcessItem (dicItem:ItemDictionary) (dicCondition:ConditionDictionary) =
    for item in dicItem do
        if dicCondition.ContainsKey(item.Value.ConditionPhysicalName)
        then
            let cond = dicCondition.[item.Value.ConditionPhysicalName]
            ignore(item.Value.ConditionRef <- Some(cond))
            ()
        else ()

// エンティティ項目に情報を付加します。
let ProcessEntityItem
    (dicEntityItem:EntityItemDictionary)
    (dicEntityIndex:EntityIndexDictionary)
    (dicItem:ItemDictionary)
    (dicCondition:ConditionDictionary)
    (dicEntity:EntityDictionary) =
    for entitem in dicEntityItem do
        // PKEY項目かどうかを判断します。
        if dicEntityIndex.ContainsKey(entitem.Key)
        then
            let indexitem = dicEntityIndex.[entitem.Key]
            ignore(entitem.Value.PkeyIndex <- Some(indexitem.ItemIndex))
            ()
        else ()
        // 項目情報を付加します。
        if dicItem.ContainsKey(entitem.Value.PhysicalName)
        then
            let item = dicItem.[entitem.Value.PhysicalName]
            ignore(entitem.Value.ItemRef <- Some(item))
            // 項目にコンディションが定義されていた場合、コンディション情報を付加します。
            match item.ConditionRef with
            | None -> ()
            | Some cond -> ignore(entitem.Value.ConditionRef <- Some(cond))
            ()
        else ()
        // エンティティ項目にコンディションが定義されていた場合、コンディション情報を付加します。
        // 既に項目情報で定義されていた場合も、上書きします。
        match entitem.Value.ConditionPhysicalName with
        | "" -> ()
        | _ ->
            if dicCondition.ContainsKey(entitem.Value.ConditionPhysicalName)
            then
                let cond = dicCondition.[entitem.Value.ConditionPhysicalName]
                ignore(entitem.Value.ConditionRef <- Some(cond))
                ()
            else ()
        // エンティティ項目をエンティティに登録します。
        if dicEntity.ContainsKey(entitem.Value.EntityPhysicalName)
        then
            let ent = dicEntity.[entitem.Value.EntityPhysicalName]
            let _ = ent.EntityItems.Add(entitem.Key, entitem.Value)
            ()
        else ()

// インデックス項目に情報を付加します。
let ProcessIndexItem
    (dicIndex:IndexDictionary)
    (dicItem:ItemDictionary) =
    for inditem in dicIndex do
        // 項目情報を付加します。
        if dicItem.ContainsKey(inditem.Value.EntItemPhysicalName)
        then
            let item = dicItem.[inditem.Value.EntItemPhysicalName]
            ignore(inditem.Value.ItemRef <- Some(item))
        else ()

// 一般の項目が前に、共通項目が後になるよう、番号を振り直します。
let ReNumberEntityItem (dicEntity:EntityDictionary) =
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

// セルに罫線を設定します。
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

// エンティティ単位に、エクセルファイルを作成します。
let WriteExcel (worksheet:_Worksheet) (ent:Entity) (dicIndex:IndexDictionary) initRow = 

    // データ属性から、１レコードあたりの必要容量を算出します。
    let mutable row = initRow
    let mutable totalSize = 0
    let mutable indexCount = 0
    let mutable indexCountWithoutPK = 0
    let objTbl = new Dictionary<string , (string * int * int)>()
    for entitem in ent.EntityItems.Values do
        // 項目を設定します。
        match entitem.ItemRef with
        | None -> ()
        | Some itemref ->
            let size = match itemref.DataType with
                | "NUMBER"     -> 1 + itemref.DataLength1 / 2
                | "CHAR"
                | "VARCHAR2" -> itemref.DataLength1
            totalSize <- totalSize + size
            // ignore <| (printfn "項目名:%s,データタイプ:%s,全体桁数:%d" entitem.PhysicalName,itemref.DataType,itemref.DataLength1)
            printfn "項目名:%s,データタイプ:%s,全体桁数:%d,サイズ:%d" entitem.PhysicalName itemref.DataType itemref.DataLength1 size
    // データ属性から、１レコードあたりのインデックス必要容量を算出します。
    for inditem in dicIndex do
        if ent.PhysicalName = inditem.Value.EntityPhysicalName
        then
            printfn "エンティティ物理名:%s" inditem.Value.EntityPhysicalName
            match inditem.Value.ItemRef with
            | None -> ()
            | Some itemref -> 
                let size = match itemref.DataType with
                    | "NUMBER"     -> 1 + itemref.DataLength1 / 2
                    | "CHAR"
                    | "VARCHAR2" -> itemref.DataLength1
                printfn "項目名:%s,データタイプ:%s,全体桁数:%d,サイズ:%d" itemref.PhysicalName itemref.DataType itemref.DataLength1 size
                if objTbl.ContainsKey(inditem.Value.IdxPhysicalName)
                then
                    let (idxType, colCount, initSize) = objTbl.[inditem.Value.IdxPhysicalName]
                    objTbl.[inditem.Value.IdxPhysicalName] <- (idxType , colCount + 1, initSize + size)
                else
                    objTbl.Add(inditem.Value.IdxPhysicalName, (inditem.Value.IdxType, 1, size)) 
        else ()
    for obj in objTbl do
        row <- row + 1
        indexCount <- indexCount + 1
        let (idxType, colCount, initSize) = obj.Value
        if idxType = "PK"
        then ()
        else indexCountWithoutPK <- indexCountWithoutPK + 1
        let numStr = match indexCount with
                       | 1 -> "①"
                       | 2 -> "②"
                       | 3 -> "③"
                       | 4 -> "④"
                       | 5 -> "⑤"
                       | 6 -> "⑥"
                       | 7 -> "⑦"
                       | 8 -> "⑧"
                       | 9 -> "⑨"
                       | 10 -> "⑩"
                       | 11 -> "⑪"
                       | 12 -> "⑫"
                       | 13 -> "⑬"
                       | 14 -> "⑭"
                       | 15 -> "⑮"
                       | 16 -> "⑯"
                       | 17 -> "⑰"
                       | 18 -> "⑱"
                       | 19 -> "⑲"
                       | 20 -> "⑳"
                       | _ -> "○"
        let suffix = if indexCountWithoutPK < 10
                     then "0" + (string indexCountWithoutPK)
                     else (string indexCountWithoutPK)
        let indexPhysicalName = if idxType = "PK"
                                then "PK_" + ent.PhysicalName
                                else 
                                     "IX_" + ent.PhysicalName + suffix
        worksheet.Range(Cell (1, row)).Value2 <- ent.PhysicalName
        worksheet.Range(Cell (2, row)).Value2 <- string totalSize
        worksheet.Range(Cell (3, row)).Value2 <- numStr
        worksheet.Range(Cell (4, row)).Value2 <- indexPhysicalName
        worksheet.Range(Cell (5, row)).Value2 <- string colCount
        worksheet.Range(Cell (6, row)).Value2 <- string initSize
    if row = initRow
    then
        row <- row + 1
        worksheet.Range(Cell (1, row)).Value2 <- ent.PhysicalName
        worksheet.Range(Cell (2, row)).Value2 <- string totalSize
    else ()
    row

// エントリーポイント
#if COMPILED
[<EntryPoint>]
#endif
let main (args:string[]) =
    let currDir = Directory.GetCurrentDirectory()
    let csvDir = 
        match args.Length with
        | 1 -> args.[0] + "\\" + CsvSubDir
        | 2 -> args.[0]
        | _ -> currDir + "\\" + CsvSubDir
    let outDir = 
        match args.Length with
        | 1 -> args.[0] + "\\" + OutSubDir
        | 2 -> args.[1]
        | _ -> currDir + "\\" + OutSubDir
    // match GetDomain csvDir with
    // | None -> -1
    // | Some dicCondition -> 
    match GetCondition csvDir with
    | None -> -1
    | Some dicCondition -> 
    match GetConditionItem csvDir with
    | None -> -1
    | Some dicConditionItem -> 
    match GetItem csvDir with
    | None -> -1
    | Some dicItem -> 
    match GetEntity csvDir with
    | None -> -1
    | Some dicEntity -> 
    match GetEntityItem csvDir with
    | None -> -1
    | Some dicEntityItem -> 
    match GetEntityIndex csvDir with
    | None -> -1
    | Some dicEntityIndex -> 
    match GetIndex csvDir with
    | None -> -1
    | Some dicIndex -> 
    ProcessConditionItem dicConditionItem dicCondition
    ProcessItem dicItem dicCondition
    ProcessEntityItem dicEntityItem dicEntityIndex dicItem dicCondition dicEntity
    ProcessIndexItem  dicIndex dicItem
    ReNumberEntityItem dicEntity
    let app = new ApplicationClass(Visible = false) 
    app.DisplayAlerts <- false
    ignore <| MakeDirectory outDir
    // Excel作成
    let workbooks = app.Workbooks
    let workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet) 
    let sheets = workbook.Worksheets 
    let worksheet = (sheets.[box 1] :?> _Worksheet)
    let mutable row = 0
    worksheet.Name <- "初期サイズ計算"
    for ent in dicEntity.Values do
        match (IsTarget ent) with
        | true ->
            printfn "エンティティ[%s]を処理中…" ent.PhysicalName
            row <- WriteExcel worksheet ent dicIndex row
        | false ->
            printfn "エンティティ[%s]をスキップしました。" ent.PhysicalName
    // app.ActiveWindow.FreezePanes <- true
    // ブック保存、終了
    workbook.SaveAs(outDir + "\\" + "CalcInitialSize.xls")
    workbook.Close()
    app.UserControl <- false
    app.Quit()
    0
#if INTERACTIVE
main
#endif
