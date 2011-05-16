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
// open ListDomain
open ListCondition
open ListConditionItem
open ListItem
open ListEntity
open ListEntityItem
open ListEntityIndex

let CsvSubDir = "一覧出力"
let OutSubDir = "登録シート"
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
let CreateExcel (app:ApplicationClass) (ent:Entity) (outDir:string) = 
    let workbooks = app.Workbooks
    let workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet) 
    let sheets = workbook.Worksheets 
    let worksheet = (sheets.[box 1] :?> _Worksheet) 

    worksheet.Name <- ent.PhysicalName
    // エンティティ項目毎に、入力欄を作成します。（共通項目はスキップします）
    for entitem in ent.EntityItems.Values do
        match IsCommonItem(entitem.PhysicalName) with
        | true     -> ()
        | false    ->
            // 罫線のセット
            for idx in [4 .. InputRow] do
                SetPattern (worksheet.Range (Cell (entitem.ItemIndex, idx)))
            // 項目論理名
            worksheet.Range(Cell (entitem.ItemIndex, 4)).Value2 <- entitem.LogicalName
            worksheet.Range(Cell (entitem.ItemIndex, 4)).Interior.ColorIndex <- TitleBackColor
            worksheet.Range(Cell (entitem.ItemIndex, 4)).Font.ColorIndex <- TitleFontColor
            worksheet.Range(Cell (entitem.ItemIndex, 4)).HorizontalAlignment <- Constants.xlLeft
            // 項目物理名
            worksheet.Range(Cell (entitem.ItemIndex, 5)).Value2 <- entitem.PhysicalName
            worksheet.Range(Cell (entitem.ItemIndex, 5)).Interior.ColorIndex <- TitleBackColor
            worksheet.Range(Cell (entitem.ItemIndex, 5)).Font.ColorIndex <- TitleFontColor
            worksheet.Range(Cell (entitem.ItemIndex, 5)).HorizontalAlignment <- Constants.xlLeft
            // 備考
            worksheet.Range(Cell (entitem.ItemIndex, 8)).Value2 <- entitem.Remarks
            worksheet.Range(Cell (entitem.ItemIndex, 8)).Interior.ColorIndex <- RemarksColor
            worksheet.Range(Cell (entitem.ItemIndex, 8)).WrapText <- true
            worksheet.Range(Cell (entitem.ItemIndex, 8)).HorizontalAlignment <- Constants.xlLeft

            // 項目を設定します。
            match entitem.ItemRef with
            | None -> ()
            | Some itemref ->
                // データ型
                worksheet.Range(Cell (entitem.ItemIndex, 6)).Value2 <- itemref.DataType
                worksheet.Range(Cell (entitem.ItemIndex, 6)).Interior.ColorIndex <- DataAttrColor
                worksheet.Range(Cell (entitem.ItemIndex, 6)).HorizontalAlignment <- Constants.xlLeft
                // データ長
                worksheet.Range(Cell (entitem.ItemIndex, 7)).Value2 <- itemref.DataLengthDsp
                worksheet.Range(Cell (entitem.ItemIndex, 7)).Interior.ColorIndex <- DataAttrColor
                worksheet.Range(Cell (entitem.ItemIndex, 7)).HorizontalAlignment <- Constants.xlLeft
                // 入力行
                worksheet.Range(Cell (entitem.ItemIndex, InputRow)).NumberFormatLocal <- itemref.NumberFormat
                // セル幅をセット
                ignore <| worksheet.Range(Column entitem.ItemIndex).EntireColumn.AutoFit()
                let width = unbox<float> (worksheet.Range(Column entitem.ItemIndex).ColumnWidth)
                if width < itemref.ColumnWidth
                then worksheet.Range(Column entitem.ItemIndex).ColumnWidth <- itemref.ColumnWidth
                else ()
                // データチェック属性を設定
                let validation = worksheet.Range(Cell (entitem.ItemIndex, InputRow)).Validation
                ignore(validation.Delete())
                ignore <| match (itemref.DataType, entitem.ConditionRef) with
                            // コンディション定義がある場合、リスト選択
                            | (_ , Some(cond)) ->
                                validation.Add(
                                    XlDVType.xlValidateList
                                    , XlDVAlertStyle.xlValidAlertStop
                                    , XlFormatConditionOperator.xlBetween
                                    , (GetDropDownList cond.ConditionItems))
                                validation.IMEMode <- int XlIMEMode.xlIMEModeOff
                            // 数値項目は最小、最大値
                            | ("NUMBER", None) ->
                                validation.Add(
                                    XlDVType.xlValidateDecimal
                                    , XlDVAlertStyle.xlValidAlertStop
                                    , XlFormatConditionOperator.xlBetween
                                    , itemref.NumberFormatMin
                                    , itemref.NumberFormatMax)
                                validation.IMEMode <- int XlIMEMode.xlIMEModeOff
                            // 文字列項目は、バイト数制約（最大値指定の時と、限定の場合がある）
                            | ("CHAR", None) -> 
                                validation.Add(
                                    XlDVType.xlValidateCustom
                                    , XlDVAlertStyle.xlValidAlertStop
                                    , XlFormatConditionOperator.xlBetween
                                    , "=LENB(" + (Cell (entitem.ItemIndex, InputRow)) + ")"
                                      + itemref.ValidationCusmomOperator
                                      + (CellRA (entitem.ItemIndex, DataLengthRow)))
                                validation.IMEMode <- int XlIMEMode.xlIMEModeOff
                            // 文字列項目は、最大バイト数制約
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
                // その他属性を設定
                validation.IgnoreBlank <- true
                validation.InCellDropdown <- true
                validation.InputTitle <- ""
                validation.ErrorTitle <- ""
                validation.InputMessage <- ""
                validation.ErrorMessage <- ""
                validation.ShowInput <- true
                validation.ShowError <- true
                // PKEY、NotNull項目は色をつけ、注釈を表示する。
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
    // ヘッダ欄設定
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
    // SQLを組み立てる参照式
    worksheet.Range(Cell (GetSqlPos ent 0)).Value2 <- (GetUpdateSql ent)
   // SELECT句、VALUE句をセットする。
    let offset = 0
    let mutable idx = offset
    for entitem in ent.EntityItems.Values do
        idx <- idx + 1
        match idx < ent.EntityItems.Count + offset with
        // | true -> worksheet.Range(Cell (GetColumnPos ent idx)).Value2 <- "=\"" + entitem.PhysicalName + "\" & \",\" & " + Cell (GetColumnPos ent (id
                  // worksheet.Range(Cell (GetSqlPos ent idx)).Value2 <- "=" + (GetSqlValue entitem) + " & \",\" & " + Cell (GetSqlPos ent (idx + 1)
        | true -> worksheet.Range(Cell (GetColumnPos ent idx)).Value2 <- "=\"" + entitem.PhysicalName + "\""
                  worksheet.Range(Cell (GetSqlPos    ent idx)).Value2 <- "="   + Cell (GetColumnPos ent idx) + " & \"=\" & " + (GetSqlValue entitem) + " & \",\" & " + Cell (GetSqlPos ent (idx + 1))
        | false -> worksheet.Range(Cell (GetColumnPos ent idx)).Value2 <- "=\"" + entitem.PhysicalName + "\""
                   worksheet.Range(Cell (GetSqlPos ent idx)).Value2 <- "=" + Cell (GetColumnPos ent idx) + " & \"=\" & " + (GetSqlValue entitem)
    // 固定枠を設定
    ignore <| worksheet.Range(Cell (GetFreezePanesPos ent)).Select()
    app.ActiveWindow.FreezePanes <- true
    // ブック保存、終了
    workbook.SaveAs(outDir + "\\" + ent.PhysicalName + "("+ ent.LogicalName + ")" + ".xls")
    workbook.Close()


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
    ProcessConditionItem dicConditionItem dicCondition
    ProcessItem dicItem dicCondition
    ProcessEntityItem dicEntityItem dicEntityIndex dicItem dicCondition dicEntity
    ReNumberEntityItem dicEntity
    let app = new ApplicationClass(Visible = false) 
    app.DisplayAlerts <- false
    ignore <| MakeDirectory outDir
    for ent in dicEntity.Values do
        match (IsTarget ent) with
        | true ->
            printfn "エンティティ[%s]を処理中…" ent.PhysicalName
            CreateExcel app ent outDir
        | false ->
            printfn "エンティティ[%s]をスキップしました。" ent.PhysicalName
    // Excel終了
    app.UserControl <- false
    app.Quit()
    0
#if INTERACTIVE
main
#endif
