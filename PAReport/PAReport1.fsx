#if INTERACTIVE
#r "Microsoft.Office.Interop.Excel.dll"
#load "ExcelAutomation.fsx"
#endif
#if COMPILED
module PAReport1
#endif
open System
open System.IO
// open System.Runtime.InteropServices
open Microsoft.Office.Interop.Excel
// open System.Text.RegularExpressions
open ExcelAutomation

// エントリーポイント
#if COMPILED
[<EntryPoint>]
#endif

let main (args:string[]) =
    let mutable outpos = 0
    let app = new ApplicationClass(Visible = false)
    app.DisplayAlerts <- false
    let outbook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet) 
    let outsheet = (outbook.Worksheets.[box 1] :?> _Worksheet) 
    outsheet.Name <- "PS～PT進捗"
    outpos <- outpos + 1
    setH outsheet outpos 1 ["";"";"";"";"";"";"";"";"";"";"";"";"";"";"";"";"";"";"";"";"";"";"";
        "プログラム設計書作成";"";"";"";"";"";"";"";
        "プログラミング";"";"";"";"";"";"";"";"";
        "単体テスト";
        ]
    outpos <- outpos + 1
    setH outsheet outpos 1 [
        "機能区分";
        "区分";
        "№";
        "ﾌﾟﾛｸﾞﾗﾑ名称";
        "システム";
        "担当";
        "進捗率";
        "ﾌﾟﾛｸﾞﾗﾑID";
        "PG種別";
        "新規改造";
        "想定規模(STEP)";
        "Ｃ＃";
        "PLSQL";
        "COBOL";
        "帳票";
        "合計";
        "予定改造率(%)";
        "予定STEP";
        "実績STEP";
        "難易度（0.8～1.2）";
        "実績改造率(%)";
        "予定評価STEP";
        "実績評価STEP";
        "作成者";
        "人日";
        "進捗率";
        "作成完了";
        "ﾚﾋﾞｭｰ完了";
        "作成完了";
        "ﾚﾋﾞｭｰ完了";
        "工数(時間)";
        "担当";
        "人日";
        "進捗率";
        "作成着手";
        "作成完了";
        "ﾚﾋﾞｭｰ完了";
        "作成完了";
        "ﾚﾋﾞｭｰ完了";
        "工数(時間)";
        "担当";
        "人日";
        "進捗率";
        "ﾃｽﾄ着手";
        "ﾃｽﾄ完了";
        "ﾚﾋﾞｭｰ完了";
        "ﾃｽﾄ完了";
        "ﾚﾋﾞｭｰ完了";
        "工数(時間)";
        "テスト項目数(a)";
        "バグ発生数(b)";
        "テスト項目数(a)";
        "バグ発生数(b)";
        "テスト項目数(ｃ)";
        "テスト密度(c)/(a)";
        "①";
        "②";
        "③";
        "④";
        "⑤";
        "⑥";
        "⑦";
        "⑧";
        "⑨";
        "バグ発生数(e)";
        "バグ密度(e)/(b)";
        "テストヒット率(%)(e)/(ｃ)";
        "テスト項目数(h)";
        "バグ発生数(I)";
        "特異点(j)";
        "総合評価";
        "備考";
        "対象";
        "担当";
        "予定日";
        "実施日";
        "指摘件数";
        "仕様書(h)";
        "PG(h)";
        "PT(h)";
        "予定日";
        "完了日";
        "対象";
        "担当";
        "予定日";
        "実施日";
        "指摘件数";
        "仕様書(h)";
        "PG(h)";
        "PT(h)";
        "予定日";
        "完了日";
        "予定日";
        "実績日";
        "対象";
        "ｿｰｽｺｰﾄﾞ診断送付";
        "ｿｰｽｺｰﾄﾞ診断対応";
        "ファイル名";
        "シート名";]
    let dirs =
        getFiles args.[0]
        |> Seq.filter (fun x -> x.ToUpper().EndsWith(".XLS"))
    for file in dirs do
        match OpenWorkbook app file with
        | None -> ()
        | Some inbook ->
            for (Some insheet) in (Seq.map (OpenWorksheet inbook) ["PS～PT進捗";]) do
                let startrow = 6
                let checkcol = 4
                let startcol = 1
                let endrow = Seq.length (Seq.takeWhile (not << IsNull) (getV insheet startrow checkcol)) + startrow - 1
                for row in startrow .. endrow do
                    outpos <- outpos + 1
                    let record =
                        (getH insheet row startcol)
                        |> Seq.take 97
                        |> flip Seq.append (seq [file.Substring(args.[0].Length + 1);])
                        |> flip Seq.append (seq [insheet.Name;])
                    setH outsheet outpos 1 record
                    Seq.iter (setNumberFormatLocal "0%" outsheet outpos) [7;17;21;26;34;43;55;66;67;]
                    Seq.iter (setNumberFormatLocal "yyyy/mm/dd" outsheet outpos) [27;28;29;30;35;36;37;38;44;45;46;47;48;75;76;81;82;85;86;91;92;93;94;96;97]
            inbook.Close()
    outbook.SaveAs(args.[1])
    outbook.Close()
    app.UserControl <- false
    app.Quit()
    printfn "処理が終了しました。"
    0
