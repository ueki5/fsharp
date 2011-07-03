#if INTERACTIVE
#r "Microsoft.Office.Interop.Excel.dll"
#load "ExcelAutomation.fsx"
#endif
#if COMPILED
module PAReport3
#endif
open System
open System.IO
open System.Runtime.InteropServices
open Microsoft.Office.Interop.Excel
open System.Text.RegularExpressions
open ExcelAutomation

// エントリーポイント
#if COMPILED
[<EntryPoint>]
#endif

let main (args:string[]) =
    let mutable outpos = 1
    let app = new ApplicationClass(Visible = false)
    app.DisplayAlerts <- false
    let outbook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet) 
    let outsheet = (outbook.Worksheets.[box 1] :?> _Worksheet) 
    outsheet.Name <- "レビュー表"
    setH outsheet outpos 1 [
        "No";
        "ﾚﾋﾞｭｰｱ";
        "担当";
        "区分";
        "機能名";
        "シート";
        "関連No";
        "レビュー実施日";
        "指摘事項";
        "対応";
        "横展開対象";
        "PS仕様書修正有無";
        "PG修正有無";
        "PT仕様書修正有無";
        "追加ﾃｽﾄ有無";
        "対応予定日";
        "対応完了日";
        "対応確認者";
        "見解修正なし";
        "協力会社見解の確認（青：対応なし、赤：対応）";
        "ファイル名";
        "シート名";]
    let dirs =
        getFiles args.[0]
        |> Seq.filter (fun x -> x.ToUpper().EndsWith(".XLS"))
    for file in dirs do
        match OpenWorkbook app file with
        | None -> ()
        | Some inbook -> 
            for (Some insheet) in (Seq.map (OpenWorksheet inbook) ["PT仕様書指摘一覧";"エビデンス指摘一覧";]) do
                let startrow = 6
                let checkcol = 10
                let startcol = 2
                let endrow = Seq.length (Seq.takeWhile (not << IsNull) (getV insheet startrow checkcol)) + startrow - 1
                for row in startrow .. endrow do
                    outpos <- outpos + 1
                    let record =
                        (getH insheet row startcol)
                        |> Seq.take 20
                        |> flip Seq.append (seq [file.Substring(args.[0].Length + 1);])
                        |> flip Seq.append (seq [insheet.Name;])
                    setH outsheet outpos 1 record
                    Seq.iter (setNumberFormatLocal "yyyy/mm/dd" outsheet outpos) [8;16;17]
            inbook.Close()
    outbook.SaveAs(args.[1])
    outbook.Close()
    app.UserControl <- false
    app.Quit()
    printfn "処理が終了しました。"
    0
