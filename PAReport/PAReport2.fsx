#if INTERACTIVE
#r "Microsoft.Office.Interop.Excel.dll"
#load "ExcelAutomation.fsx"
#endif
#if COMPILED
module PAReport2
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
        "№";
        "指摘箇所";
        "指摘内容";
        "レビュー者";
        "レビュー日";
        "処置";
        "修正者";
        "修正日";
        "確認者";
        "完了日";
        "ファイル名";
        "シート名";]
    let dirs =
        getFiles args.[0]
        |> Seq.filter (fun x -> x.ToUpper().EndsWith(".XLS"))
    for file in dirs do
        match OpenWorkbook app file with
        | None -> ()
        | Some inbook -> 
            for (Some insheet) in (Seq.map (OpenWorksheet inbook) ["設計書";"仕様書";"ソース";"テスト結果 ";]) do
                let startrow = 7
                let checkcol = 3
                let startcol = 1
                let endrow = Seq.length (Seq.takeWhile (not << IsNull) (getV insheet startrow checkcol)) + startrow - 1
                for row in startrow .. endrow do
                    outpos <- outpos + 1
                    let record =
                        (getH insheet row startcol)
                        |> Seq.take 10
                        |> flip Seq.append (seq [file.Substring(args.[0].Length + 1);])
                        |> flip Seq.append (seq [insheet.Name;])
                    setH outsheet outpos 1 record
                    Seq.iter (setNumberFormatLocal "yyyy/mm/dd" outsheet outpos) [5; 8; 10;]
            inbook.Close()
    outbook.SaveAs(args.[1])
    outbook.Close()
    app.UserControl <- false
    app.Quit()
    printfn "処理が終了しました。"
    0
