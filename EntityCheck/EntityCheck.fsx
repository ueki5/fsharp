#if INTERACTIVE
#r "Microsoft.Office.Interop.Excel.dll"
#load "ExcelAutomation.fsx"
#endif
#if COMPILED
module EntityCheck
#endif
open System
open System.IO
open System.Text
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
    let dirs =
        getFiles args.[0]
        |> Seq.filter (fun x -> x.ToUpper().EndsWith(".XLS"))
    let outdir = args.[1]
    for file in dirs do
        match OpenWorkbook app file with
        | None -> ()
        | Some inbook ->
            for (Some insheet) in (Seq.map (OpenWorksheet inbook) ["エンティティ詳細";]) do
                let startrow = 7
                let checkcol = 2
                let startcol = 2
                let endrow = Seq.length (Seq.takeWhile (not << IsNull) (getV insheet startrow checkcol)) + startrow - 1
                let outfile = (outdir + "\\" + file.Substring(args.[0].Length + 1) + ".txt"):string
                let w = new StreamWriter(outfile, false, Encoding.GetEncoding("Shift-JIS"))
                for row in startrow .. endrow do
                    outpos <- outpos + 1
                    let record =
                        (getH insheet row startcol)
                        |> Seq.take 50
                        |> Seq.toArray
                    w.Write(record.[0])
                    w.Write(",")
                    w.Write(record.[2])
                    w.Write(",")
                    w.Write(record.[4])
                    w.Write(",")
                    w.Write(record.[19])
                    w.Write(",")
                    w.Write(record.[32])
                    w.Write(",")
                    w.Write(record.[36])
                    w.Write(",")
                    w.Write(if (unbox<string> record.[40]) = "" then (box "0") else record.[40])
                    w.Write(",")
                    w.WriteLine(record.[44])
                w.Close();
            inbook.Close()
    // outbook.SaveAs(args.[1])
    // outbook.Close()
    app.UserControl <- false
    app.Quit()
    printfn "処理が終了しました。"
    0
