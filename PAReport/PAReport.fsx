#if INTERACTIVE
#r "Microsoft.Office.Interop.Excel.dll"
#endif
#if COMPILED
module PAReport
#endif
open System
open System.IO
open System.Collections.Generic
open System.Runtime.InteropServices
open Microsoft.Office.Interop.Excel
// open ExcelAutomation

let set (ws:Worksheet) (row:int) (col:int) v = 
    ws.Cells.Item(row,col) <- v;; 
let setH (ws:Worksheet) (sRow:int) (sCol:int) (vSeq :seq<'a>) = 
    vSeq  
    |> Seq.mapi (fun i v -> (sRow,sCol + i ,v))  
    |> Seq.iter (fun (r,c,v) -> set ws r c v);; 
let setV (ws:Worksheet) (sRow:int) (sCol:int) (vSeq :seq<'a>) = 
    vSeq  
    |> Seq.mapi (fun i v -> (sRow+i,sCol,v))  
    |> Seq.iter (fun (r,c,v) -> set ws r c v);; 
let OpenExcel (app:ApplicationClass) (fileName:string) = 
    let workbooks = app.Workbooks
    let workbook = workbooks.Open(fileName)
    // let workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet) 
    let sheets = workbook.Worksheets 
    let worksheet = (sheets.[box 1] :?> Worksheet)
    set worksheet 1 2 "aaaa"
    // workbook.Close


// エントリーポイント
#if COMPILED
[<EntryPoint>]
#endif
let main (args:string[]) =
    let app = new ApplicationClass(Visible = true)
    app.DisplayAlerts <- false
    // let workbooks = app.Workbooks
    // let workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet) 
    // let sheets = workbook.Worksheets 
    // let worksheet = (sheets.[box 1] :?> _Worksheet) 
    OpenExcel app args.[0]
    // app.UserControl <- false
    // app.Quit()
    0
#if INTERACTIVE
main
#endif
