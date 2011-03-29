// #if INTERACTIVE
// #r "Microsoft.Office.Interop.Excel.dll"
// #endif

// For COMException 
// open Microsoft.Office.Interop.Excel
// open System.Runtime.InteropServices
open System
open System.IO
open System.Text

let rec File2Line (filename:string) =
    let r = new StreamReader(filename, Encoding.GetEncoding("Shift-JIS"))
    let rec ReadLines (stream:StreamReader) (line:string) =
        let newline = stream.ReadLine()
        match newline with
        | null -> line
        | _    -> ReadLines stream (line + newline)
    ReadLines r ""
let LineToList (line:string) =
    let list = line.ToCharArray()
               |> Array.toList
    list
// let rec ListToLines list line lines =
//     match (list, line) with
//     | ([], "") -> lines
//     | ('\n'::cs) -> ListToLines cs 
    
// File2Lines "csv/List_Entity_Item.csv"
