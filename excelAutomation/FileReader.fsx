#if COMPILED
module FileReader
#endif
// #if INTERACTIVE
// #r "Microsoft.Office.Interop.Excel.dll"
// #endif

// For COMException 
// open Microsoft.Office.Interop.Excel
// open System.Runtime.InteropServices
open System
open System.IO
open System.Text

let rec FileToLine (filename:string) =
    let r = new StreamReader(filename, Encoding.GetEncoding("Shift-JIS"))
    r.ReadToEnd()
let LineToList (line:string) =
    line.ToCharArray()
    |> Array.toList
type ReadStatus = Normal
                | Quoted
let ListToLines list =
    let rec ListToLines' sta list line lines =
        match (sta, list) with
        | (Normal, []) -> if line = ""
                            then lines
                            else line::lines
                          |> List.rev
                          |> List.toArray
        | (Normal, '\r'::'\n'::cs)
        | (Normal, '\n'::cs) -> (ListToLines' Normal cs "" (line::lines))
        | (Normal, '\"'::cs) -> (ListToLines' Quoted cs line lines)
        | (Normal, c::cs) -> (ListToLines' Normal cs (line + string c) lines)
        | (Quoted, '\"'::'\"'::cs) -> (ListToLines' Quoted cs (line + string '\"') lines)
        | (Quoted, '\"'::cs) -> (ListToLines' Normal cs line lines)
        | (Quoted, c::cs) -> (ListToLines' Quoted cs (line + string c) lines)
        | (_, _) -> [|"Case is not match! line="; line|]
    ListToLines' Normal list "" []
let SplitToLines line =
    line
    |> LineToList
    |> ListToLines
let FileToLines =
    FileToLine
    >> LineToList
    >> ListToLines
let FileToArray filename =
    let lines = FileToLines filename
    List.toArray [ for line in lines -> line.Split([|','|])]

// [<EntryPoint>]
// let main (args : string[]) =
//     let printarray arys =
//         match arys with
//         | [||] -> ()
//         | _ -> (Array.iter (fun s -> printfn "%s" s) arys)
//     match Array.length args with
//     | 1 -> let arys = FileToArray args.[0]
//            ignore <| Array.map printarray arys
//            0
//     | _ -> -1
