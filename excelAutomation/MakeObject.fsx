#if COMPILED
module MakeObject
#endif
open FileReader
// #if INTERACTIVE
// #r "Microsoft.Office.Interop.Excel.dll"
// #endif

// For COMException 
// open Microsoft.Office.Interop.Excel
// open System.Runtime.InteropServices
open System

let MakeListEntityItem (ary:string[][]) =
    // æ“ª‚Ì‚Ps‚Í€–ÚŒ©o‚µ‚Ìˆ×AÌ‚Ä‚é
    let ary' = ary.[1..]
    ary'

[<EntryPoint>]
let main (args : string[]) =
    let printarray arys =
        match arys with
        | [||] -> ()
        | _ -> (Array.iter (fun s -> printfn "%s" s) arys)
    match Array.length args with
    | 1 -> let arys = FileToArray args.[0]
           ignore <| Array.map printarray arys
           0
    | _ -> -1
