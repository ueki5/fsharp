#if INTERACTIVE
#r "Microsoft.Office.Interop.Excel.dll"
#endif
#if COMPILED
module ExcelAutomation
#endif

open System
open System.IO
open System.Runtime.InteropServices
open Microsoft.Office.Interop.Excel
open System.Text.RegularExpressions

let rec getFiles targetDir = 
    seq {   for file in Directory.GetFiles(targetDir) do  
               yield (file)
            for subDir in Directory.GetDirectories targetDir do  
                yield! (getFiles subDir) }

let MAX_COL = 256
let MAX_ROW = 65536

let rToN (n :int) = 
    let charArr = [|for c in 'A' .. 'Z' -> c|]
    let len = Array.length charArr
    let rec rToNSub rem result = //26進数もどきへの変換、0(="A")は[0],1(="B")は[1],26(="AA")は[0;0]となる 
        if rem < len then 
            rem :: result 
        else 
            let red = rem % len
            rToNSub ((rem - red)/len - 1 ) (red :: result) // -1 がミソ  
    let makeUpCharArr = 
        (rToNSub (n-1) []) // -1 は列が1から始まる為 
        |>  List.map (fun i -> charArr.[i]) 
        |> Array.ofList 
    new string(makeUpCharArr)

let myMap =
    [|for c in 'A' .. 'Z' -> c|] 
    |> Array.mapi (fun i c -> (c,i)) 
    |> Map.ofArray

let nToC (ntag : string) =
    let rec nToCSub (rem:list<char>) resC = 
        match rem with 
        |[] -> resC
        |h :: tl ->  
            let num = myMap.[(System.Char.ToUpper h)]
            nToCSub tl (26*(resC+1) + num)
    let c = (nToCSub (List.ofArray (ntag.ToCharArray())) -1)
    c+1

let nToRC (ntag : string) =
    let rec nToRCSub (rem:list<char>) resR resC   = 
        match rem with 
        |[] ->  
            (resR,resC) 
        |h :: tl when System.Char.IsDigit h -> 
            nToRCSub tl (10 * resR + System.Int32.Parse(h.ToString())) resC 
        |h :: tl ->  
            let num = myMap.[(System.Char.ToUpper h)] 
            nToRCSub tl resR (26*(resC+1) + num) 
    let (r,c) = (nToRCSub (List.ofArray (ntag.ToCharArray())) 0 -1)
    (r,c+1)

let set (ws:_Worksheet) (row:int) (col:int) v = 
    ws.Cells.Item(row,col) <- v
    // let range = ws.Cells.Item(row,col) :?> Range
    // range.Borders.[XlBordersIndex.xlDiagonalDown].LineStyle <- Constants.xlNone
    // range.Borders.[XlBordersIndex.xlDiagonalDown].LineStyle <- Constants.xlNone
    // range.Borders.[XlBordersIndex.xlDiagonalUp].LineStyle <- Constants.xlNone
    // range.Borders.[XlBordersIndex.xlEdgeLeft].LineStyle <- XlLineStyle.xlContinuous
    // range.Borders.[XlBordersIndex.xlEdgeLeft].Weight <- XlBorderWeight.xlThin
    // range.Borders.[XlBordersIndex.xlEdgeLeft].ColorIndex <- Constants.xlAutomatic
    // range.Borders.[XlBordersIndex.xlEdgeTop].LineStyle <- XlLineStyle.xlContinuous
    // range.Borders.[XlBordersIndex.xlEdgeTop].Weight <- XlBorderWeight.xlThin
    // range.Borders.[XlBordersIndex.xlEdgeTop].ColorIndex <- Constants.xlAutomatic
    // range.Borders.[XlBordersIndex.xlEdgeBottom].LineStyle <- XlLineStyle.xlContinuous
    // range.Borders.[XlBordersIndex.xlEdgeBottom].Weight <- XlBorderWeight.xlThin
    // range.Borders.[XlBordersIndex.xlEdgeBottom].ColorIndex <- Constants.xlAutomatic
    // range.Borders.[XlBordersIndex.xlEdgeRight].LineStyle <- XlLineStyle.xlContinuous
    // range.Borders.[XlBordersIndex.xlEdgeRight].Weight <- XlBorderWeight.xlThin
    // range.Borders.[XlBordersIndex.xlEdgeRight].ColorIndex <- Constants.xlAutomatic

let setH (ws:_Worksheet) (sRow:int) (sCol:int) (vSeq :seq<'a>) = 
    vSeq  
    |> Seq.mapi (fun i v -> (sRow,sCol + i ,v))  
    |> Seq.iter (fun (r,c,v) -> set ws r c v)
let setV (ws:_Worksheet) (sRow:int) (sCol:int) (vSeq :seq<'a>) = 
    vSeq  
    |> Seq.mapi (fun i v -> (sRow+i,sCol,v))  
    |> Seq.iter (fun (r,c,v) -> set ws r c v)

let get (ws:_Worksheet) (row:int) (col:int) = 
    (ws.Cells.Item(row,col) :?> Range).Value2
let getRange (ws:_Worksheet) (row:int) (col:int) = 
    (ws.Cells.Item(row,col) :?> Range)
let setNumberFormatLocal format sheet row col = 
    (getRange sheet row col).NumberFormatLocal <- format
// let getG<'a> (ws:_Worksheet) (row : int) (col :int) =
//     let value = get ws row col
//     match value with
//     | null -> None
//     | _    -> Some (value :?> 'a)
// let getFloat = getG<'float>
// let getInt  = getG<'int>
// let getString = getG<'string>
// let getGA1<'a> (ws:_Worksheet) (sR,sC) (eR,eC) = 
//     Array2D.init (eR - sR + 1) (eC - sC + 1) (fun r c -> getG<'a> ws (sR+r) (sC+c))
// let getGA2<'a> (ws:_Worksheet) (sA:string) (eA:string) = 
//     getGA1<'a> ws (nToRC sA) (nToRC eA)
let getH (ws:_Worksheet) (sRow:int) (sCol:int) = 
    seq {for i in 0 .. MAX_COL do yield (get ws sRow (sCol + i))}
let getV (ws:_Worksheet) (sRow:int) (sCol:int) = 
    seq {for i in 0 .. MAX_ROW do yield (get ws (sRow + i) sCol)}

let OpenWorkbook (app:ApplicationClass) (fileName:string) = 
    try
        let workbook = app.Workbooks.Open(fileName)
        Some workbook
    with
        | ex -> printfn "error:%s" (ex.ToString()); None
let OpenWorksheet (workbook:Workbook) (sheetName:string) = 
    try
        let worksheet = (workbook.Worksheets.[box sheetName] :?> _Worksheet)
        Some worksheet
    with
        | ex -> printfn "error:%s" (ex.ToString()); None

let IsNull value =
    if value = null
    then true
    else false

let flip f a b = f b a
