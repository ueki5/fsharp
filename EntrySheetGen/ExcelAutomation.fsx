#if INTERACTIVE
#r "Microsoft.Office.Interop.Excel.dll"
#endif
#if COMPILED
module ExcelAutomation
#endif

// For COMException 
//open Microsoft.Office.Interop.Excel
open System
open System.IO
//open System.Runtime.InteropServices

let PkeyColor = 38
let NotNullColor = 6
let TitleBackColor = 11
let TitleFontColor = 2
let DataAttrColor = 37
let RemarksColor = 24

let CommonColumnCount = 15 - 1
let CommonColumnOffset = 1
let CommonRow = 2
let InputColumnOffset = 0
let InputRow = 9
let DataLengthRow = 7
let ColumnPhysicalNameRow = 5

let NumToAlph n =
    match n with
    |  0 -> "A"
    |  1 -> "B"
    |  2 -> "C"
    |  3 -> "D"
    |  4 -> "E"
    |  5 -> "F"
    |  6 -> "G"
    |  7 -> "H"
    |  8 -> "I"
    |  9 -> "J"
    | 10 -> "K"
    | 11 -> "L"
    | 12 -> "M"
    | 13 -> "N"
    | 14 -> "O"
    | 15 -> "P"
    | 16 -> "Q"
    | 17 -> "R"
    | 18 -> "S"
    | 19 -> "T"
    | 20 -> "U"
    | 21 -> "V"
    | 22 -> "W"
    | 23 -> "X"
    | 24 -> "Y"
    | 25 -> "Z"
    |  _ -> ""
let rec PosX col =
    let col' = col - 1
    let division = (col - 1) / 26
    let remainder = (col - 1) % 26
    match division with
    | 0 -> NumToAlph remainder
    | _ -> PosX division + NumToAlph remainder
let PosY row = string row
let CellRR (col, row) = (PosX col) + (PosY row)
let CellRA (col, row) = (PosX col) + "$" + (PosY row)
let CellAR (col, row) = "$" + (PosX col) + (PosY row)
let CellAA (col, row) = "$" + (PosX col) + "$" + (PosY row)
let Cell = CellRR
let ColumnR col = PosX col + ":" + PosX col
let ColumnA col = "$" + (PosX col) + ":" + "$" + (PosX col)
let Column = ColumnR
let RowR row = PosY row + ":" + PosY row
let RowA row = "$" + (PosY row) + ":" + "$" + (PosY row)
let Row = RowR
