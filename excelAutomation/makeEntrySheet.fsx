// Copyright (c) Microsoft Corporation 2005-2008.
// This sample code is provided "as is" without warranty of any kind. 
// We disclaim all warranties, either express or implied, including the 
// warranties of merchantability and fitness for a particular purpose. 
#if INTERACTIVE
#r "Microsoft.Office.Interop.Excel.dll"
#endif

// For COMException 
open Microsoft.Office.Interop.Excel
open System
open System.Runtime.InteropServices
open System.IO

//
// Create new Excel.Application
let app = new ApplicationClass(Visible = true) 
let workbooks = app.Workbooks
let workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet) 
let sheets = workbook.Worksheets 
let worksheet = (sheets.[box 1] :?> _Worksheet) 
// Console.WriteLine ("Setting the value for cell")

// This puts the value 5 to the cell
// Console.WriteLine (List.head lines)
let r = new StreamReader("csv/List_Entity_Item.csv",Encoding.GetEncoding("Shift-JIS")) ;;
let 

worksheet.Range("A1").Value2 <- line
worksheet.Range("A2").Value2 <- "エンティティ物理名"

worksheet.Range("A4", "E4").Value2 <- [| "項目論理名１";"項目論理名２";"項目論理名３";"項目論理名４";"項目論理名５" |]
worksheet.Range("A5", "E5").Value2 <- [| "項目物理名１";"項目物理名２";"項目物理名３";"項目物理名４";"項目物理名５" |]
worksheet.Range("A6", "E6").Value2 <- [| "属性１";"属性２";"属性３";"属性４";"属性５" |]
worksheet.Range("A7", "E7").Value2 <- [| "桁数１";"桁数２";"桁数３";"桁数４";"桁数５" |]
worksheet.Range("A8", "E8").Value2 <- [| "備考１";"備考２";"備考３";"備考４";"備考５" |]
// let Cell i j =
//     match i

let Cnv10To26 n =
    match n % 26 with
    |  1 -> "A"
    |  2 -> "B"
    |  3 -> "C"
    |  4 -> "D"
    |  5 -> "E"
    |  6 -> "F"
    |  7 -> "G"
    |  8 -> "H"
    |  9 -> "I"
    | 10 -> "J"
    | 11 -> "K"
    | 12 -> "L"
    | 13 -> "M"
    | 14 -> "N"
    | 15 -> "O"
    | 16 -> "P"
    | 17 -> "Q"
    | 18 -> "R"
    | 19 -> "S"
    | 20 -> "T"
    | 21 -> "U"
    | 22 -> "V"
    | 23 -> "W"
    | 24 -> "X"
    | 25 -> "Y"
    |  _ -> "Z"
let rec CnvNum2Alph n = (CnvNum2Alph (n / 26)) + (Cnv10To26 n)

#if COMPILED
System.Threading.Thread.Sleep(1000)

//  If user interacted with Excel it will not close when the app object is destroyed, so we close it explicitly 
let _ = 
  try 
    workbook.Saved <- true
    app.UserControl <- false
    app.Quit()
  with e -> Console.WriteLine ("User closed Excel manually, so we don't have to do that")

let _ = Console.WriteLine ("Sample successfully finished!")
#endif
