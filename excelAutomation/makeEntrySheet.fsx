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

// Create new Excel.Application
let app = new ApplicationClass(Visible = true) 
let workbooks = app.Workbooks
let workbook = workbooks.Add(XlWBATemplate.xlWBATWorksheet) 
let sheets = workbook.Worksheets 
let worksheet = (sheets.[box 1] :?> _Worksheet) 
// Console.WriteLine ("Setting the value for cell")

// This puts the value 5 to the cell
worksheet.Cells(1,1).Value2 <- "エンティティ論理名"
// worksheet.Range("A1").Value2 <- "エンティティ論理名"
worksheet.Range("A2").Value2 <- "エンティティ物理名"

worksheet.Range("A4", "E4").Value2 <- [| "項目論理名１";"項目論理名２";"項目論理名３";"項目論理名４";"項目論理名５" |]
worksheet.Range("A5", "E5").Value2 <- [| "項目物理名１";"項目物理名２";"項目物理名３";"項目物理名４";"項目物理名５" |]
worksheet.Range("A6", "E6").Value2 <- [| "属性１";"属性２";"属性３";"属性４";"属性５" |]
worksheet.Range("A7", "E7").Value2 <- [| "桁数１";"桁数２";"桁数３";"桁数４";"桁数５" |]
worksheet.Range("A8", "E8").Value2 <- [| "備考１";"備考２";"備考３";"備考４";"備考５" |]

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
