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

// エントリーポイント
#if COMPILED
[<EntryPoint>]
#endif
let main (args:string[]) =
    let app = new ApplicationClass(Visible = false) 
    app.DisplayAlerts <- false
    
    app.UserControl <- false
    app.Quit()
    0
#if INTERACTIVE
main
#endif
