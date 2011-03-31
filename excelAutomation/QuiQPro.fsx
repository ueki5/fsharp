#if COMPILED
module QuiQPro
#endif

open System
let TryInt s = match Int32.TryParse(s) with
               | (false, _) -> 0
               | (true, n) -> n

let DateType s = match s with
                 | "9" -> "NUMBER"
                 | "X" -> "CHAR"
                 | "N" -> "VARCHAR2"
                 | _   -> ""

