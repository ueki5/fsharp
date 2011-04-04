#if COMPILED
module QuiQPro
#endif

open System
let TryInt s = match Int32.TryParse(s) with
               | (false, _) -> 0
               | (true, n) -> n
let LengthPairToString len1 len2 = match len2 with
                                   | 0 -> string len1
                                   | _ -> (string len1) + "," + (string len2)

let CnvDateType s = match s with
                    | "9" -> "NUMBER"
                    | "X" -> "CHAR"
                    | "N" -> "VARCHAR2"
                    | _   -> ""

let AppendWith s1 delim s2 =
    match s1 with
    | "" -> s2
    | _  -> s1 + delim + s2
let WrapWith (s:string) (wrap:string) = "\"" + wrap + "\" & " + s + " & \"" + wrap + "\""
