#if INTERACTIVE
#load "QuiQPro.fsx"
#load "FileReader.fsx"
#load "ListCondition.fsx"
#endif
#if COMPILED
module ListItem
#endif
open System
open System.Collections.Generic
open QuiQPro
open FileReader
open ListCondition
open ListConditionItem

let GetNumberFormat (len1:int) (len2:int) (s:string) =
    if len2 > 0 && len1 > len2
    then (String.replicate (len1 - len2) s) + "." + (String.replicate len2 s)
    else (String.replicate len1 s)
let EndWith src chk =
    let srclen = String.length src
    let chklen = String.length chk
    match (chklen > 0) && (srclen > chklen) with
    | true -> src.[(srclen - chklen) .. srclen - 1] = chk
    | false -> false
type Item = {
    PhysicalName:string
    ;LogicalName:string
    ;DomainPhysicalName:string
    ;DomainLogicalName:string
    ;DataType:string
    ;DataLength1:int
    ;DataLength2:int
    ;DataLengthDsp:string
    ;ColumnWidth:int
    ;ValidationCusmomOperator:string
    ;NumberFormat:string
    ;NumberFormatMin:string
    ;NumberFormatMax:string
    ;ConditionPhysicalName:string
    ;ConditionLogicalName:string
    ;mutable ConditionRef:(Condition option)
    ;Remarks:string
    ;InsID:string
    ;InsDate:string
    ;UpdID:string
    ;UpdDate:string
    }
    
let MakeListItem (ary2d:string[][]) =
    let objTbl = new Dictionary<string, Item>()
    let MakeObject' (ary:string []) =  {
        PhysicalName = ary.[0]
        LogicalName = ary.[1]
        DomainPhysicalName = ary.[2]
        DomainLogicalName = ary.[3]
        DataType = CnvDateType ary.[4]
        DataLength1 = TryInt ary.[5]
        DataLength2 = TryInt ary.[6]
        DataLengthDsp = ""
        ColumnWidth = 0
        ValidationCusmomOperator = ""
        NumberFormat = ""
        NumberFormatMin = ""
        NumberFormatMax = ""
        ConditionPhysicalName = ary.[7]
        ConditionLogicalName = ary.[8]
        ConditionRef = None
        Remarks = ary.[9]
        InsID = ary.[10]
        InsDate = ary.[11]
        UpdID = ary.[12]
        UpdDate = ary.[13]}
    let GetColumnWidth (len:int) =
        if len < 10 then 10
        elif len <= 40 then len
        else 40
    let GetValidationCusmomOperator (item:Item) =
        match item.DataType with
        | "CHAR" -> "="
        // | "CHAR" ->
        //     match (EndWith item.PhysicalName "_CD") || (EndWith item.PhysicalName "_KBN") || (EndWith item.PhysicalName "_BI" && item.DataLength1 = 8)  || (EndWith item.PhysicalName "_YMD" && item.DataLength1 = 8) with
        //     | true -> "="
        //     | false -> "<="
        | "VARCHAR2" -> "<="
        | _ -> ""
    let GetFormat(item:Item) =
        match item.DataType with
        | "NUMBER" -> (GetNumberFormat (item.DataLength2 + 1) item.DataLength2 "0") + "_ "
        | "CHAR" -> "@"
        | "VARCHAR2" -> "@"
        | _ -> ""
    let MakeObject (ary:string []) =
        ary
        |> MakeObject' 
        |> (fun item -> {item with
                           DataLengthDsp = LengthPairToString item.DataLength1 item.DataLength2
                           ColumnWidth = GetColumnWidth item.DataLength1
                           ValidationCusmomOperator = GetValidationCusmomOperator item
                           NumberFormat = GetFormat item
                           NumberFormatMin = "-" + (GetNumberFormat item.DataLength1 item.DataLength2 "9")
                           NumberFormatMax = GetNumberFormat item.DataLength1 item.DataLength2 "9"
                         })
            
    match (Array.length ary2d) > 1 with
    | false -> objTbl
    | true  -> 
        // 先頭の１行は項目見出しの為、捨てる
        let ary2d' = ary2d.[1..]
        let objAry = Array.map MakeObject ary2d'
        for obj in objAry do
            objTbl.Add(obj.PhysicalName, obj)
        objTbl
