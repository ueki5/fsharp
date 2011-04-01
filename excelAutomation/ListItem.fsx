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
        // DataLengthDsp = LengthPairToString (TryInt ary.[5]) (TryInt ary.[6])
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
        | "VARCHAR2" -> "<="
        | _ -> ""
    let Padding (s:string) (len:int) =
        let mutable out = ""
        if len > 0
        then for i = 1 to len do
             out <- out + s
        else out <- ""
        out
    let GetNumberFormat (len1:int) (len2:int) (s:string) =
        if len2 > 0 && len1 > len2 && len1 >= 0
        then (Padding s (len1 - len2)) + "." + (Padding s len2)
        else (Padding s len1)
    let GetFormat(item:Item) =
        match item.DataType with
        | "NUMBER" -> (GetNumberFormat 1 item.DataLength2 "0") + "_ "
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
