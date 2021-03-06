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
let StartWith src chk =
    let srclen = String.length src
    let chklen = String.length chk
    match (chklen > 0) && (srclen > chklen) with
    | true -> src.[1  .. srclen] = chk
    | false -> false
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
    ;ColumnWidth:float
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
    
type ItemDictionary = Dictionary<string,Item>
let MakeListItem (ary2d:string[][]) =
    let objTbl = new ItemDictionary()
    let MakeObject' (ary:string []) =  {
        PhysicalName = ary.[0]
        LogicalName = ary.[1]
        DomainPhysicalName = ary.[2]
        DomainLogicalName = ary.[3]
        DataType = CnvDateType ary.[4]
        DataLength1 = TryInt ary.[5]
        DataLength2 = TryInt ary.[6]
        DataLengthDsp = ""
        ColumnWidth = 0.0
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
    let GetColumnWidth (len:float) =
        if len < 10.0 then 10.0
        elif len <= 40.0 then len
        else 40.0
    let GetValidationCusmomOperator (item:Item) =
        match item.DataType with
        // | "CHAR" -> "="
        | "CHAR" ->
            // 桁数完全一致とする項目
            if
                  (item.DomainPhysicalName <> "SHN_CD")
                  && (not (item.PhysicalName.Contains "SHN_CD") || item.DataLength1 <> 4)
                  && (item.DomainPhysicalName <> "TORIHIKI_ZYOUKEN_CD")
                  && (not (item.PhysicalName.Contains "TORIHIKI_ZYOUKEN_CD") || item.DataLength1 <> 4)
                  && (not (item.PhysicalName.Contains "ZIHANKI_BAIKA_ZYOUKEN_CD") || item.DataLength1 <> 4)
                  && (not (item.PhysicalName.Contains "KAMOKU_CD") || item.DataLength1 <> 5)
                  && ((item.PhysicalName.EndsWith "_CD")
                      || (item.PhysicalName.EndsWith "_KBN")
                      || (item.PhysicalName.EndsWith "_BI" && item.DataLength1 = 8)
                      || (item.PhysicalName.EndsWith "_YMD" && item.DataLength1 = 8)
                      || (item.PhysicalName.EndsWith "_HIDUKE" && item.DataLength1 = 8)
                      || (item.PhysicalName.Contains "_BI_" && item.DataLength1 = 8)
                      || (item.PhysicalName.Contains "_YMD_" && item.DataLength1 = 8)
                      || (item.PhysicalName.Contains "_HIDUKE_" && item.DataLength1 = 8)
                      || (item.DomainPhysicalName = "YYYYMMDD")
                      || (item.DomainPhysicalName = "YYYYMM")
                      || (item.DomainPhysicalName = "YYYY")
                      || (item.DomainPhysicalName = "YY")
                      || (item.DomainPhysicalName = "MM")
                      || (item.DomainPhysicalName = "DD")
                      || (item.DomainPhysicalName.StartsWith "DATETIME"))
                  && (item.PhysicalName <> "SIKIBETU_CD")
                  && (item.PhysicalName <> "GENERAL_KBN_CD")
                  && (item.PhysicalName <> "SYUBETU_CD")
                  && (item.PhysicalName <> "SIKIBETU_CD")
                  && (item.PhysicalName <> "KBN_CD")
            then "="
            // 上限のみ
            else "<="
        | "VARCHAR2" -> "<="
        | _ -> ""
    let GetFormat(item:Item) =
        match item.DataType with
        | "NUMBER" ->
            if item.DomainPhysicalName.StartsWith "KINGAKU"
               || item.DomainPhysicalName = "TANKA"
               || (item.PhysicalName.Contains "_KINGAKU" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_KIN8" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_KIN10" && item.DataLength1 = 10 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_KIN12" && item.DataLength1 = 12 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_KIN15" && item.DataLength1 = 15 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_ZAN8" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_ZAN10" && item.DataLength1 = 10 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_ZAN12" && item.DataLength1 = 12 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_ZAN15" && item.DataLength1 = 15 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_TAX" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_TAX" && item.DataLength1 = 10 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_TAX" && item.DataLength1 = 12 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_TAX" && item.DataLength1 = 15 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_ZANDAKA" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_ZANDAKA" && item.DataLength1 = 10 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_ZANDAKA" && item.DataLength1 = 12 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_ZANDAKA" && item.DataLength1 = 15 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_HIYOU" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_HIYOU" && item.DataLength1 = 10 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_HIYOU" && item.DataLength1 = 12 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_HIYOU" && item.DataLength1 = 15 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_RYOUKIN" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "BUHINDAI" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "SHNDAI" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "BUZAI" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_GETUGAKU" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_SOUGAKU" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_KOURYOU" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "_KAKAKU" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "TESURYOU" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "DENKIRYOU" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "SIYOURYOU" && item.DataLength1 = 8 && item.DataLength2 = 0)
               || (item.PhysicalName.Contains "TANKA" && item.DataLength1 = 8 && item.DataLength2 = 2)
               || (item.PhysicalName.Contains "HATARI" && item.DataLength1 = 8 && item.DataLength2 = 2)
               || (item.PhysicalName.Contains "KINGAKU" && item.DataLength1 = 8 && item.DataLength2 = 2)
               || (item.PhysicalName.Contains "TATENE" && item.DataLength1 = 8 && item.DataLength2 = 2)
            then "#,##" + (GetNumberFormat (item.DataLength2 + 1) item.DataLength2 "0")
                + ";[赤]-#,##" + (GetNumberFormat (item.DataLength2 + 1) item.DataLength2 "0")
            elif item.DomainPhysicalName = "TAX_RITU"
                 || (item.PhysicalName.Contains "TAX"
                     && item.PhysicalName.Contains "RITU"
                     && item.DataLength1 = 4
                     && item.DataLength2 = 3)
            then "0.0%"
            else (GetNumberFormat (item.DataLength2 + 1) item.DataLength2 "0") + "_ "
        | "CHAR" -> "@"
        | "VARCHAR2" -> "@"
        | _ -> ""
    let MakeObject (ary:string []) =
        ary
        |> MakeObject' 
        |> (fun item -> {item with
                           DataLengthDsp = LengthPairToString item.DataLength1 item.DataLength2
                           ColumnWidth = GetColumnWidth (float item.DataLength1)
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
