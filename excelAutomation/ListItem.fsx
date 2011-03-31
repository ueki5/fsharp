#if INTERACTIVE
#load "FileReader.fsx"
#endif
#if COMPILED
module ListItem
#endif
open FileReader
open System
open System.Collections.Generic

type Item = {
    PhysicalName:string
    ;LogicalName:string
    ;DomainPhysicalName:string
    ;DomainLogicalName:string
    ;DataType:string
    ;DataLength:int
    ;DataLength2:int
    ;ConditionPhysicalName:string
    ;ConditionLogicalName:string
    ;Remarks:string
    ;InsID:string
    ;InsDate:string
    ;UpdID:string
    ;UpdDate:string
    }

let MakeListItem (ary2d:string[][]) =
    let TryInt s = match Int32.TryParse(s) with
                         | (false, _) -> 0
                         | (true, n) -> n
    let objTbl = new Dictionary<string, Item>()
    let MakeObject (ary:string []) = {
        PhysicalName = ary.[0]
        LogicalName = ary.[1]
        DomainPhysicalName = ary.[2]
        DomainLogicalName = ary.[3]
        DataType = ary.[4]
        DataLength = TryInt ary.[5]
        DataLength2 = TryInt ary.[6]
        ConditionPhysicalName = ary.[7]
        ConditionLogicalName = ary.[8]
        Remarks = ary.[9]
        InsID = ary.[10]
        InsDate = ary.[11]
        UpdID = ary.[12]
        UpdDate = ary.[13]
        }
            
    match (Array.length ary2d) > 1 with
    | false -> objTbl
    | true  -> 
        // 先頭の１行は項目見出しの為、捨てる
        let ary2d' = ary2d.[1..]
        let objAry = Array.map MakeObject ary2d'
        for obj in objAry do
            objTbl.Add(obj.PhysicalName, obj)
        objTbl
