#if INTERACTIVE
#load "QuiQPro.fsx"
#load "FileReader.fsx"
#endif
#if COMPILED
module ListConditionItem
#endif
open QuiQPro
open FileReader
open System
open System.Collections.Generic

type ConditionItem = {
     ConditionPhysicalName:string
    ;ConditionLogicalName:string
    ;ItemIndex:int
    ;ConditionItemId:string
    ;ConditionValue:string
    ;ConditionLabel:string
    ;InsID:string
    ;InsDate:string
    ;UpdID:string
    ;UpdDate:string
    }

type ConditionItemDictionary = Dictionary<string * string, ConditionItem>
let MakeListConditionItem (ary2d:string[][]) =
    let TryInt s = match Int32.TryParse(s) with
                         | (false, _) -> 0
                         | (true, n) -> n
    let objTbl = new ConditionItemDictionary()
    let MakeObject (ary:string []) = {
        ConditionPhysicalName = ary.[0]
        ConditionLogicalName = ary.[1]
        ItemIndex = TryInt ary.[2]
        ConditionItemId = ary.[3]
        ConditionValue = ary.[4]
        ConditionLabel = ary.[5]
        InsID = ary.[6]
        InsDate = ary.[7]
        UpdID = ary.[8]
        UpdDate = ary.[9]
        }
            
    match (Array.length ary2d) > 1 with
    | false -> objTbl
    | true  -> 
        // 先頭の１行は項目見出しの為、捨てる
        let ary2d' = ary2d.[1..]
        let objAry = Array.map MakeObject ary2d'
        for obj in objAry do
            objTbl.Add((obj.ConditionPhysicalName, obj.ConditionItemId), obj)
        objTbl
