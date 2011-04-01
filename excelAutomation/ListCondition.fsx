#if INTERACTIVE
#load "QuiQPro.fsx"
#load "FileReader.fsx"
#load "ListConditionItem.fsx"
#endif
#if COMPILED
module ListCondition
#endif
open System
open System.Collections.Generic
open QuiQPro
open FileReader
open ListConditionItem

type Condition = {
    ConditionPhysicalName:string;
    ConditionLogicalName:string;
    FixedItemPosition:string;
    FixedItemValue:string;
    FixedItemLable:string;
    Remarks:string;
    InsID:string;
    InsDate:string;
    UpdID:string;
    UpdDate:string;
    ConditionItems:ConditionItemDictionary;
    }

type ConditionDictionary = Dictionary<string, Condition>
let MakeListCondition (ary2d:string[][]) =
    let TryInt s = match Int32.TryParse(s) with
                         | (false, _) -> 0
                         | (true, n) -> n
    let objTbl = new ConditionDictionary()
    let MakeObject (ary:string []) = {
        ConditionPhysicalName = ary.[0]
        ConditionLogicalName = ary.[1]
        FixedItemPosition = ary.[2]
        FixedItemValue = ary.[3]
        FixedItemLable = ary.[4]
        Remarks = ary.[5]
        InsID = ary.[6]
        InsDate = ary.[7]
        UpdID = ary.[8]
        UpdDate = ary.[9]
        ConditionItems = new ConditionItemDictionary()
        }
            
    match (Array.length ary2d) > 1 with
    | false -> objTbl
    | true  -> 
        // 先頭の１行は項目見出しの為、捨てる
        let ary2d' = ary2d.[1..]
        let objAry = Array.map MakeObject ary2d'
        for obj in objAry do
            objTbl.Add(obj.ConditionPhysicalName, obj)
        objTbl
let GetDropDownList (condItems:ConditionItemDictionary) =
    let mutable s = ""
    for cond in condItems do
        match s with
        | "" ->
            s <- cond.Value.ConditionValue + ":" + cond.Value.ConditionLabel
        | _ ->
            s <- s + "," + cond.Value.ConditionValue + ":" + cond.Value.ConditionLabel
    s
