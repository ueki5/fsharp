#if INTERACTIVE
#load "QuiQPro.fsx"
#load "FileReader.fsx"
#load "ListCondition.fsx"
#load "ListItem.fsx"
#endif
#if COMPILED
module ListEntityItem
#endif
open System
open System.Collections.Generic
open FileReader
open ListCondition
open ListItem

type EntityItem = {
    EntityPhysicalName:string
    ;EntityLogicalName:string
    ;ItemIndex:int
    ;mutable PkeyIndex:(int option)
    ;PhysicalName:string
    ;LogicalName:string
    ;mutable ItemRef:(Item option)
    ;FkFlg:string
    ;NotNull:string
    ;Default:string
    ;ConditionPhysicalName:string
    ;ConditionLogicalName:string
    ;mutable ConditionRef:(Condition option)
    ;Remarks:string
    ;InsID:string
    ;InsDate:string
    ;UpdID:string
    ;UpdDate:string
    }
type EntityItemDictionary = Dictionary<string * string,EntityItem>
let MakeListEntityItem (ary2d:string[][]) =
    let objTbl = new EntityItemDictionary()
    let MakeObject (ary:string []) = {
            EntityPhysicalName = ary.[0]
            EntityLogicalName = ary.[1]
            ItemIndex = int ary.[2]
            PkeyIndex = None
            PhysicalName = ary.[3]
            LogicalName = ary.[4]
            ItemRef = None
            FkFlg = ary.[5]
            NotNull = ary.[6]
            Default = ary.[7]
            ConditionPhysicalName = ary.[8]
            ConditionLogicalName = ary.[9]
            ConditionRef = None
            Remarks = ary.[10]
            InsID = ary.[11]
            InsDate = ary.[12]
            UpdID = ary.[13]
            UpdDate = ary.[14]
            }
            
    match (Array.length ary2d) > 1 with
    | false -> objTbl
    | true  -> 
        // 先頭の１行は項目見出しの為、捨てる
        let ary2d' = ary2d.[1..]
        let objAry = Array.map MakeObject ary2d'
        for obj in objAry do
            objTbl.Add((obj.EntityPhysicalName, obj.PhysicalName), obj)
        objTbl
