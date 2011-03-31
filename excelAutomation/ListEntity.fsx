#if INTERACTIVE
#load "QuiQPro.fsx"
#load "FileReader.fsx"
#load "ListEntityItem.fsx"
#endif
#if COMPILED
module ListEntity
#endif
open System
open System.Collections.Generic
open QuiQPro
open FileReader
open ListEntityItem

type Entity = {
    PhysicalName:string;
    LogicalName:string;
    Remarks:string;
    InsID:string;
    InsDate:string;
    UpdID:string;
    UpdDate:string;
    EntityItems:EntityItemDictionary
    }
type EntityDictionary = Dictionary<string, Entity>
let MakeListEntity (ary2d:string[][]) =
    let objTbl = new EntityDictionary()
    let MakeObject (ary:string []) = {
        PhysicalName = ary.[0]
        LogicalName = ary.[1]
        Remarks = ary.[2]
        InsID = ary.[3]
        InsDate = ary.[4]
        UpdID = ary.[5]
        UpdDate = ary.[6]
        EntityItems = new EntityItemDictionary()
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
