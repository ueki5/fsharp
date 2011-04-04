#if INTERACTIVE
#load "QuiQPro.fsx"
#load "ExcelAutomation.fsx"
#load "FileReader.fsx"
#load "ListEntityItem.fsx"
#endif
#if COMPILED
module ListEntity
#endif
open System
open System.Collections.Generic
open QuiQPro
open ExcelAutomation
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
let entityitems (entity:Entity) = 
    let mutable entityitems' = ""
    for entitem in entity.EntityItems do
        ignore <| entityitems' <- AppendWith entityitems' "," (entitem.Value.PhysicalName)
    entityitems'
let itemvalues (entity:Entity) = 
    let mutable values' = ""
    for entitem in entity.EntityItems do
        ignore <| values'<- AppendWith values' "," (GetSqlValue entitem.Value)
    values'
let GetSqlPos (entity:Entity) offset =
    match entity.EntityItems.Count <= CommonColumnCount with
    | true  -> (1 + offset , InputRow)
    | false -> (InputColumnOffset + entity.EntityItems.Count - CommonColumnCount + offset, InputRow)
let GetInsertSql1 (entity:Entity) = "=\"INSERT INTO " + entity.PhysicalName + "(\" & " + (Cell (GetSqlPos entity 1)) + " & \") VALUES(\" & " + (Cell (GetSqlPos entity 2)) + "& \");\""
let GetInsertSql2 (entity:Entity) = entityitems entity
let IsTarget (entity:Entity) = (String.length entity.PhysicalName > 3) && (entity.PhysicalName.[0..2] <> "ZV_")
