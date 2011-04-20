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
let GetColumnPos (entity:Entity) offset =
    match entity.EntityItems.Count <= CommonColumnCount with
    | true  -> (1 + offset , ColumnPhysicalNameRow)
    | false -> (InputColumnOffset + entity.EntityItems.Count - CommonColumnCount + offset, ColumnPhysicalNameRow)
let GetSqlPos (entity:Entity) offset =
    match entity.EntityItems.Count <= CommonColumnCount with
    | true  -> (1 + offset , InputRow)
    | false -> (InputColumnOffset + entity.EntityItems.Count - CommonColumnCount + offset, InputRow)
let GetInsertSql (entity:Entity) = "=\"INSERT INTO " + entity.PhysicalName + "(\" & " + (CellRA (GetColumnPos entity 1)) + " & \") VALUES(\" & " + (Cell (GetSqlPos entity 1)) + "& \");\""
let IsTarget (entity:Entity) = (String.length entity.PhysicalName > 3) && (entity.PhysicalName.[0..2] <> "ZV_")
let GetFreezePanesPos (entity:Entity) =
    let mutable pos = 0
    for entitem in entity.EntityItems.Values do
        match (pos, entitem.PkeyIndex) with
        | (0, None) -> pos <- entitem.ItemIndex
        | _ -> ()
    if pos > 0
    then (pos, InputRow)
    else (  1, InputRow)
