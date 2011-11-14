#if INTERACTIVE
#load "QuiQPro.fsx"
#load "ExcelAutomation.fsx"
#load "FileReader.fsx"
#load "ListCondition.fsx"
#load "ListItem.fsx"
#endif
#if COMPILED
module ListEntityItem
#endif
open System
open System.Collections.Generic
open QuiQPro
open ExcelAutomation
open FileReader
open ListCondition
open ListItem

type EntityItem = {
    EntityPhysicalName:string
    ;EntityLogicalName:string
    ;mutable ItemIndex:int
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
let IsCommonItem name = 
        match name with
        | "X_INS_DATETIME"
        | "X_INS_USER_ID"
        | "X_INS_CLIENT_IP"
        | "X_INS_APSERVER_IP"
        | "X_INS_PG_ID"
        | "D_UPD_DATETIME"
        | "D_UPD_USER_ID"
        | "D_UPD_CLIENT_IP"
        | "D_UPD_APSERVER_IP"
        | "D_UPD_PG_ID"
        | "B_UPD_DATETIME"
        | "B_UPD_USER_ID"
        | "B_UPD_CLIENT_IP"
        | "B_UPD_APSERVER_IP"
        | "B_UPD_PG_ID" -> true
        | _ -> false
let GetSqlValue entityitem = 
        match entityitem.PhysicalName with
        | "X_INS_DATETIME"
        | "D_UPD_DATETIME"
        | "B_UPD_DATETIME" -> CellAA (CommonColumnOffset + 1, CommonRow)
        | "X_INS_USER_ID"
        | "D_UPD_USER_ID"
        | "B_UPD_USER_ID" -> WrapWith (CellAA (CommonColumnOffset + 2, CommonRow)) "'"
        | "X_INS_CLIENT_IP"
        | "D_UPD_CLIENT_IP"
        | "B_UPD_CLIENT_IP" -> CellAA (CommonColumnOffset + 3, CommonRow)
        | "X_INS_APSERVER_IP"
        | "D_UPD_APSERVER_IP"
        | "B_UPD_APSERVER_IP" -> CellAA (CommonColumnOffset + 4, CommonRow)
        | "X_INS_PG_ID"
        | "D_UPD_PG_ID"
        | "B_UPD_PG_ID" -> WrapWith (CellAA(CommonColumnOffset + 5, CommonRow)) "'"
        | _ ->
            match entityitem.ItemRef with
            | Some(item) when item.DataType = "NUMBER" ->
                let cellpos = Cell (InputColumnOffset + entityitem.ItemIndex, InputRow)
                "IF(len(" + cellpos + ")=0,\"Null\"," + cellpos + ")"
            | Some(item) when (item.DataType = "CHAR" || item.DataType = "VARCHAR2") ->
                let cellpos = Cell (InputColumnOffset + entityitem.ItemIndex, InputRow)
                WrapWith ("MIDB(" + cellpos + ",1," + (string item.DataLength1) + ")") "'"
            | _ -> ""
