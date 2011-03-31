#if INTERACTIVE
#load "QuiQPro.fsx"
#load "FileReader.fsx"
#endif
#if COMPILED
module ListEntityItem
#endif
open FileReader
open System
open System.Collections.Generic

type EntityItem = {
    EntPhysicalName:string
   ;EntLogicalName:string
   ;ItemIndex:int
   ;PhysicalName:string
   ;LogicalName:string
   ;FkFlg:string
   ;NotNull:string
   ;Default:string
   ;CondPhysicalName:string
   ;CondLogicalName:string
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
            EntPhysicalName = ary.[0]
            EntLogicalName = ary.[1]
            ItemIndex = int ary.[2]
            PhysicalName = ary.[3]
            LogicalName = ary.[4]
            FkFlg = ary.[5]
            NotNull = ary.[6]
            Default = ary.[7]
            CondPhysicalName = ary.[8]
            CondLogicalName = ary.[9]
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
            objTbl.Add((obj.EntPhysicalName, obj.PhysicalName), obj)
        objTbl
