#if INTERACTIVE
#load "FileReader.fsx"
#endif
#if COMPILED
module ListEntityIndex
#endif
open FileReader
open System
open System.Collections.Generic

type EntityIndex = {
    EntityPhysicalName:string
   ;EntityLogicalName:string
   ;IdxPhysicalName:string
   ;IdxLogicalName:string
   ;IdxType:string
   ;EntItemPhysicalName:string
   ;EntItemLogicalName:string
   ;ItemIndex:int
   ;InsID:string
   ;InsDate:string
   ;UpdID:string
   ;UpdDate:string
   }

let MakeListEntityIndex (ary2d:string[][]) =
    let objTbl = new Dictionary<string * string, EntityIndex>()
    let MakeObject (ary:string []) = {
            EntityPhysicalName = ary.[0]
            EntityLogicalName = ary.[1]
            IdxPhysicalName = ary.[2]
            IdxLogicalName = ary.[3]
            IdxType = ary.[4]
            EntItemPhysicalName = ary.[5]
            EntItemLogicalName = ary.[6]
            ItemIndex = int ary.[7]
            InsID = ary.[8]
            InsDate = ary.[9]
            UpdID = ary.[10]
            UpdDate = ary.[11]
            }
            
    match (Array.length ary2d) > 1 with
    | false -> objTbl
    | true  -> 
        // 先頭の１行は項目見出しの為、捨てる
        let ary2d' = ary2d.[1..]
        let objAry = Array.map MakeObject ary2d'
        for obj in objAry do
            match obj.IdxType with
            | "PK" -> objTbl.Add((obj.EntityPhysicalName, obj.EntItemPhysicalName), obj)
            | _    -> ()
        objTbl
