#if INTERACTIVE
#load "FileReader.fsx"
#load "ListItem.fsx"
#endif
#if COMPILED
module ListIndex
#endif
open FileReader
open System
open System.Collections.Generic
open ListItem

type Index = {
    EntityPhysicalName:string
   ;EntityLogicalName:string
   ;IdxPhysicalName:string
   ;IdxLogicalName:string
   ;IdxType:string
   ;EntItemPhysicalName:string
   ;EntItemLogicalName:string
   ;ItemIndex:int
   ;mutable ItemRef:(Item option)
   ;InsID:string
   ;InsDate:string
   ;UpdID:string
   ;UpdDate:string
   }

type IndexDictionary = Dictionary<string * string * string, Index>
let MakeListIndex (ary2d:string[][]) =
    let objTbl = new IndexDictionary()
    let MakeObject (ary:string []) = {
            EntityPhysicalName = ary.[0]
            EntityLogicalName = ary.[1]
            IdxPhysicalName = ary.[2]
            IdxLogicalName = ary.[3]
            IdxType = ary.[4]
            EntItemPhysicalName = ary.[5]
            EntItemLogicalName = ary.[6]
            ItemIndex = int ary.[7]
            ItemRef = None
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
            objTbl.Add((obj.EntityPhysicalName, obj.IdxPhysicalName, obj.EntItemPhysicalName), obj)
        objTbl
