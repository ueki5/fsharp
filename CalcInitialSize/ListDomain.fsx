#if INTERACTIVE
#load "QuiQPro.fsx"
#load "FileReader.fsx"
#endif
#if COMPILED
module ListDomain
#endif
open QuiQPro
open FileReader
open System
open System.Collections.Generic

type Domain = {
      PhysicalName:string
      ;LogicalName:string
      ;DataType:string
      ;DataLength:int
      ;DataLength2:int
      ;ConditionPhysicalName:string
      ;ConditionLogicalName:string
      ;Remarks:string
      ;InsID:string
      ;InsDate:string
      ;UpdID:string
      ;UpdDate:string
      }
type DomainDictionary = Dictionary<string, Domain>
let MakeListDomain (ary2d:string[][]) =
    let TryInt s = match Int32.TryParse(s) with
                         | (false, _) -> 0
                         | (true, n) -> n
    let objTbl = new DomainDictionary()
    let MakeObject (ary:string []) = {
        PhysicalName = ary.[0]
        LogicalName = ary.[1]
        DataType = ary.[2]
        DataLength = TryInt ary.[3]
        DataLength2 = TryInt ary.[4]
        ConditionPhysicalName = ary.[5]
        ConditionLogicalName = ary.[6]
        Remarks = ary.[7]
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
            objTbl.Add(obj.PhysicalName, obj)
        objTbl
