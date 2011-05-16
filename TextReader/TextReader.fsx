#if COMPILATION
module TextReader
#end if
open System
open System.IO

let rec getFiles targetDir = 
    seq {   for file in Directory.GetFiles(targetDir) do  
                yield (file)
            for subDir in Directory.GetDirectories targetDir do  
                yield! (getFiles subDir) }
printfn "Command line arguments: start"
for arg in fsi.CommandLineArgs do
    printfn "%s" arg
printfn "Command line arguments: end"
Seq.iter (fun x -> printfn "%s" x) (getFiles "d:\work")
