// open System
// open System.IO
// let rec testFiles testdir =  //��1
//     seq {   for testfile in Directory.GetFiles(testdir) do  //��2
//                yield (testfile)
//             for testsubdir in Directory.GetDirectories testdir do  //��3
//                 yield! (testFiles testsubdir) }
// testFiles "d:\work"

let myFunction x y = x + 2 * y

printfn "Command line arguments: "

for arg in fsi.CommandLineArgs do
    printfn "%s" arg

printfn "%A" (myFunction 10 40)
