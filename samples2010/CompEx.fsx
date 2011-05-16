open System
open System.IO
let rec testFiles testdir =  //��1
    seq {   for testfile in Directory.GetFiles(testdir) do  //��2
               yield (testfile)
            for testsubdir in Directory.GetDirectories testdir do  //��3
                yield! (testFiles testsubdir) }

testFiles "d:\work"
