open System
open System.IO
let rec testFiles testdir =  //Å¶1
    seq {   for testfile in Directory.GetFiles(testdir) do  //Å¶2
               yield (testfile)
            for testsubdir in Directory.GetDirectories testdir do  //Å¶3
                yield! (testFiles testsubdir) }

testFiles "d:\work"
