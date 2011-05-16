module TextReader
open System
open System.IO
open System.Text
open FParsec.Primitives
open FParsec.CharParsers
open FParsec.Error

let rec getFiles targetDir =  
    seq {   for file in Directory.GetFiles(targetDir) do  
               yield (file)
            for subDir in Directory.GetDirectories targetDir do  
                yield! (getFiles subDir) }

let FileToLine (filename:string) =
    let r = new StreamReader (filename, Encoding.GetEncoding("Shift-JIS"))
    r.ReadToEnd()

let LineToList (line:string) =
    line.ToCharArray()
    |> Array.toList

type ReadStatus = Normal
                | Quoted

let ListToLines list =
    let rec ListToLines' sta list line lines =
        match (sta, list) with
        | (Normal, []) -> if line = ""
                            then lines
                            else line::lines
                          |> List.rev
                          // |> List.toArray
        | (Normal, '\r'::'\n'::cs)
        | (Normal, '\n'::cs) -> (ListToLines' Normal cs "" (line::lines))
        | (Normal, '\"'::cs) -> (ListToLines' Quoted cs line lines)
        | (Normal, c::cs) -> (ListToLines' Normal cs (line + string c) lines)
        | (Quoted, '\"'::'\"'::cs) -> (ListToLines' Quoted cs (line + string '\"') lines)
        | (Quoted, '\"'::cs) -> (ListToLines' Normal cs line lines)
        | (Quoted, c::cs) -> (ListToLines' Quoted cs (line + string c) lines)
        | (_, _) -> ["Case is not match! line="; line]
    ListToLines' Normal list "" []

let FileToLines =
    FileToLine
    >> LineToList
    >> ListToLines

FileToLines "D:\\WORK\\CHK_GLOVIA.TXT"
for line in (FileToLines "D:\\WORK\\CHK_GLOVIA.TXT") do
    printfn "%s" line
done
