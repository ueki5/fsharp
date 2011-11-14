#if COMPILED
module RemoveTmStamp
#endif
open System
open System.IO
open System.Text
// open Seq

let StartWith src chk = 
    let srclen = String.length src
    let chklen = String.length chk
    match (chklen > 0) && (srclen > chklen) with
    | true -> src.[0 .. chklen - 1] = chk
    | false -> false

let getFiles targetDir = 
    seq {   for file in Directory.GetFiles(targetDir) do  
                yield (file)}

// エントリーポイント
#if COMPILED
[<EntryPoint>]
#endif

let main (args:string[]) =
    if Array.length args <> 2
    then
        ignore <| printfn "usage:RemoveTmStamp.exe ディレクトリ名 ファイル名"
        -1
    else
    let dirname = args.[0]
    let dirbackup = dirname + "\\" + "ot_backup"
    let filename = args.[1]
    let filepath = dirname + "\\" + filename
    if Directory.Exists(dirbackup)
    then ()
    else ignore <| Directory.CreateDirectory(dirbackup)
    let files =
        getFiles dirname
        |> Seq.filter (fun x -> StartWith x filepath)
        |> Seq.filter (fun x -> String.length x = String.length (filepath + ".YYYYMMDDHHMMSSFF"))
    if (Seq.isEmpty files)
    then
        ignore <| printfn "対象のファイル（%s）が存在しません。" (filename + ".YYYYMMDDHHMMSSFF")
        -1
    else
        if File.Exists filepath
        then
            ignore <| printfn "既にファイル（%s）が存在します。" filename
            -1
        else
            let targetfile = Seq.head (Seq.sort files)
            File.Copy(targetfile, dirbackup + "\\" + Path.GetFileName(targetfile))
            File.Move(targetfile, filepath)
            ignore <| printfn "ファイル名変更（%s⇒%s）" (Path.GetFileName targetfile) filename
            0
