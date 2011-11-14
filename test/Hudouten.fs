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

// �G���g���[�|�C���g
#if COMPILED
[<EntryPoint>]
#endif

let main (args:string[]) =
    if Array.length args <> 2
    then
        ignore <| printfn "usage:RemoveTmStamp.exe �f�B���N�g���� �t�@�C����"
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
        ignore <| printfn "�Ώۂ̃t�@�C���i%s�j�����݂��܂���B" (filename + ".YYYYMMDDHHMMSSFF")
        -1
    else
        if File.Exists filepath
        then
            ignore <| printfn "���Ƀt�@�C���i%s�j�����݂��܂��B" filename
            -1
        else
            let targetfile = Seq.head (Seq.sort files)
            File.Copy(targetfile, dirbackup + "\\" + Path.GetFileName(targetfile))
            File.Move(targetfile, filepath)
            ignore <| printfn "�t�@�C�����ύX�i%s��%s�j" (Path.GetFileName targetfile) filename
            0
