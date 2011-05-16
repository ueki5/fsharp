open System
open System.IO
open System.Text.RegularExpressions

let str = @"こんにちは みなさん おげんき ですか？ わたしは げんき です。
この ぶんしょう は いぎりす の ケンブリッジ だいがく の けんきゅう の けっか
にんげん は もじ を にんしき する とき その さいしょ と さいご の もじさえ あっていれば
じゅんばん は めちゃくちゃ でも ちゃんと よめる という けんきゅう に もとづいて
わざと もじの じゅんばん を いれかえて あります。
どうです？ ちゃんと よめちゃう でしょ？
ちゃんと よめたら はんのう よろしく"

let cambrigde (s : string) =
    let rand = new Random()
    let rec convert (s : string) = 
        let shuffle =
            s
            |> Seq.sortBy (ignore >> rand.Next)
            |> Seq.toArray
        if s.Length <= 1
        then s.ToCharArray()
        elif s.ToCharArray() <> shuffle
        then shuffle
        else convert s 
    Regex.Replace(s, @"(?<=\w)\w{2,}(?=\w)", fun (m : Match) -> new string(m.Value |> convert))

str |> cambrigde |> printfn "%s"
Console.Read()


let rec testFiles testdir =  //※1
    seq {   for testfile in Directory.GetFiles(testdir) do  //※2
               yield (testfile)
            for testsubdir in Directory.GetDirectories testdir do  //※3
                yield! (testFiles testsubdir) };;
