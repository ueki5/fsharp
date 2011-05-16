module App1
open System
open System.IO
open System.Text
open FParsec.Primitives
open FParsec.CharParsers
open FParsec.Error
// open ExtString.String

let test s = match run letter s with
             | Success (r,us,p)   -> printfn "success: %A" r
             | Failure (msg,err,us) -> printfn "failed: %s" msg

let _ = test "ふじこｌｐ777"
let _ = test "1ふじこｌｐ777"

let test2 s = match many1 letter |> run <| s with
              | Success (r, s, p) -> printfn "%A" r
              | Failure (msg, err, s) -> printfn "%s" msg

let _ = test2 "ふじこｌｐ777"
let _ = test2 "1ふじこｌｐ777"

let ld = (letter <|> digit |> many1Chars) 
         |> run <| "ABC12D34;EF5"
match ld with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

let ld' = parse {let! anc = choice [letter;digit] |> many1Chars
                 return anc}
          |> run <| "ABC12D34;EF5"
match ld' with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg
Console.WriteLine () |> ignore

let e1 = run (pchar '(' >>. (many1 digit) .>> pchar ')') "(123456)"
match e1 with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

let f = pchar '(' >>. (many1Chars digit) .>> pchar ')'
        |> run <| "(123456)"
match f with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

let notlower1 = (letter <|> anyOf "#+.") .>> manyChars lower |> manyChars
               |> run <| "F#C#C++javaVB.NET"
match notlower1 with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

let cn = (pipe2 letter (many1Chars digit) (fun x y -> string(x) + string(y)) |> manyStrings) 
         |> run <| "A1B2C345;D6"
match cn with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

let cn' = parse {let! anc = letter 
                 let! d = many1Chars digit 
                 return string(anc)+string(d)} |> manyStrings
          |> run <| "A1B2C345;D6"
match cn' with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

let cn'' = parse {let! anc = letter 
                  let! d = many1Chars digit <|> (anyOf ";" |> many1Chars)
                  return string(anc)+string(d)} |> manyStrings
           |> run <| "A1B2C3D;E45;F6"
match cn'' with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

let str = @"(*comme
nt123a*)bc4d*)"

let comment1 = pstring "(*" >>. many1Chars (choice [digit;letter;newline] ) .>> pstring "*)"
               |> run <| str
match comment1 with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

let comment2 = between (pstring "(*") (pstring "*)") (regex "[^*)]+") 
               |> run <| str
match comment2 with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

let kanji = regex "[一-龠]" <?> "kanji"
let kr = kanji |> many1 |> run <| "読書百遍意自ずから通ず"
match kr with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

let repeat n p =
  let rec repeat n p result =
    parse {if n > 0 then
             let! x = p
             let! xs = repeat (n - 1) p (result@[x])
             return xs
           else
             return result}
  repeat n p []

let d3 = letter <|> digit .>> many letter >>. repeat 3 digit
         |> run <| "aBc1234dEf"

match d3 with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

let w = repeat 3 (pstring "うぇ")
         |> run <| "うぇうぇうぇうぇうぇうえぇえぇええｗｗｗｗ"

match w with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

let w2 = repeat 3 (pstring "うぇうぇ")
         |> run <| "うぇうぇうぇうぇうぇうえぇえぇええｗｗｗ"

match w2 with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg


let w3 = attempt (repeat 3 (pstring "うぇうぇ")) <|> (repeat 5 (pstring "うぇ"))
         |> run <| "うぇうぇうぇうぇうぇうえぇえぇええｗｗｗ"

match w3 with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

let r = parse {let! d = regex "^\d{3}-\d{4}$" in return d}
let zip = (r, "001-0016") ||> run
match zip with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

// let zip2 = parse {let! p= pipe3 (repeat 3 digit) (pchar '-') (repeat 4 digit) (fun x y z -> x@[y]@z)
//                   do! notFollowedBy (digit <|> letter)
//                   return implode p}
//            |> run <| "001-0016"
// match zip2 with
// | Success (r, s, p) -> printfn "%A" r
// | Failure (msg, err, s) -> printfn "%s" msg

let pline =
   parse {let! first = anyChar 
          if first = '\n' then 
            return "" 
          else 
            let! txt = restOfLine 
            return (first.ToString()+txt)} 

let strings' = run (many pline) "\n\nHoge1\nFuga\n\nPiyo" 
match strings' with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg

let pline': Parser<string, unit> =
    fun state ->
       let mutable str = null
       let newState = state.SkipRestOfLine(true, &str)
       if not (LanguagePrimitives.PhysicalEquality state newState) then
           Reply(str, newState)
       else
           Reply(Error, NoErrorMessages, newState)

let strings'' = run (many pline') "\n\nHoge1\nFuga\n\nPiyo"
match strings'' with
| Success (r, s, p) -> printfn "%A" r
| Failure (msg, err, s) -> printfn "%s" msg
