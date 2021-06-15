open Fxcel

let read () = System.Console.ReadLine()
let toInt (s: string) = System.Convert.ToInt32(s)

//printf "アタッチするExcelを指定してください。---> "
//let ps = show()
//let index = read() |> toInt

//let ps = enumerate()
//let app = attach ps.[0]
//let sheet = app |> workbook(1) |> worksheet(1)

//sheet.["A1:A3"] |> set 100
//sheet.["B1"] |> fx "SUM(A1:A3)"
//printfn "%A" (get sheet.["B1"])

//sheet.["A1:A3"]
//|> address
//|> printfn "%s"

let ps = enumerate()
let app = attach ps.[0]
let sheet = app |> workbook(1) |> worksheet(1)

//sheet.["A1:B3"]
//|> gets
//|> iteri (fun i j row -> printfn $"[%d{i}, %d{j}] {row}")

sheet.["A1:B3"]
|> rows
|> iter (fun row -> printfn $"%A{row}")

sheet.["A1:B3"]
|> rowsi
|> iter (fun (i, row) -> printfn $"[%d{i}] %A{row}")

sheet.["A1:B3"]
|> columns
|> iter (fun col -> printfn $"%A{col}")

sheet.["A1:B3"]
|> columnsi
|> iter (fun (i, col) -> printfn $"[%d{i}] %A{col}")

detach app

//let range = sheet.["A1:B3"] |> gets
//let (x, y) = range |> len
//for i = 0 to x - 1 do
//  for j = 0 to y - 1 do
//    printf $"{range.[i, j]} "
//  printfn ""

//detach app
