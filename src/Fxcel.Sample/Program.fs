open Fxcel

let read () = System.Console.ReadLine()
let toInt (s: string) = System.Convert.ToInt32(s)

printf "アタッチするExcelを指定してください。---> "
let ps = show()
let index = read() |> toInt

let app = attach ps.[index]
let sheet = app |> workbook(1) |> worksheet(1)

sheet.["A1:A3"] |> set 100
sheet.["B1"] |> fx "SUM(A1:A3)"
printfn "%A" (get sheet.["B1"])

app.Dispose()
