open Fxcel
open Midoliy.Office.Interop

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

sheet.["A1:A3"] |> set 100
sheet.["B1"] |> fx "SUM(A1:A3)"

//sheet.["A1:B3"]
//|> gets
//|> iteri (fun i j row -> printfn $"[%d{i}, %d{j}] {row}")

//sheet.["A1:B3"]
//|> rows
//|> iter (fun row -> printfn $"%A{row}")

//sheet.["A1:B3"]
//|> rowsi
//|> iter (fun (i, row) ->  printfn $"[%d{i}] %A{row}")

//sheet.["A1:B3"]
//|> columns
//|> iter (fun col -> printfn $"%A{col}")

//sheet.["A1:B3"]
//|> columnsi
//|> iter (fun (i, col) -> printfn $"[%d{i}] %A{col}")


////let enumerator = sheet.["A1:B3"].GetEnumerator()
////printfn $"%A{enumerator.Current}"
////enumerator.MoveNext() |> ignore
////let v = enumerator.Current
////let t = v.GetType()
////printfn $"%A{t}"

//let range = sheet.["A1:B3"]
//printfn $"Height= {range.Height}, Width= {range.Width}"
//printfn $"RowHeight= {range.RowHeight}, ColumnWidth= {range.ColumnWidth}"

//let rs = range.Rows |> Seq.toArray
//let rs2 = range |> get :?> obj[,]
//printfn $"%A{rs2.[1, 1]}"




//let range = sheet.["A1:B3"]
//range |> rows |> iter (fun row -> 
//  row |> cells |> iter (fun cell ->
//    printf $"%A{cell |> integer} ")
//  printfn "")

//printfn ""

//range |> columns |> iter (fun col -> 
//  col |> cells |> iter (fun cell ->
//    printf $"%A{cell |> integer} ")
//  printfn "")
  


let range = sheet.["A1:B3"]
range |> rows |> iteri (fun i row ->
  if i % 2 = 0 then row |> bgcolor Color.Azure
  else row |> bgpattern Pattern.Gray16)
















//range.Rows
//|> Seq.iter (fun row -> 
//  row |> Seq.iter (fun cell -> printf $"%A{cell.Address} ")
//  printfn "")

detach app

//let range = sheet.["A1:B3"] |> gets
//let (x, y) = range |> len
//for i = 0 to x - 1 do
//  for j = 0 to y - 1 do
//    printf $"{range.[i, j]} "
//  printfn ""

//detach app
