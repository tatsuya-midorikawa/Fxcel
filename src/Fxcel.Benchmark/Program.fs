// Learn more about F# at http://docs.microsoft.com/dotnet/fsharp

open System
open BenchmarkDotNet.Attributes
open BenchmarkDotNet.Running

type Benchmark() =
  [<Benchmark>]
  member __.M100() =
    for i in [| 1..100 |] do
      let a = 100
      ()

  [<Benchmark>]
  member __.M1000() =
    for i in [| 1..1000 |] do
      let a = 100
      ()


[<EntryPoint>]
let main argv =
  BenchmarkRunner.Run<Benchmark>() |> ignore
  
  0
