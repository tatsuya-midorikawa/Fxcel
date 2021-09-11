// Learn more about F# at http://docs.microsoft.com/dotnet/fsharp

open System
open BenchmarkDotNet.Attributes
open BenchmarkDotNet.Running

open Fxcel.Core.Interop

type Benchmark() =
  [<Benchmark>]
  member __.Struct() =
    use a = XlApplication.BlankWorkbook()
    ()

[<EntryPoint>]
let main argv =
  BenchmarkRunner.Run<Benchmark>() |> ignore
  
  0
