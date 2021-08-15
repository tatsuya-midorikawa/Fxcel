namespace Fxcel.Core

open System.Reflection
open System.Buffers
open System.Runtime.CompilerServices
open System.Runtime.InteropServices
open System.Collections
open Microsoft.FSharp.NativeInterop

module rec ExcelInterop =
  [<Interface;ComImport;TypeLibType(4288s);Guid("000208DB-0000-0000-C000-000000000046");>]
  type IWorkbooks =
    inherit IEnumerable
    [<DispId(148);return: MarshalAs(UnmanagedType.Interface)>]
    abstract member Application : IApplication with get

  [<Interface;ComImport;TypeLibType(4160s);Guid("000208DA-0000-0000-C000-000000000046");>]
  type IWorkbook =
    [<DispId(148);return: MarshalAs(UnmanagedType.Interface)>]
    abstract member Application : IApplication with get

  [<Interface;ComImport;TypeLibType(4160s);Guid("000208D5-0000-0000-C000-000000000046");DefaultMember("_Default")>]
  type IApplication =
    [<DispId(148);return: MarshalAs(UnmanagedType.Interface)>]
    abstract member Application : IApplication with get
    [<DispId(308);return: MarshalAs(UnmanagedType.Interface)>]
    abstract member ActiveWorkbook : IWorkbook with get
    //[<DispId(572);return: MarshalAs(UnmanagedType.Interface)>]
    [<DispId(572);return: MarshalAs(UnmanagedType.Interface)>]
    abstract member Workbooks : obj with get

  [<AbstractClass;Guid("00024500-0000-0000-C000-000000000046");>]
  [<ComImport;ClassInterface(0s);DefaultMember("_Default");TypeLibType(2s)>]
  type ApplicationClass =
    [<DispId(148);return: MarshalAs(UnmanagedType.Interface)>]
    abstract member Application : ApplicationClass with get
    [<DispId(308);return: MarshalAs(UnmanagedType.Interface)>]
    abstract member ActiveWorkbook : WorkbookClass with get
    [<DispId(572);return: MarshalAs(UnmanagedType.Interface)>]
    abstract member Workbooks : IWorkbooks with get
      
  [<AbstractClass;Guid("00020819-0000-0000-C000-000000000046");>]
  [<ComImport;ClassInterface(0s);TypeLibType(2s)>]
  type WorkbookClass =
    [<DispId(148);return: MarshalAs(UnmanagedType.Interface)>]
    abstract member Application : ApplicationClass with get
