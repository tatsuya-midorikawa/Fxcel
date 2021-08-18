// See https://aka.ms/new-console-template for more information
using System.Runtime.InteropServices;
using Fxcel.Core.Interop;

Console.WriteLine($"{IntPtr.Size}");
var size = Marshal.SizeOf<XlApplication>();
Console.WriteLine($"{size}");
