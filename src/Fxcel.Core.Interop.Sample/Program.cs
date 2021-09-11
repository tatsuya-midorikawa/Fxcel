// See https://aka.ms/new-console-template for more information

using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using Fxcel.Core.Interop;
using Fxcel.Core.Interop.Common;
using Excel = Microsoft.Office.Interop.Excel;



//using var app = XlApplication.BlankWorkbook();
//var a = app.Workbooks;
//var b = a.Parent;
//Console.WriteLine(a.GetType());


static class C
{
    struct Struct
    {
        public int Value = 100;
        public void M(int v) => Value = v;

        [return: NotNullIfNotNull("value")]
        public string? M(string? value) => value;
    }

    class Class
    {
        public int Value = 100;
        public void M(int v) => Value = v;
    }

    private static readonly Struct s = new Struct();
    private static readonly Class c = new Class();

    public static int Main(string[] args)
    {
        Console.WriteLine(s.Value);
        s.M(500);
        Console.WriteLine(s.Value);

        Console.WriteLine(c.Value);
        c.M(500);
        Console.WriteLine(c.Value);

        //Console.WriteLine("SizeOf({0}) = {1}", typeof(IntPtr), IntPtr.Size);
        //Console.WriteLine("SizeOf({0}) = {1}", typeof(XlApplication), Marshal.SizeOf(typeof(XlApplication)));
        //Console.WriteLine("SizeOf({0}) = {1}", typeof(XlWorkbooks), Marshal.SizeOf(typeof(XlWorkbooks)));
        //using var app = XlApplication.BlankWorkbook();
        //var a = app.Workbooks;
        //var b = a.Parent;
        //Console.WriteLine(a.GetType());

        return 0;
    }
}




//public readonly ref struct S
//{
//    public readonly int Number;
//    public readonly string Value;
//}

//public sealed class C
//{
//    public ref readonly S M(in S lhs, in S rhs)
//    {
//        ref readonly var l = ref lhs;
//        ref readonly var r = ref rhs;
//        return ref (l.Number < r.Number ? ref r : ref l);
//    }
//}

//readonly record struct Foo (int Bar);

//static class C
//{
//    public static int Main(string[] args)
//    {
//        Console.WriteLine("SizeOf({0}) = {1}", typeof(IntPtr), IntPtr.Size);
//        Console.WriteLine("SizeOf({0}) = {1}", typeof(XlApplication), Marshal.SizeOf(typeof(XlApplication)));
//        Console.WriteLine("SizeOf({0}) = {1}", typeof(XlWorkbooks), Marshal.SizeOf(typeof(XlWorkbooks)));
//        using var app = XlApplication.BlankWorkbook();
//        var a = app.Workbooks;
//        var b = a.Parent;
//        Console.WriteLine(a.GetType());
//        return 0;
//    }
//}


//var books = app.Workbooks;
//var book = books[1];
//var colors = book.Colors;
//Console.WriteLine(colors.GetType());
//ComHelper.FinalRelease(colors);
//book.FinalRelease();
//books.FinalRelease();
//app.Quit();
//app.FinalRelease();
//Console.WriteLine($"{IntPtr.Size}");
//var size = Marshal.SizeOf<XlApplication>();
//Console.WriteLine($"{size}");

//// 必要な変数は try の外で宣言する
//Excel.Application xlApplication = null;

//// COM オブジェクトの解放を保証するために try ～ finally を使用する
//try
//{
//    xlApplication = new Excel.Application();

//    // 警告メッセージなどを表示しないようにする
//    xlApplication.DisplayAlerts = false;

//    Excel.Workbooks xlBooks = xlApplication.Workbooks;

//    try
//    {
//        Excel.Workbook xlBook = xlBooks.Add(string.Empty);

//        try
//        {
//            Excel.Sheets xlSheets = xlBook.Worksheets;

//            try
//            {
//                Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets[1];

//                try
//                {
//                    Excel.Range xlCells = xlSheet.Cells;

//                    try
//                    {
//                        Excel.Range xlRange = (Excel.Range)xlCells[6, 4];

//                        try
//                        {
//                            // Microsoft Excel を表示する
//                            xlApplication.Visible = true;

//                            // 1000 ミリ秒 (1秒) 待機する
//                            System.Threading.Thread.Sleep(1000);

//                            // Row=6, Column=4 の位置に文字をセットする
//                            xlRange.Value2 = "あと 1 秒で終了します";

//                            // 1000 ミリ秒 (1秒) 待機する
//                            System.Threading.Thread.Sleep(1000);
//                        }
//                        finally
//                        {
//                            //Marshal.AddRef(Marshal.GetComInterfaceForObject(xlRange, typeof(Excel.Range)));
//                            //Console.WriteLine($"xlRange is COM object= {Marshal.IsComObject(xlRange)}, ref count= {Marshal.ReleaseComObject(xlRange)}");
//                            if (xlRange != null)
//                            {
//                                //xlRange = null;
//                                Marshal.FinalReleaseComObject(xlRange);
//                            }
//                        }
//                    }
//                    finally
//                    {
//                        //Marshal.AddRef(Marshal.GetComInterfaceForObject(xlCells, typeof(Excel.Range)));
//                        //Console.WriteLine($"xlCells is COM object= {Marshal.IsComObject(xlCells)}, ref count= {Marshal.ReleaseComObject(xlCells)}");
//                        if (xlCells != null)
//                        {
//                            //xlCells = null;
//                            Marshal.FinalReleaseComObject(xlCells);
//                        }
//                    }
//                }
//                finally
//                {
//                    //Marshal.AddRef(Marshal.GetComInterfaceForObject(xlSheet, typeof(Excel.Worksheet)));
//                    //Console.WriteLine($"xlSheet is COM object= {Marshal.IsComObject(xlSheet)}, ref count= {Marshal.ReleaseComObject(xlSheet)}");
//                    if (xlSheet != null)
//                    {
//                        //xlSheet = null;
//                        Marshal.FinalReleaseComObject(xlSheet);
//                    }
//                }
//            }
//            finally
//            {
//                //Marshal.AddRef(Marshal.GetComInterfaceForObject(xlSheets, typeof(Excel.Sheets)));
//                //Console.WriteLine($"xlSheets is COM object= {Marshal.IsComObject(xlSheets)}, ref count= {Marshal.ReleaseComObject(xlSheets)}");
//                if (xlSheets != null)
//                {
//                    //xlSheets = null;
//                    Marshal.FinalReleaseComObject(xlSheets);
//                }
//            }
//        }
//        finally
//        {
//            //Console.WriteLine($"xlBook is COM object= {Marshal.IsComObject(xlBook)}");
//            if (xlBook != null)
//            {
//                try
//                {
//                    xlBook.Close();
//                }
//                finally
//                {
//                    //xlBook = null;
//                    Marshal.FinalReleaseComObject(xlBook);
//                }
//            }
//        }
//    }
//    finally
//    {
//        //Console.WriteLine($"xlBooks is COM object= {Marshal.IsComObject(xlBooks)}");
//        if (xlBooks != null)
//        {
//            //xlBooks = null;
//            Marshal.FinalReleaseComObject(xlBooks);
//        }
//    }
//}
//finally
//{
//    //Console.WriteLine($"xlApplication is COM object= {Marshal.IsComObject(xlApplication)}");
//    if (xlApplication != null)
//    {
//        try
//        {
//            xlApplication.Quit();
//        }
//        finally
//        {
//            //xlApplication = null;
//            Marshal.FinalReleaseComObject(xlApplication);
//        }
//    }
//}
