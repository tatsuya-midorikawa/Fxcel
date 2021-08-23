﻿namespace Fxcel.Core.Interop
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    public enum XlClipboardFormat
    {
        BIFF12 = 0x3F,
        BIFF = 8,
        BIFF2 = 18,
        BIFF3 = 20,
        BIFF4 = 30,
        Binary = 0xF,
        Bitmap = 9,
        CGM = 13,
        CSV = 5,
        DIF = 4,
        DspText = 12,
        EmbeddedObject = 21,
        EmbedSource = 22,
        Link = 11,
        LinkSource = 23,
        LinkSourceDesc = 0x20,
        Movie = 24,
        Native = 14,
        ObjectDesc = 0x1F,
        ObjectLink = 19,
        OwnerLink = 17,
        PICT = 2,
        PrintPICT = 3,
        RTF = 7,
        ScreenPICT = 29,
        StandardFont = 28,
        StandardScale = 27,
        SYLK = 6,
        Table = 0x10,
        Text = 0,
        ToolFace = 25,
        ToolFacePICT = 26,
        VALU = 1,
        WK1 = 10
    }
}
