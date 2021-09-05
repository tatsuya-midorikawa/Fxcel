﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftMenu = Microsoft.Office.Interop.Excel.Menu;

    [SupportedOSPlatform("windows")]
    public class XlMenu : XlComObject
    {
        public XlMenu(MicrosoftMenu menu) : base(menu) { }
        private MicrosoftMenu raw => (MicrosoftMenu)_raw;
    }
}
