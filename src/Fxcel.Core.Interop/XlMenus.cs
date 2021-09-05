using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftMenu = Microsoft.Office.Interop.Excel.Menu;

    [SupportedOSPlatform("windows")]
    public readonly ref struct XlMenus
    {
        private readonly XlApplication _app;
        public XlMenus(XlApplication app) => _app = app;
        public XlMenu this[int index] => _app.ManageCom(new XlMenu(_app.raw.ShortcutMenus[index]));
    }
}
