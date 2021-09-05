using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlMenus
    {
        private readonly XlApplication _app;
        public XlMenus(XlApplication app) => _app = app;
        public readonly XlMenu this[int index] => _app.collector.Mark(new XlMenu(_app.raw.ShortcutMenus[index]));
    }
}
