using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using MicrosoftLanguageSettings = Microsoft.Office.Core.LanguageSettings;

namespace Fxcel.Core.Interop
{
    [SupportedOSPlatform("windows")]
    public readonly ref struct XlLanguageSettings
    {
        internal readonly MicrosoftLanguageSettings raw;
        public XlLanguageSettings(MicrosoftLanguageSettings settings) => raw = settings;

        public int Release() => ComHelper.Release(raw);
    }
}
