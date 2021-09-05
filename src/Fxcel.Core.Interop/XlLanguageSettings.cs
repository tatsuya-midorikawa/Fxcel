using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop
{
    using MicrosoftLanguageSettings = Microsoft.Office.Core.LanguageSettings;

    [SupportedOSPlatform("windows")]
    public class XlLanguageSettings : XlComObject
    {
        public XlLanguageSettings(MicrosoftLanguageSettings settings) : base(settings) { }
        private MicrosoftLanguageSettings raw => (MicrosoftLanguageSettings)_raw;
    }
}
