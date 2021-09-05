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
    public sealed class XlLanguageSettings : XlComObject
    {
        internal XlLanguageSettings(MicrosoftLanguageSettings com) => raw = com;
        internal MicrosoftLanguageSettings raw;

        public override int Release() => ComHelper.Release(raw);
        public override void ForceRelease() => ComHelper.FinalRelease(raw);
        protected override void DidDispose()
        {
            raw = default!;
            base.DidDispose();
        }
    }
}
