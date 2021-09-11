﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Versioning;
using Fxcel.Core.Interop.Common;
using System.Runtime.CompilerServices;

namespace Fxcel.Core.Interop
{
    using MicrosoftXmlNamespaces = Microsoft.Office.Interop.Excel.XmlNamespaces;

    [SupportedOSPlatform("windows")]
    public readonly struct XlXmlNamespaces : IComObject
    {
        internal readonly MicrosoftXmlNamespaces raw;
        private readonly ComCollector collector;
        private readonly bool disposed;

        internal XlXmlNamespaces(MicrosoftXmlNamespaces com)
        {
            raw = com;
            collector = new();
            disposed = false;
        }

        public readonly void Dispose()
        {
            if (!disposed)
            {
                // release managed objects
                collector.Collect();
                ForceRelease();

                // update status
                Unsafe.AsRef(disposed) = true;
            }
            GC.SuppressFinalize(this);
        }

        public readonly int Release() => ComHelper.Release(raw);
        public readonly void ForceRelease() => ComHelper.FinalRelease(raw);
    }
}