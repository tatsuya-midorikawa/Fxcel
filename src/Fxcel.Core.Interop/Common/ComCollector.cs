using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop.Common
{
    public readonly struct ComCollector
    {
        private readonly List<IComObject> _collection = new();

        public readonly ref readonly T Mark<T>(in T target) where T : IComObject
        {
            _collection.Add(target);
            return ref target;
        }

        public readonly void Collect()
        {
            foreach (var item in _collection)
                try { item?.Dispose(); } catch(Exception e) { Debug.WriteLine(e.Message); }
        }

        public readonly void Collect(Action didCollect)
        {
            Collect();
            didCollect();
        }
    }
}
