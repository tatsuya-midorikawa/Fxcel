using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Fxcel.Core.Interop.Common
{
    public readonly struct ComCollector
    {
        private readonly List<IComObject> _collection = new();

        public readonly T Mark<T>(T target) where T : IComObject
        {
            _collection.Add(target);
            return target;
        }

        public readonly void Collect()
        {
            foreach (var item in _collection)
                try { item?.Dispose(); } catch { /* ignore */ }
        }

        public readonly void Collect(Action didCollect)
        {
            Collect();
            didCollect();
        }
    }
}
