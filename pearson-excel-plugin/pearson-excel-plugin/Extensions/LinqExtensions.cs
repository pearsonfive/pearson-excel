using System;
using System.Collections.Generic;
using System.Linq;

namespace Pearson.Excel.Plugin.Extensions
{
    public static class LinqExtensions
    {
        public static void ForEach<T>(this IEnumerable<T> enumerable, Action<T> action)
        {
            foreach (var item in enumerable ?? Enumerable.Empty<T>())
            {
                action(item);
            }
        }
    }
}