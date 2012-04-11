
namespace DirkSarodnick.GoogleSync.Core.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Defines the Linq Extensions.
    /// </summary>
    public static class LinqExtensions
    {
        /// <summary>
        /// Firsts the or instance.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="enumerable">The enumerable.</param>
        /// <returns>The First Element or a new Instance of same type.</returns>
        public static T FirstOrInstance<T>(this IEnumerable<T> enumerable)
            where T : class, new()
        {
            return enumerable.FirstOrDefault() ?? new T();
        }

        /// <summary>
        /// Firsts the or instance.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="enumerable">The enumerable.</param>
        /// <param name="predicate">The predicate.</param>
        /// <returns>The First Element or a new Instance of same type.</returns>
        public static T FirstOrInstance<T>(this IEnumerable<T> enumerable, Func<T, bool> predicate)
            where T : class, new()
        {
            return enumerable.FirstOrDefault(predicate) ?? new T();
        }
    }
}
