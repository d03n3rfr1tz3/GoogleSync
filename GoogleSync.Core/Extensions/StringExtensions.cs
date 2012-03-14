
namespace DirkSarodnick.GoogleSync.Core.Extensions
{
    using System;
    using System.Linq.Expressions;
    using System.Reflection;

    /// <summary>
    /// Defines the String Extensions.
    /// </summary>
    public static class StringExtensions
    {
        /// <summary>
        /// Applies the property.
        /// </summary>
        /// <typeparam name="TObject">The type of the object.</typeparam>
        /// <typeparam name="TValue">The type of the value.</typeparam>
        /// <param name="baseObject">The base object.</param>
        /// <param name="expression">The expression.</param>
        /// <param name="value">The value.</param>
        /// <returns>True if Changed.</returns>
        public static bool ApplyProperty<TObject, TValue>(this TObject baseObject, Expression<Func<TObject, TValue>> expression, TValue value)
            where TObject : class
        {
            var propertyValue = expression.Compile()(baseObject);
            var equals = Equals(propertyValue, value);

            var propertyString = propertyValue == null ? string.Empty : propertyValue.ToString();
            var valueString = value == null ? string.Empty : value.ToString();

            if (string.IsNullOrWhiteSpace(propertyString) && string.IsNullOrWhiteSpace(valueString))
            {
                equals = true;
            }

            if (!equals)
            {
                var property = (PropertyInfo)((MemberExpression)expression.Body).Member;
                property.SetValue(baseObject, value, null);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Formats the phone.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>The formatted Phone.</returns>
        public static string FormatPhone(this string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            return value.Replace(@"+", @" +");
        }
    }
}
