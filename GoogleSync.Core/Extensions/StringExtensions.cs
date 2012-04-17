
namespace DirkSarodnick.GoogleSync.Core.Extensions
{
    using System;
    using System.Linq.Expressions;
    using System.Reflection;
    using System.Text.RegularExpressions;

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

            var propertyString = Equals(propertyValue, default(TValue)) ? string.Empty : propertyValue.ToString();
            var valueString = Equals(value, default(TValue)) ? string.Empty : value.ToString();

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
        /// Formats the value simpler, cause there are sometimes string.Empty strings compared against googles preferable nulls.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        public static string FormatSimple(this string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }

            return value;
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

            return string.Concat(" ", value);
        }

        /// <summary>
        /// Formats the phone as clean numbers.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        public static string FormatPhoneClean(this string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            return new Regex(@"\D", RegexOptions.CultureInvariant).Replace(value, string.Empty);
        }
    }
}