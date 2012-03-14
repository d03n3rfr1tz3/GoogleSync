
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Outlook;

namespace DirkSarodnick.GoogleSync.Core.Data.Recurrence
{
    /// <summary>
    /// Defines the RecurrenceSerializer class.
    /// </summary>
    public static class RecurrenceSerializer
    {
        private static readonly string[] DateTimeFormat = new[] { @"yyyyMMddTHHmmss", @"yyyyMMddTHHmmssZ", @"yyyyMMdd" };
        private static readonly Regex StartRegex = new Regex(@"DTSTART(?<Zone>\S*):(?<Date>\w+)");
        private static readonly Regex EndRegex = new Regex(@"DTEND(?<Zone>\S*):(?<Date>\w+)");
        private static readonly Regex RuleRegex = new Regex(@"RRULE:(?<Rule>\S*)");
        private static readonly Regex FrequenceRegex = new Regex(@"FREQ=(?<Value>\w+)");
        private static readonly Regex ByDayRegex = new Regex(@"BYMONTHDAY=(?<Value>\d+)");
        private static readonly Regex ByWeekDayRegex = new Regex(@"BYDAY=(?<Value>\w+)");
        private static readonly Regex ByMonthRegex = new Regex(@"BYMONTH=(?<Value>\w+)");
        private static readonly Regex IntervalRegex = new Regex(@"INTERVAL=(?<Value>\d+)");
        private static readonly Regex CountRegex = new Regex(@"COUNT=(?<Value>\d+)");
        private static readonly Regex UntilRegex = new Regex(@"UNTIL=(?<Date>\w+)");

        /// <summary>
        /// Deserializes the specified recurrence.
        /// </summary>
        /// <param name="recurrence">The recurrence.</param>
        /// <returns>The Recurrence Data.</returns>
        public static RecurrenceData Deserialize(string recurrence)
        {
            var result = new RecurrenceData();

            if (StartRegex.IsMatch(recurrence))
            {
                var dateString = StartRegex.Match(recurrence).Groups["Date"].Value;

                bool allDayEvent;
                DateTime dateTime;
                if(dateString.ToDate(out dateTime, out allDayEvent))
                {
                    result.StartTime = dateTime;
                    result.AllDayEvent = allDayEvent;
                }
            }

            if (EndRegex.IsMatch(recurrence))
            {
                var dateString = EndRegex.Match(recurrence).Groups["Date"].Value;

                bool allDayEvent;
                DateTime dateTime;
                if (dateString.ToDate(out dateTime, out allDayEvent))
                {
                    result.EndTime = dateTime;
                    result.AllDayEvent = allDayEvent;
                }
            }

            if (RuleRegex.IsMatch(recurrence))
            {
                var ruleString = RuleRegex.Match(recurrence).Groups["Rule"].Value;
                var rules = ruleString.Split(';');

                foreach (var rule in rules)
                {
                    if (FrequenceRegex.IsMatch(rule))
                    {
                        var value = FrequenceRegex.Match(rule).Groups["Value"].Value;
                        result.RecurrenceType = GetRecurrenceType(value);
                    }

                    if (ByDayRegex.IsMatch(rule))
                    {
                        var value = ByDayRegex.Match(rule).Groups["Value"].Value;
                        int integer;
                        if (int.TryParse(value, out integer))
                        {
                            result.DayOfMonth = integer;
                        }
                    }

                    if (ByWeekDayRegex.IsMatch(rule))
                    {
                        var value = ByWeekDayRegex.Match(rule).Groups["Value"].Value;
                        result.DayOfWeek = GetDayOfWeek(value);
                    }

                    if (ByMonthRegex.IsMatch(rule))
                    {
                        var value = ByMonthRegex.Match(rule).Groups["Value"].Value;
                        int integer;
                        if (int.TryParse(value, out integer))
                        {
                            result.MonthOfYear = integer;
                        }
                    }

                    if (IntervalRegex.IsMatch(rule))
                    {
                        var value = IntervalRegex.Match(rule).Groups["Value"].Value;
                        int integer;
                        if (int.TryParse(value, out integer))
                        {
                            result.Interval = integer;
                        }
                    }

                    if (UntilRegex.IsMatch(rule))
                    {
                        var dateString = UntilRegex.Match(rule).Groups["Date"].Value;

                        bool allDayEvent;
                        DateTime dateTime;
                        if (dateString.ToDate(out dateTime, out allDayEvent))
                        {
                            result.EndPattern = dateTime;
                        }
                    }
                    else if (CountRegex.IsMatch(rule))
                    {
                        var value = CountRegex.Match(rule).Groups["Value"].Value;
                        int integer;
                        if (int.TryParse(value, out integer))
                        {
                            result.Count = integer;
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Serializes the specified get recurrence pattern.
        /// </summary>
        /// <param name="recurrencePattern">The recurrence pattern.</param>
        /// <param name="appointmentStart">The appointment start.</param>
        /// <param name="allday">if set to <c>true</c> [allday].</param>
        /// <returns>The serialized Recurrence.</returns>
        public static string Serialize(RecurrencePattern recurrencePattern, DateTime appointmentStart, bool allday)
        {
            var startTime = appointmentStart.Date.Add(recurrencePattern.StartTime - recurrencePattern.StartTime.Date);
            var endTime = appointmentStart.Date.Add(recurrencePattern.EndTime - recurrencePattern.EndTime.Date);
            var values = allday
            ? new List<string>
            {
                string.Format("DTSTART;VALUE=DATE:{0}", startTime.ToString("yyyyMMdd", CultureInfo.InvariantCulture)),
                string.Format("DTEND;VALUE=DATE:{0}", endTime.ToString("yyyyMMdd", CultureInfo.InvariantCulture))
            }
            : new List<string>
            {
                string.Format("DTSTART:{0}", startTime.ToUniversalTime().ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture)),
                string.Format("DTEND:{0}", endTime.ToUniversalTime().ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture))
            };
            
            var ruleValues = new List<string>
            {
                string.Format("FREQ={0}", recurrencePattern.RecurrenceType.GetRecurrenceString())
            };

            if (!recurrencePattern.NoEndDate)
            {
                ruleValues.Add(string.Format("UNTIL={0}", recurrencePattern.PatternEndDate.ToUniversalTime().ToString("yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture)));
            }

            if ((recurrencePattern.RecurrenceType == OlRecurrenceType.olRecursYearly && recurrencePattern.Interval / 12 > 1) ||
                (recurrencePattern.RecurrenceType != OlRecurrenceType.olRecursYearly && recurrencePattern.Interval > 1))
            {
                ruleValues.Add(string.Format("INTERVAL={0}", recurrencePattern.Interval));
            }

            if (recurrencePattern.NoEndDate && recurrencePattern.Occurrences > 0)
            {
                ruleValues.Add(string.Format("COUNT={0}", recurrencePattern.Occurrences));
            }
            
            switch (recurrencePattern.RecurrenceType)
            {
                case OlRecurrenceType.olRecursWeekly:
                    ruleValues.Add(string.Format("BYDAY={0}", recurrencePattern.DayOfWeekMask.GetDayOfWeek()));
                    break;
                case OlRecurrenceType.olRecursDaily:
                case OlRecurrenceType.olRecursMonthly:
                case OlRecurrenceType.olRecursYearly:
                    if (recurrencePattern.MonthOfYear > 0)
                        ruleValues.Add(string.Format("BYMONTH={0}", recurrencePattern.MonthOfYear));
                    if (recurrencePattern.DayOfMonth > 0)
                        ruleValues.Add(string.Format("BYMONTHDAY={0}", recurrencePattern.DayOfMonth));
                    break;
            }

            values.Add(string.Format("RRULE:{0}", string.Join(";", ruleValues)));
            return string.Join("\r\n", values);
        }

        /// <summary>
        /// Gets the type of the recurrence.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>The RecurrenceType.</returns>
        public static OlRecurrenceType GetRecurrenceType(string value)
        {
            switch (value)
            {
                case "DAILY":
                    return OlRecurrenceType.olRecursDaily;
                case "WEEKLY":
                    return OlRecurrenceType.olRecursWeekly;
                case "MONTHLY":
                    return OlRecurrenceType.olRecursMonthly;
                case "YEARLY":
                    return OlRecurrenceType.olRecursYearly;
            }

            return OlRecurrenceType.olRecursYearly;
        }

        /// <summary>
        /// Gets the recurrence string.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>The Recurrence string.</returns>
        public static string GetRecurrenceString(this OlRecurrenceType value)
        {
            switch (value)
            {
                case OlRecurrenceType.olRecursDaily:
                    return "DAILY";
                case OlRecurrenceType.olRecursWeekly:
                    return "WEEKLY";
                case OlRecurrenceType.olRecursMonthly:
                    return "MONTHLY";
                case OlRecurrenceType.olRecursYearly:
                    return "YEARLY";
            }

            return "YEARLY";
        }

        /// <summary>
        /// Gets the day of week.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>The DayOfWeek.</returns>
        public static OlDaysOfWeek GetDayOfWeek(string value)
        {
            switch (value)
            {
                case "MO":
                    return OlDaysOfWeek.olMonday;
                case "TU":
                    return OlDaysOfWeek.olTuesday;
                case "WE":
                    return OlDaysOfWeek.olWednesday;
                case "TH":
                    return OlDaysOfWeek.olThursday;
                case "FR":
                    return OlDaysOfWeek.olFriday;
                case "SA":
                    return OlDaysOfWeek.olSaturday;
                case "SU":
                    return OlDaysOfWeek.olSunday;
            }

            return OlDaysOfWeek.olMonday;
        }

        /// <summary>
        /// Gets the day of week.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>The Day of Week.</returns>
        public static string GetDayOfWeek(this OlDaysOfWeek value)
        {
            switch (value)
            {
                case OlDaysOfWeek.olMonday:
                    return "MO";
                case OlDaysOfWeek.olTuesday:
                    return "TU";
                case OlDaysOfWeek.olWednesday:
                    return "WE";
                case OlDaysOfWeek.olThursday:
                    return "TH";
                case OlDaysOfWeek.olFriday:
                    return "FR";
                case OlDaysOfWeek.olSaturday:
                    return "SA";
                case OlDaysOfWeek.olSunday:
                    return "SU";
            }

            return "MO";
        }

        /// <summary>
        /// Parses a DateTime of specified string.
        /// </summary>
        /// <param name="dateString">The date string.</param>
        /// <param name="resultDateTime">The result date time.</param>
        /// <param name="allDayEvent">if set to <c>true</c> [all day event].</param>
        /// <returns>True if Changed.</returns>
        public static bool ToDate(this string dateString, out DateTime resultDateTime, out bool allDayEvent)
        {
            bool result = false;
            
            DateTime dateTime;
            resultDateTime = new DateTime();
            allDayEvent = false;

            if (DateTime.TryParseExact(dateString, "yyyyMMddTHHmmssZ", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out dateTime))
            {
                resultDateTime = dateTime;
                result = true;
            }
            else if (DateTime.TryParseExact(dateString, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime))
            {
                resultDateTime = dateTime;
                allDayEvent = true;
                result = true;
            }
            else if (DateTime.TryParseExact(dateString, "yyyyMMddTHHmmss", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime))
            {
                resultDateTime = dateTime;
                result = true;
            }

            return result;
        }
    }
}
