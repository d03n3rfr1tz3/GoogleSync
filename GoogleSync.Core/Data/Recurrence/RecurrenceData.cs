
namespace DirkSarodnick.GoogleSync.Core.Data.Recurrence
{
    using System;
    using Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Defines the RecurrenceData class.
    /// </summary>
    public class RecurrenceData
    {
        /// <summary>
        /// Gets or sets the type of the recurrence.
        /// </summary>
        /// <value>The type of the recurrence.</value>
        public OlRecurrenceType RecurrenceType { get; set; }

        /// <summary>
        /// Gets or sets the start time.
        /// </summary>
        /// <value>The start time.</value>
        public DateTime StartTime { get; set; }

        /// <summary>
        /// Gets or sets the end time.
        /// </summary>
        /// <value>The end time.</value>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// Gets or sets the start pattern.
        /// </summary>
        /// <value>The start pattern.</value>
        public DateTime? StartPattern { get; set; }

        /// <summary>
        /// Gets or sets the end pattern.
        /// </summary>
        /// <value>The end pattern.</value>
        public DateTime? EndPattern { get; set; }

        /// <summary>
        /// Gets or sets the day of month.
        /// </summary>
        /// <value>The day of month.</value>
        public int DayOfMonth { get; set; }

        /// <summary>
        /// Gets or sets the day of week.
        /// </summary>
        /// <value>The day of week.</value>
        public OlDaysOfWeek DayOfWeek { get; set; }

        /// <summary>
        /// Gets or sets the month of year.
        /// </summary>
        /// <value>The month of year.</value>
        public int MonthOfYear { get; set; }

        /// <summary>
        /// Gets or sets the interval.
        /// </summary>
        /// <value>The interval.</value>
        public int? Interval { get; set; }

        /// <summary>
        /// Gets or sets the count.
        /// </summary>
        /// <value>The count.</value>
        public int? Count { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether [all day event].
        /// </summary>
        /// <value><c>true</c> if [all day event]; otherwise, <c>false</c>.</value>
        public bool AllDayEvent { get; set; }
    }
}
