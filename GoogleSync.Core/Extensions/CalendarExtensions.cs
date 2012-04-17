
namespace DirkSarodnick.GoogleSync.Core.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using Data;
    using Data.Recurrence;
    using Google.GData.Calendar;
    using Google.GData.Extensions;
    using Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Defines the Calendar Extensions.
    /// </summary>
    public static class CalendarExtensions
    {
        /// <summary>
        /// Determines wether the calendar items are mergable.
        /// </summary>
        /// <param name="outlookCalendarItem">The outlook calendar item.</param>
        /// <param name="googleCalendarItem">The google calendar item.</param>
        /// <returns>True or False.</returns>
        public static bool Mergeable(AppointmentItem outlookCalendarItem, EventEntry googleCalendarItem)
        {
            return outlookCalendarItem.UserProperties.GetProperty("GoogleId") == googleCalendarItem.EventId ||
                   (outlookCalendarItem.Subject == googleCalendarItem.Title.Text && googleCalendarItem.Locations.Any(l => l.ValueString == outlookCalendarItem.Location)) ||
                   googleCalendarItem.Times.Any(g => g.StartTime == outlookCalendarItem.Start && g.EndTime == outlookCalendarItem.End);
        }

        #region Google > Outlook

        /// <summary>
        /// Determines wether the calendar items are mergable.
        /// </summary>
        /// <param name="calendarItems">The calendar items.</param>
        /// <param name="calendarItem">The calendar item.</param>
        /// <returns>The mergeable Calendar Items.</returns>
        public static IEnumerable<AppointmentItem> Mergeable(this IEnumerable<AppointmentItem> calendarItems, EventEntry calendarItem)
        {
            return calendarItems.Where(c => Mergeable(c, calendarItem));
        }

        /// <summary>
        /// Merges the with.
        /// </summary>
        /// <param name="outlookCalendarItem">The outlook calendar item.</param>
        /// <param name="googleCalendarItem">The google calendar item.</param>
        /// <returns>True if Changed.</returns>
        public static bool MergeWith(this AppointmentItem outlookCalendarItem, EventEntry googleCalendarItem)
        {
            var result = false;

            result |= outlookCalendarItem.ApplyProperty(c => c.Subject, googleCalendarItem.Title.Text);
            result |= outlookCalendarItem.ApplyProperty(c => c.Body, googleCalendarItem.Content.Content);
            result |= outlookCalendarItem.ApplyProperty(c => c.Location, googleCalendarItem.Locations.FirstOrInstance(l => l.Rel == Where.RelType.EVENT).ValueString);
            result |= outlookCalendarItem.ApplyProperty(c => c.BusyStatus, googleCalendarItem.EventTransparency.GetStatus());
            result |= outlookCalendarItem.ApplyProperty(c => c.Sensitivity, googleCalendarItem.EventVisibility.GetStatus());
            result |= outlookCalendarItem.MergeRecipients(googleCalendarItem.Participants);


            if (googleCalendarItem.Times.Any())
            {
                var time = googleCalendarItem.Times.First();
                result |= outlookCalendarItem.ApplyProperty(c => c.AllDayEvent, time.AllDay);
                result |= outlookCalendarItem.ApplyProperty(c => c.Start, time.StartTime);
                result |= outlookCalendarItem.ApplyProperty(c => c.End, time.EndTime);

                if (time.Reminders.Any(r => r.Method == Google.GData.Extensions.Reminder.ReminderMethod.alert))
                {
                    var reminder = time.Reminders.First(r => r.Method == Google.GData.Extensions.Reminder.ReminderMethod.alert);
                    result |= outlookCalendarItem.ApplyProperty(c => c.ReminderSet, true);
                    result |= outlookCalendarItem.ApplyProperty(c => c.ReminderOverrideDefault, true);
                    result |= outlookCalendarItem.ApplyProperty(c => c.ReminderMinutesBeforeStart, reminder.GetMinutes());
                }
                else
                {
                    result |= outlookCalendarItem.ApplyProperty(c => c.ReminderSet, false);
                    result |= outlookCalendarItem.ApplyProperty(c => c.ReminderOverrideDefault, false);
                }
            }

            result |= outlookCalendarItem.MergeRecurrence(googleCalendarItem.Recurrence);

            return result;
        }

        /// <summary>
        /// Merges the specified recipients.
        /// </summary>
        /// <param name="outlookCalendarItem">The outlook calendar item.</param>
        /// <param name="googleParticipants">The google participants.</param>
        /// <returns>True if Changed.</returns>
        public static bool MergeRecipients(this AppointmentItem outlookCalendarItem, IEnumerable<Who> googleParticipants)
        {
            var result = false;
            var recipients = outlookCalendarItem.Recipients.Cast<Recipient>().Where(r => r.Address != null).ToList();
            var participants = googleParticipants.Where(p => p.Rel != Who.RelType.EVENT_ORGANIZER).ToList();

            foreach (var participant in participants)
            {
                if (!recipients.Any(r => r.Address == participant.Email || r.Name == participant.Email))
                {
                    var recipient = outlookCalendarItem.Recipients.Add(participant.Email);
                    recipient.Type = (int)participant.Attendee_Type.GetRecipientType();
                    outlookCalendarItem.MeetingStatus = OlMeetingStatus.olMeeting;
                    result = true;
                }
            }

            foreach (Recipient recipient in recipients)
            {
                if (!participants.Any(p => p.Email == (recipient.Address ?? recipient.Name)))
                {
                    outlookCalendarItem.Recipients.Remove(recipient.Index);
                    result = true;
                }
            }

            return result;
        }

        /// <summary>
        /// Merges the recurrence.
        /// </summary>
        /// <param name="outlookCalendarItem">The outlook calendar item.</param>
        /// <param name="recurrence">The recurrence.</param>
        /// <returns>True if Change.</returns>
        public static bool MergeRecurrence(this AppointmentItem outlookCalendarItem, Recurrence recurrence)
        {
            if (recurrence != null && !string.IsNullOrWhiteSpace(recurrence.Value))
            {
                var result = false;
                var outlookRecurrencePattern = outlookCalendarItem.GetRecurrencePattern();
                var googleRecurrencePattern = RecurrenceSerializer.Deserialize(recurrence.Value);

                try { result |= outlookCalendarItem.ApplyProperty(r => r.AllDayEvent, googleRecurrencePattern.AllDayEvent); } catch (TargetInvocationException) { }
                result |= outlookRecurrencePattern.ApplyProperty(r => r.RecurrenceType, googleRecurrencePattern.RecurrenceType);
                result |= outlookRecurrencePattern.ApplyProperty(r => r.PatternStartDate, googleRecurrencePattern.StartPattern.HasValue ? googleRecurrencePattern.StartPattern.Value.Date : DateTime.Today);
                if (googleRecurrencePattern.EndPattern.HasValue)
                    result |= outlookRecurrencePattern.ApplyProperty(r => r.PatternEndDate, googleRecurrencePattern.EndPattern.Value.Date);
                else
                    result |= outlookRecurrencePattern.ApplyProperty(r => r.NoEndDate, true);

                result |= outlookRecurrencePattern.ApplyProperty(r => r.DayOfMonth, googleRecurrencePattern.DayOfMonth);
                result |= outlookRecurrencePattern.ApplyProperty(r => r.DayOfWeekMask, googleRecurrencePattern.DayOfWeek);
                result |= outlookRecurrencePattern.ApplyProperty(r => r.MonthOfYear, googleRecurrencePattern.MonthOfYear);

                if (outlookRecurrencePattern.StartTime.TimeOfDay.Ticks != googleRecurrencePattern.StartTime.TimeOfDay.Ticks)
                    result |= outlookRecurrencePattern.ApplyProperty(r => r.StartTime, googleRecurrencePattern.StartTime);

                if (outlookRecurrencePattern.EndTime.TimeOfDay.Ticks != googleRecurrencePattern.EndTime.TimeOfDay.Ticks)
                    result |= outlookRecurrencePattern.ApplyProperty(r => r.EndTime, googleRecurrencePattern.EndTime);

                if (googleRecurrencePattern.Count.HasValue)
                    result |= outlookRecurrencePattern.ApplyProperty(r => r.Occurrences, googleRecurrencePattern.Count.Value);

                if (googleRecurrencePattern.Interval.HasValue)
                    result |= outlookRecurrencePattern.ApplyProperty(r => r.Interval, googleRecurrencePattern.Interval.Value);

                return result;
            }

            if (outlookCalendarItem.RecurrenceState == OlRecurrenceState.olApptNotRecurring)
            {
                outlookCalendarItem.ClearRecurrencePattern();
                return true;
            }

            return false;
        }

        /// <summary>
        /// Gets the minutes.
        /// </summary>
        /// <param name="reminder">The reminder.</param>
        /// <returns>The Minutes.</returns>
        public static int GetMinutes(this Google.GData.Extensions.Reminder reminder)
        {
            int result = reminder.Minutes;
            result += reminder.Hours * 60;
            result += reminder.Days * 24 * 60;
            return result;
        }

        /// <summary>
        /// Gets the status.
        /// </summary>
        /// <param name="transparency">The transparency.</param>
        /// <returns>The Busy status.</returns>
        public static OlBusyStatus GetStatus(this EventEntry.Transparency transparency)
        {
            switch (transparency.Value)
            {
                case EventEntry.Transparency.TRANSPARENT_VALUE:
                    return OlBusyStatus.olFree;
                case EventEntry.Transparency.OPAQUE_VALUE:
                    return OlBusyStatus.olBusy;
            }

            return OlBusyStatus.olBusy;
        }

        /// <summary>
        /// Gets the status.
        /// </summary>
        /// <param name="visibility">The visibility.</param>
        /// <returns>The Sensitity.</returns>
        public static OlSensitivity GetStatus(this EventEntry.Visibility visibility)
        {
            switch (visibility.Value)
            {
                case EventEntry.Visibility.DEFAULT_VALUE:
                    return OlSensitivity.olNormal;
                case EventEntry.Visibility.PRIVATE_VALUE:
                    return OlSensitivity.olPrivate;
                case EventEntry.Visibility.CONFIDENTIAL_VALUE:
                    return OlSensitivity.olConfidential;
            }

            return OlSensitivity.olNormal;
        }

        /// <summary>
        /// Gets the status.
        /// </summary>
        /// <param name="status">The status.</param>
        /// <returns>The Response status.</returns>
        public static OlResponseStatus GetStatus(this EventEntry.EventStatus status)
        {
            switch (status.Value)
            {
                case EventEntry.EventStatus.CONFIRMED_VALUE:
                    return OlResponseStatus.olResponseAccepted;
                case EventEntry.EventStatus.CANCELED_VALUE:
                    return OlResponseStatus.olResponseDeclined;
                case EventEntry.EventStatus.TENTATIVE_VALUE:
                    return OlResponseStatus.olResponseTentative;
            }

            return OlResponseStatus.olResponseNone;
        }

        /// <summary>
        /// Gets the type of the recipient.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <returns>The Meeting Recipient Type.</returns>
        public static OlMeetingRecipientType GetRecipientType(this Who.AttendeeType type)
        {
            if (type == null) return OlMeetingRecipientType.olOptional;

            switch (type.Value)
            {
                case Who.AttendeeType.EVENT_REQUIRED:
                    return OlMeetingRecipientType.olRequired;
                case Who.AttendeeType.EVENT_OPTIONAL:
                    return OlMeetingRecipientType.olOptional;
            }

            return OlMeetingRecipientType.olOptional;
        }

        #endregion

        #region Outlook > Google

        /// <summary>
        /// Determines wether the calendar items are mergable.
        /// </summary>
        /// <param name="calendarItems">The calendar items.</param>
        /// <param name="calendarItem">The calendar item.</param>
        /// <returns>The mergeable Calendar Items.</returns>
        public static IEnumerable<EventEntry> Mergeable(this IEnumerable<EventEntry> calendarItems, AppointmentItem calendarItem)
        {
            return calendarItems.Where(c => Mergeable(calendarItem, c));
        }

        /// <summary>
        /// Merges the with.
        /// </summary>
        /// <param name="googleCalendarItem">The google calendar item.</param>
        /// <param name="outlookCalendarItem">The outlook calendar item.</param>
        /// <returns>True if Changed.</returns>
        public static bool MergeWith(this EventEntry googleCalendarItem, AppointmentItem outlookCalendarItem)
        {
            var result = false;

            result |= googleCalendarItem.Title.ApplyProperty(g => g.Text, outlookCalendarItem.Subject);
            result |= googleCalendarItem.Content.ApplyProperty(g => g.Content, outlookCalendarItem.Body);
            
            if (googleCalendarItem.EventTransparency == null) googleCalendarItem.EventTransparency = new EventEntry.Transparency();
            result |= googleCalendarItem.EventTransparency.ApplyProperty(g => g.Value, outlookCalendarItem.BusyStatus.GetStatus());

            if (googleCalendarItem.EventVisibility == null) googleCalendarItem.EventVisibility = new EventEntry.Visibility();
            result |= googleCalendarItem.EventVisibility.ApplyProperty(g => g.Value, outlookCalendarItem.Sensitivity.GetStatus());
            
            result |= googleCalendarItem.Locations.Merge(outlookCalendarItem.Location);
            result |= googleCalendarItem.Participants.Merge(outlookCalendarItem);

            if (outlookCalendarItem.RecurrenceState == OlRecurrenceState.olApptNotRecurring)
            {
                When time = googleCalendarItem.Times.FirstOrInstance();
                if (!googleCalendarItem.Times.Any())
                    googleCalendarItem.Times.Add(time);

                result |= time.ApplyProperty(t => t.AllDay, outlookCalendarItem.AllDayEvent);
                result |= time.ApplyProperty(t => t.StartTime, outlookCalendarItem.Start);
                result |= time.ApplyProperty(t => t.EndTime, outlookCalendarItem.End);

                if (outlookCalendarItem.ReminderSet)
                {
                    Google.GData.Extensions.Reminder reminder = time.Reminders.FirstOrInstance(t => t.Method == Google.GData.Extensions.Reminder.ReminderMethod.alert);
                    var timespan = TimeSpan.FromMinutes(outlookCalendarItem.ReminderMinutesBeforeStart);
                    result |= reminder.ApplyProperty(r => r.Method, Google.GData.Extensions.Reminder.ReminderMethod.alert);
                    result |= reminder.ApplyProperty(r => r.Minutes, timespan.Minutes);
                    result |= reminder.ApplyProperty(r => r.Hours, timespan.Hours);
                    result |= reminder.ApplyProperty(r => r.Days, timespan.Days);
                }
            }
            else
            {
                if (googleCalendarItem.Recurrence == null) googleCalendarItem.Recurrence = new Recurrence();
                result |= googleCalendarItem.Recurrence.ApplyProperty(r => r.Value, RecurrenceSerializer.Serialize(outlookCalendarItem.GetRecurrencePattern(), outlookCalendarItem.Start, outlookCalendarItem.AllDayEvent));
            }

            return result;
        }

        /// <summary>
        /// Merges the specified locations.
        /// </summary>
        /// <param name="locations">The locations.</param>
        /// <param name="location">The location.</param>
        /// <returns>True if Changed.</returns>
        public static bool Merge(this ExtensionCollection<Where> locations, string location)
        {
            if (!string.IsNullOrWhiteSpace(location) && !locations.Any(e => e.ValueString == location))
            {
                locations.Add(new Where { ValueString = location, Rel = Where.RelType.EVENT });
                return true;
            }

            return false;
        }

        /// <summary>
        /// Merges the specified participants.
        /// </summary>
        /// <param name="participants">The participants.</param>
        /// <param name="outlookCalendarItem">The outlook calendar item.</param>
        /// <returns>True if Changed.</returns>
        public static bool Merge(this ExtensionCollection<Who> participants, AppointmentItem outlookCalendarItem)
        {
            var result = false;

            foreach (Recipient recipient in outlookCalendarItem.Recipients.Cast<Recipient>().Where(recipient => !participants.Any(e => e.Email == (recipient.Address ?? recipient.Name))))
            {
                participants.Add(new Who
                                     {
                                         Attendee_Type = ((OlMeetingRecipientType)Enum.Parse(typeof(OlMeetingRecipientType), recipient.Type.ToString())).GetRecipientType(),
                                         Email = recipient.Address ?? recipient.Name,
                                         Rel = Who.RelType.EVENT_ATTENDEE
                                     });
                result |= true;
            }
            
            return result;
        }

        /// <summary>
        /// Gets the status.
        /// </summary>
        /// <param name="status">The status.</param>
        /// <returns>The status string.</returns>
        public static string GetStatus(this OlBusyStatus status)
        {
            switch (status)
            {
                case OlBusyStatus.olFree:
                case OlBusyStatus.olTentative:
                    return EventEntry.Transparency.TRANSPARENT_VALUE;
                case OlBusyStatus.olBusy:
                case OlBusyStatus.olOutOfOffice:
                    return EventEntry.Transparency.OPAQUE_VALUE;
            }

            return EventEntry.Transparency.OPAQUE_VALUE;
        }

        /// <summary>
        /// Gets the status.
        /// </summary>
        /// <param name="status">The status.</param>
        /// <returns>The visibility string.</returns>
        public static string GetStatus(this OlSensitivity status)
        {
            switch (status)
            {
                case OlSensitivity.olNormal:
                    return EventEntry.Visibility.DEFAULT_VALUE;
                case OlSensitivity.olPrivate:
                    return EventEntry.Visibility.PRIVATE_VALUE;
                case OlSensitivity.olConfidential:
                    return EventEntry.Visibility.CONFIDENTIAL_VALUE;
            }

            return EventEntry.Visibility.DEFAULT_VALUE;
        }

        /// <summary>
        /// Gets the status.
        /// </summary>
        /// <param name="status">The status.</param>
        /// <returns>The status string.</returns>
        public static string GetStatus(this OlResponseStatus status)
        {
            switch (status)
            {
                case OlResponseStatus.olResponseAccepted:
                    return EventEntry.EventStatus.CONFIRMED_VALUE;
                case OlResponseStatus.olResponseDeclined:
                    return EventEntry.EventStatus.CANCELED_VALUE;
                case OlResponseStatus.olResponseTentative:
                    return EventEntry.EventStatus.TENTATIVE_VALUE;
            }

            return EventEntry.EventStatus.TENTATIVE_VALUE;
        }

        /// <summary>
        /// Gets the type of the recipient.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <returns>The AttendeeType.</returns>
        public static Who.AttendeeType GetRecipientType(this OlMeetingRecipientType type)
        {
            switch (type)
            {
                case OlMeetingRecipientType.olRequired:
                    return new Who.AttendeeType { Value = Who.AttendeeType.EVENT_REQUIRED };
                case OlMeetingRecipientType.olOptional:
                    return new Who.AttendeeType { Value = Who.AttendeeType.EVENT_OPTIONAL };
            }

            return new Who.AttendeeType { Value = Who.AttendeeType.EVENT_REQUIRED };
        }

        #endregion
    }
}
