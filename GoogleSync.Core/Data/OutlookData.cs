
namespace DirkSarodnick.GoogleSync.Core.Data
{
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Defines the OutlookData class.
    /// </summary>
    public class OutlookData : BaseData
    {
        private Items contactItems;
        private Items calendarItems;

        /// <summary>
        /// Gets the contacts folder items.
        /// </summary>
        /// <value>The contacts folder items.</value>
        public Items ContactsFolderItems
        {
            get
            {
                return this.contactItems ?? (this.contactItems = ApplicationData.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts).Items);
            }
        }

        /// <summary>
        /// Gets the calendar folder items.
        /// </summary>
        /// <value>The calendar folder items.</value>
        public Items CalendarFolderItems
        {
            get
            {
                return this.calendarItems ?? (this.calendarItems = ApplicationData.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar).Items);
            }
        }

        /// <summary>
        /// Gets the contacts.
        /// </summary>
        /// <returns>The Contacts.</returns>
        public IEnumerable<ContactItem> GetContacts()
        {
            var items = ApplicationData.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts).Items;
            return items.Cast<ContactItem>();
        }

        /// <summary>
        /// Gets the calendar items.
        /// </summary>
        /// <returns>The Calendar Items.</returns>
        public IEnumerable<AppointmentItem> GetCalendarItems()
        {
            var items = ApplicationData.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar).Items;
            return items.Cast<AppointmentItem>().Where(o => o.RecurrenceState == OlRecurrenceState.olApptNotRecurring || o.RecurrenceState == OlRecurrenceState.olApptMaster);
        }
    }
}