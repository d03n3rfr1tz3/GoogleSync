
namespace DirkSarodnick.GoogleSync.Core.Data
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Google.Contacts;
    using Google.GData.Calendar;
    using Google.GData.Client;
    using Google.GData.Contacts;

    /// <summary>
    /// Defines the GoogleData class.
    /// </summary>
    public class GoogleData : BaseData
    {
        private ContactsRequest contactsRequest;
        private CalendarService calendarService;

        /// <summary>
        /// Gets the contacts request.
        /// </summary>
        /// <value>The contacts request.</value>
        public ContactsRequest ContactsRequest
        {
            get
            {
                if (this.contactsRequest == null)
                {
                    var settings = new RequestSettings(ApplicationData.GoogleApplication, ApplicationData.GoogleUsername, ApplicationData.GooglePassword);
                    this.contactsRequest = new ContactsRequest(settings);
                }

                return this.contactsRequest;
            }
        }

        public CalendarService CalendarService
        {
            get
            {
                if (this.calendarService == null)
                {
                    this.calendarService = new CalendarService(ApplicationData.GoogleApplication);
                    this.calendarService.setUserCredentials(ApplicationData.GoogleUsername, ApplicationData.GooglePassword);
                }

                return this.calendarService;
            }
        }

        /// <summary>
        /// Gets the contacts.
        /// </summary>
        /// <returns>The Contacts.</returns>
        public IEnumerable<Contact> GetContacts()
        {
            try
            {
                return this.ContactsRequest.GetContacts().Entries.ToList();
            }
            catch (GDataRequestException)
            {
                return new List<Contact>();
            }
        }

        /// <summary>
        /// Gets the contact.
        /// </summary>
        /// <param name="googleId">The google id.</param>
        /// <returns>The contact.</returns>
        public Contact GetContact(string googleId)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the calendar items.
        /// </summary>
        /// <returns>The Calendar Items.</returns>
        public IEnumerable<EventEntry> GetCalendarItems()
        {
            var query = new EventQuery(ApplicationData.GoogleCalendarUri.ToString());
            var result = this.CalendarService.Query(query);
            if (result != null)
            {
                return result.Entries.Cast<EventEntry>();
            }

            return new List<EventEntry>();
        }

        /// <summary>
        /// Gets the Google Groups.
        /// </summary>
        /// <returns>The Groups.</returns>
        public IEnumerable<Group> GetGroups()
        {
            var response = ContactsRequest.GetGroups();
            if (response != null)
            {
                return response.Entries.AsEnumerable();
            }

            return new List<Group>();
        }
    }
}
