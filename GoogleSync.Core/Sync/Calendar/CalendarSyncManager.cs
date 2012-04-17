
namespace DirkSarodnick.GoogleSync.Core.Sync.Calendar
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using Data;
    using Extensions;
    using Google.GData.Calendar;
    using Google.GData.Client;
    using Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Defines the CalendarSyncManager class.
    /// </summary>
    public class CalendarSyncManager : SyncBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CalendarSyncManager"/> class.
        /// </summary>
        /// <param name="repository">The repository.</param>
        public CalendarSyncManager(DataRepository repository)
            : base(repository)
        {
        }

        /// <summary>
        /// Syncs this instance.
        /// </summary>
        public override void Sync()
        {
            if (string.IsNullOrWhiteSpace(ApplicationData.GoogleCalendarUri.ToString()))
                return;

            if (ApplicationData.ContactBehavior == ContactBehavior.OutlookOverGoogle)
            {
                SyncOutlookToGoogle();
                SyncGoogleToOutlook();
            }
            else
            {
                SyncGoogleToOutlook();
                SyncOutlookToGoogle();
            }
        }

        /// <summary>
        /// Items the changed.
        /// </summary>
        /// <param name="item">The item.</param>
        public override void ItemChanged(object item)
        {
            if (item is AppointmentItem)
            {
                var googleCalendarItems = this.Repository.GoogleData.GetCalendarItems();
                SyncOutlookCalendarItem((AppointmentItem)item, googleCalendarItems);
            }
        }

        /// <summary>
        /// Releases unmanaged and - optionally - managed resources
        /// </summary>
        public override void Dispose()
        {
        }

        /// <summary>
        /// Syncs the google to outlook.
        /// </summary>
        private void SyncGoogleToOutlook()
        {
            var googleCalendarItems = this.Repository.GoogleData.GetCalendarItems().ToList();
            var outlookCalendarItems = this.Repository.OutlookData.GetCalendarItems().ToList();

            foreach (var googleCalendarItem in googleCalendarItems)
            {
                var mergeables = outlookCalendarItems.Mergeable(googleCalendarItem).ToList();

                var outlookCalendarItem = mergeables.Any() ? mergeables.First() : ApplicationData.Application.CreateItem(OlItemType.olAppointmentItem) as AppointmentItem;
                var changed = outlookCalendarItem.MergeWith(googleCalendarItem);
                changed |= outlookCalendarItem.UserProperties.SetProperty("GoogleId", googleCalendarItem.EventId);
                if (changed) outlookCalendarItem.Save();
            }
        }

        /// <summary>
        /// Syncs the outlook to google.
        /// </summary>
        private void SyncOutlookToGoogle()
        {
            var googleCalendarItems = this.Repository.GoogleData.GetCalendarItems().ToList();
            var outlookCalendarItems = this.Repository.OutlookData.GetCalendarItems().ToList();

            foreach (var outlookCalendarItem in outlookCalendarItems)
            {
                SyncOutlookCalendarItem(outlookCalendarItem, googleCalendarItems);
            }
        }

        /// <summary>
        /// Syncs the outlook calendar item.
        /// </summary>
        /// <param name="outlookCalendarItem">The outlook calendar item.</param>
        /// <param name="googleCalendarItems">The google calendar items.</param>
        private void SyncOutlookCalendarItem(AppointmentItem outlookCalendarItem, IEnumerable<EventEntry> googleCalendarItems)
        {
            var mergeables = googleCalendarItems.Mergeable(outlookCalendarItem).ToList();

            EventEntry googleCalendarItem;
            if (mergeables.Any())
            {
                googleCalendarItem = mergeables.First();
                var changed = googleCalendarItem.MergeWith(outlookCalendarItem);

                try
                {
                    if (changed) this.Repository.GoogleData.CalendarService.Update(googleCalendarItem);
                }
                catch (GDataRequestException ex)
                {
                    var response = ex.Response as HttpWebResponse;
                    if (response != null && response.StatusCode != HttpStatusCode.Conflict)
                    {
                        // TODO: Add functionality to resolve update conflict.
                        throw;
                    }
                }
            }
            else
            {
                googleCalendarItem = new EventEntry();
                var changed = googleCalendarItem.MergeWith(outlookCalendarItem);
                if (changed) googleCalendarItem = this.Repository.GoogleData.CalendarService.Insert(ApplicationData.GoogleCalendarUri, googleCalendarItem);
            }

            var outlookChanged = outlookCalendarItem.UserProperties.SetProperty("GoogleId", googleCalendarItem.EventId);
            if (outlookChanged)
                outlookCalendarItem.Save();
        }
    }
}
