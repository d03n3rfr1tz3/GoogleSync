
namespace DirkSarodnick.GoogleSync.Core.Sync
{
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using Data;
    using Extensions;
    using Google.Contacts;
    using Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Defines the ContactSyncManager class.
    /// </summary>
    public class ContactSyncManager : SyncBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ContactSyncManager"/> class.
        /// </summary>
        /// <param name="repository">The repository.</param>
        public ContactSyncManager(DataRepository repository)
            : base(repository)
        {
        }

        /// <summary>
        /// Syncs this instance.
        /// </summary>
        public override void Sync()
        {
            if (string.IsNullOrWhiteSpace(ApplicationData.GoogleUsername))
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
            if (item is ContactItem)
            {
                var googleContacts = ApplicationData.IncludeContactWithoutEmail
                                     ? this.Repository.GoogleData.GetContacts()
                                     : this.Repository.GoogleData.GetContacts().Where(c => c.Emails.Any()).ToList();
                SyncOutlookContact((ContactItem)item, googleContacts);
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
            var googleContacts = ApplicationData.IncludeContactWithoutEmail
                                     ? this.Repository.GoogleData.GetContacts()
                                     : this.Repository.GoogleData.GetContacts().Where(c => c.Emails.Any()).ToList();
            var outlookContacts = this.Repository.OutlookData.GetContacts().ToList();

            foreach (var googleContact in googleContacts)
            {
                var mergeables = outlookContacts.Mergeable(googleContact);

                // continue with next, if modificationdate is older then other side.
                if (ApplicationData.ContactBehavior == ContactBehavior.Automatic && mergeables.Any() &&
                    googleContact.ContactEntry.Edited.DateValue < mergeables.First().LastModificationTime)
                    continue;

                // Get or Create Contact
                ContactItem outlookContact;
                if (mergeables.Any())
                {
                    outlookContact = mergeables.First();
                }
                else
                {
                    outlookContact = ApplicationData.Application.CreateItem(OlItemType.olContactItem) as ContactItem;
                }

                // Set Informations
                var changed = outlookContact.MergeWith(googleContact, Repository.GoogleData.GetGroups());
                changed |= outlookContact.UserProperties.SetProperty("GoogleId", googleContact.Id);
                if (ApplicationData.ContactBehavior != ContactBehavior.OutlookOverGoogle && outlookContact.UserProperties.GetProperty("GooglePicture") != googleContact.PhotoEtag)
                {
                    changed |= outlookContact.AddPicture(this.Repository.GoogleData.ContactsRequest.GetPicture(googleContact));
                    changed |= outlookContact.UserProperties.SetProperty("GooglePicture", googleContact.PhotoEtag);
                }

                if (changed)
                    outlookContact.Save();
            }
        }

        /// <summary>
        /// Syncs the outlook to google.
        /// </summary>
        private void SyncOutlookToGoogle()
        {
            var googleContacts = ApplicationData.IncludeContactWithoutEmail
                                     ? this.Repository.GoogleData.GetContacts()
                                     : this.Repository.GoogleData.GetContacts().Where(c => c.Emails.Any()).ToList();
            var outlookContacts = this.Repository.OutlookData.GetContacts().ToList();

            foreach (var outlookContact in outlookContacts)
            {
                SyncOutlookContact(outlookContact, googleContacts);
            }
        }

        /// <summary>
        /// Syncs the outlook contact.
        /// </summary>
        /// <param name="outlookContact">The outlook contact.</param>
        /// <param name="googleContacts">The google contacts.</param>
        private void SyncOutlookContact(ContactItem outlookContact, IEnumerable<Contact> googleContacts)
        {
            var mergeables = googleContacts.Mergeable(outlookContact);

            // continue with next, if modificationdate is older then other side.
            if (ApplicationData.ContactBehavior == ContactBehavior.Automatic && mergeables.Any() &&
                outlookContact.LastModificationTime < mergeables.First().ContactEntry.Edited.DateValue)
                return;

            // Get or Create Contact and merge informations
            Contact googleContact;
            var googleGroups = Repository.GoogleData.GetGroups();
            if (mergeables.Any())
            {
                googleContact = mergeables.First();
                var changed = googleContact.MergeWith(outlookContact, googleGroups);
                if (changed)
                    this.Repository.GoogleData.ContactsRequest.Update(googleContact);
            }
            else
            {
                googleContact = new Contact();
                googleContact.MergeWith(outlookContact, googleGroups);
                this.Repository.GoogleData.ContactsRequest.Insert(ApplicationData.GoogleContactsUri, googleContact);
            }

            // Set GoogleId and Picture
            var outlookChanged = outlookContact.UserProperties.SetProperty("GoogleId", googleContact.Id);
            if (ApplicationData.ContactBehavior != ContactBehavior.GoogleOverOutlook && outlookContact.HasNewPicture())
            {
                this.Repository.GoogleData.ContactsRequest.SetPhoto(googleContact, outlookContact.GetPicture(), "image/jpg");
                googleContact = this.Repository.GoogleData.ContactsRequest.Retrieve(googleContact);
                outlookChanged |= outlookContact.UserProperties.SetProperty("GooglePicture", googleContact.PhotoEtag);
            }

            if (outlookChanged)
                outlookContact.Save();
        }
    }
}