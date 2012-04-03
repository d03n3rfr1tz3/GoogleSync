
namespace DirkSarodnick.GoogleSync.Core.Sync
{

    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using Data;

    /// <summary>
    /// Defines the SyncManager class.
    /// </summary>
    public class SyncManager : IDisposable
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="SyncManager"/> class.
        /// </summary>
        public SyncManager()
        {
            this.Repository = new DataRepository();
            this.SyncCollection = new List<SyncBase>
            {
                new ContactSyncManager(this.Repository),
                new CalendarSyncManager(this.Repository)
            };
        }

        /// <summary>
        /// Gets or sets the repository.
        /// </summary>
        /// <value>The repository.</value>
        protected DataRepository Repository { get; set; }

        /// <summary>
        /// Gets or sets the sync collection.
        /// </summary>
        /// <value>The sync collection.</value>
        protected List<SyncBase> SyncCollection { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is busy.
        /// </summary>
        /// <value><c>true</c> if this instance is busy; otherwise, <c>false</c>.</value>
        protected bool IsBusy { get; set; }

        /// <summary>
        /// Starts this instance.
        /// </summary>
        public void Start()
        {
            if (this.IsBusy) return;
            IsBusy = true;

            this.ClearEvents();
            this.SyncCollection.ForEach(s =>
            {
                try
                {
                    s.Sync();
                }
                catch (Exception ex)
                {
                    Debug.Write(ex.Message);
                    new EventLogPermission(EventLogPermissionAccess.Administer, ".").PermitOnly();
                    EventLog.WriteEntry("GoogleSync Addin", ex.ToString(), EventLogEntryType.Warning);
                }
            });
            this.AddEvents();

            IsBusy = false;
        }

        /// <summary>
        /// Clears the events.
        /// </summary>
        public void ClearEvents()
        {
            try
            {
                this.Repository.OutlookData.ContactsFolderItems.ItemAdd -= Items_ItemChange;
                this.Repository.OutlookData.CalendarFolderItems.ItemAdd -= Items_ItemChange;

                this.Repository.OutlookData.ContactsFolderItems.ItemChange -= Items_ItemChange;
                this.Repository.OutlookData.CalendarFolderItems.ItemChange -= Items_ItemChange;
            }
            catch (Exception ex)
            {
                Debug.Write(ex.Message);
            }
        }

        /// <summary>
        /// Adds the events.
        /// </summary>
        public void AddEvents()
        {
            this.Repository.OutlookData.ContactsFolderItems.ItemAdd += Items_ItemChange;
            this.Repository.OutlookData.CalendarFolderItems.ItemAdd += Items_ItemChange;

            this.Repository.OutlookData.ContactsFolderItems.ItemChange += Items_ItemChange;
            this.Repository.OutlookData.CalendarFolderItems.ItemChange += Items_ItemChange;
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            this.SyncCollection.ForEach(s => s.Dispose());
            this.SyncCollection.Clear();
        }

        /// <summary>
        /// Items_s the item change.
        /// </summary>
        /// <param name="item">The item.</param>
        private void Items_ItemChange(object item)
        {
            this.SyncCollection.ForEach(s => s.ItemChanged(item));
        }
    }
}
