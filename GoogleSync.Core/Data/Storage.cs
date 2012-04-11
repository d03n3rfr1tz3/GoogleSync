
namespace DirkSarodnick.GoogleSync.Core.Data
{
    using Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Defines the Storage class.
    /// </summary>
    public static class Storage
    {
        private static StorageItem store;

        /// <summary>
        /// Gets the store.
        /// </summary>
        /// <value>The store.</value>
        private static StorageItem Store
        {
            get
            {
                if (store == null)
                {
                    var folder = ApplicationData.Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                    store = folder.GetStorage("GoogleSync.Settings", OlStorageIdentifierType.olIdentifyBySubject);
                }

                return store;
            }
        }

        /// <summary>
        /// Gets the property.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns>The Property value.</returns>
        public static string GetProperty(string key)
        {
            return Store.UserProperties.GetProperty(key);
        }

        /// <summary>
        /// Sets the property.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="value">The value.</param>
        public static void SetProperty(string key, string value)
        {
            Store.UserProperties.SetProperty(key, value);
            Store.Save();
        }

        /// <summary>
        /// Gets the property.
        /// </summary>
        /// <param name="properties">The properties.</param>
        /// <param name="key">The key.</param>
        /// <returns>The Property value.</returns>
        public static string GetProperty(this UserProperties properties, string key)
        {
            var property = properties.Find(key, true);
            if (property != null)
            {
                return property.Value;
            }

            return null;
        }

        /// <summary>
        /// Sets the property.
        /// </summary>
        /// <param name="properties">The properties.</param>
        /// <param name="key">The key.</param>
        /// <param name="value">The value.</param>
        /// <returns>True if Changed.</returns>
        public static bool SetProperty(this UserProperties properties, string key, string value)
        {
            var property = properties.Find(key, true);
            if (property == null)
            {
                properties.Add(key, OlUserPropertyType.olText, false, System.Type.Missing);
                properties.Find(key, true).Value = value;
                return true;
            }

            if (property.Value != value)
            {
                property.Value = value;
                return true;
            }

            return false;
        }
    }
}