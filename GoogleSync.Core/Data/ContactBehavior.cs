
namespace DirkSarodnick.GoogleSync.Core.Data
{
    /// <summary>
    /// Defines the ContactBehavior class.
    /// </summary>
    public enum ContactBehavior
    {
        /// <summary>
        /// Automatic merging of Contacts.
        /// </summary>
        Automatic,

        /// <summary>
        /// Googles overwrites Outlook
        /// </summary>
        GoogleOverOutlook,

        /// <summary>
        /// Outlook overwrites Google
        /// </summary>
        OutlookOverGoogle
    }
}