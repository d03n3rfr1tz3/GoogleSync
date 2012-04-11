
namespace DirkSarodnick.GoogleSync.Core.Data
{
    using System;
    using System.Globalization;
    using System.Linq;
    using Google.GData.Contacts;
    using Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Defines the ApplicationData class.
    /// </summary>
    public static class ApplicationData
    {
        /// <summary>
        /// Gets or sets the application.
        /// </summary>
        /// <value>The application.</value>
        public static Application Application { get; set; }

        /// <summary>
        /// Gets or sets the google application.
        /// </summary>
        /// <value>The google application.</value>
        public static string GoogleApplication { get; set; }

        #region Self Saving Properties

        /// <summary>
        /// Gets or sets the google username.
        /// </summary>
        /// <value>The google username.</value>
        public static string GoogleUsername
        {
            get
            {
                return Storage.GetProperty("GoogleUsername");
            }

            set
            {
                Storage.SetProperty("GoogleUsername", value);
            }
        }

        /// <summary>
        /// Gets or sets the google password.
        /// </summary>
        /// <value>The google password.</value>
        public static string GooglePassword
        {
            get
            {
                return Storage.GetProperty("GooglePassword");
            }

            set
            {
                Storage.SetProperty("GooglePassword", value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether [include contact without email].
        /// </summary>
        /// <value>
        ///  <c>true</c> if [include contact without email]; otherwise, <c>false</c>.
        /// </value>
        public static bool IncludeContactWithoutEmail
        {
            get
            {
                return Storage.GetProperty("IncludeContactWithoutEmail") == true.ToString(CultureInfo.InvariantCulture);
            }

            set
            {
                Storage.SetProperty("IncludeContactWithoutEmail", value.ToString(CultureInfo.InvariantCulture));
            }
        }

        /// <summary>
        /// Gets or sets the contact behavior.
        /// </summary>
        /// <value>The contact behavior.</value>
        public static ContactBehavior ContactBehavior
        {
            get
            {
                var property = Storage.GetProperty("ContactBehavior");
                
                ContactBehavior behavior;
                if (Enum.TryParse(property, out behavior))
                {
                    return behavior;
                }

                return ContactBehavior.Automatic;
            }

            set
            {
                Storage.SetProperty("ContactBehavior", value.ToString());
            }
        }

        #endregion

        /// <summary>
        /// Gets or sets the google calendar URI.
        /// </summary>
        /// <value>The google calendar URI.</value>
        public static Uri GoogleCalendarUri
        {
            get
            {
                return new Uri("https://www.google.com/calendar/feeds/default/private/full");
            }
        }

        /// <summary>
        /// Gets the google contacts URI.
        /// </summary>
        /// <value>The google contacts URI.</value>
        public static Uri GoogleContactsUri
        {
            get
            {
                return new Uri(ContactsQuery.CreateContactsUri("default") + "?v=2");
            }
        }

        /// <summary>
        /// Resolves the google account.
        /// </summary>
        public static void ResolveGoogleAccount()
        {
            var googleAccount = GetGoogleAccount();
            if (googleAccount != null && string.IsNullOrWhiteSpace(GoogleUsername))
            {
                GoogleUsername = googleAccount.SmtpAddress;
            }
        }

        /// <summary>
        /// Gets the google account.
        /// </summary>
        /// <returns>The Account.</returns>
        private static Account GetGoogleAccount()
        {
            return Application.Session.Accounts.Cast<Account>().FirstOrDefault(account => account.SmtpAddress.Contains("gmail") || account.SmtpAddress.Contains("googlemail"));
        }
    }
}