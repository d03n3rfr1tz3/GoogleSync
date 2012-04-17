
namespace DirkSarodnick.GoogleSync.Core.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Net;
    using Data;
    using Google.Contacts;
    using Google.GData.Client;
    using Google.GData.Contacts;
    using Google.GData.Extensions;
    using Microsoft.Office.Interop.Outlook;

    /// <summary>
    /// Defines the Contact Extensions.
    /// </summary>
    public static class ContactExtensions
    {
        private static readonly string[] DateFormats = new[] { @"yyyy-MM-dd", @"--MM-dd" };

        /// <summary>
        /// Determines wether the contacts are mergable.
        /// </summary>
        /// <param name="outlookContact">The outlook contact.</param>
        /// <param name="googleContact">The google contact.</param>
        /// <returns></returns>
        public static bool Mergeable(ContactItem outlookContact, Contact googleContact)
        {
            return outlookContact.UserProperties.GetProperty("GoogleId") == googleContact.Id ||
                   googleContact.Emails.Any(g => g.Address == outlookContact.Email1Address) ||
                   googleContact.Emails.Any(g => g.Address == outlookContact.Email2Address) ||
                   googleContact.Emails.Any(g => g.Address == outlookContact.Email3Address) ||
                   (googleContact.Name.FullName == outlookContact.FullName && googleContact.Phonenumbers.Any(g => g.Value == outlookContact.BusinessTelephoneNumber)) ||
                   (googleContact.Name.FullName == outlookContact.FullName && googleContact.Phonenumbers.Any(g => g.Value == outlookContact.Business2TelephoneNumber)) ||
                   (googleContact.Name.FullName == outlookContact.FullName && googleContact.Phonenumbers.Any(g => g.Value == outlookContact.HomeTelephoneNumber)) ||
                   (googleContact.Name.FullName == outlookContact.FullName && googleContact.Phonenumbers.Any(g => g.Value == outlookContact.Home2TelephoneNumber)) ||
                   (googleContact.Name.FullName == outlookContact.FullName && googleContact.Phonenumbers.Any(g => g.Value == outlookContact.OtherTelephoneNumber));
        }

        #region Google > Outlook

        /// <summary>
        /// Determines wether the contacts are mergable.
        /// </summary>
        /// <param name="outlookContacts">The outlook contacts.</param>
        /// <param name="googleContact">The google contact.</param>
        /// <returns>
        /// The mergeable Contacts.
        /// </returns>
        public static IEnumerable<ContactItem> Mergeable(this IEnumerable<ContactItem> outlookContacts, Contact googleContact)
        {
            return outlookContacts.Where(c => Mergeable(c, googleContact));
        }

        /// <summary>
        /// Merges the with.
        /// </summary>
        /// <param name="outlookContact">The outlook contact.</param>
        /// <param name="googleContact">The google contact.</param>
        /// <param name="outlookGroups">The outlook groups.</param>
        /// <returns>
        /// True if Changed.
        /// </returns>
        public static bool MergeWith(this ContactItem outlookContact, Contact googleContact, IEnumerable<Group> outlookGroups)
        {
            var result = false;

            result |= outlookContact.ApplyProperty(c => c.FullName, googleContact.Name.FullName);
            result |= outlookContact.ApplyProperty(c => c.FirstName, googleContact.Name.GivenName);
            result |= outlookContact.ApplyProperty(c => c.LastName, googleContact.Name.FamilyName);

            result |= outlookContact.ApplyProperty(c => c.Email1Address, googleContact.Emails.FirstOrInstance(e => e.Primary).Address);
            result |= outlookContact.ApplyProperty(c => c.Email2Address, googleContact.Emails.FirstOrInstance(e => !e.Primary).Address);
            result |= outlookContact.ApplyProperty(c => c.Email3Address, googleContact.Emails.FirstOrInstance(e => !e.Primary && e.Address != outlookContact.Email2Address).Address);

            result |= outlookContact.ApplyProperty(c => c.PrimaryTelephoneNumber, (googleContact.Phonenumbers.FirstOrDefault(p => p.Primary) ?? googleContact.Phonenumbers.FirstOrInstance()).Value.FormatPhone());
            result |= outlookContact.ApplyProperty(c => c.HomeTelephoneNumber, googleContact.Phonenumbers.FirstOrInstance(p => p.Rel == ContactsRelationships.IsHome).Value.FormatPhone());
            result |= outlookContact.ApplyProperty(c => c.HomeFaxNumber, googleContact.Phonenumbers.FirstOrInstance(p => p.Rel == ContactsRelationships.IsHomeFax).Value.FormatPhone());
            result |= outlookContact.ApplyProperty(c => c.BusinessTelephoneNumber, googleContact.Phonenumbers.FirstOrInstance(p => p.Rel == ContactsRelationships.IsWork).Value.FormatPhone());
            result |= outlookContact.ApplyProperty(c => c.BusinessFaxNumber, googleContact.Phonenumbers.FirstOrInstance(p => p.Rel == ContactsRelationships.IsWorkFax).Value.FormatPhone());
            result |= outlookContact.ApplyProperty(c => c.OtherFaxNumber, googleContact.Phonenumbers.FirstOrInstance(p => p.Rel == ContactsRelationships.IsOther).Value.FormatPhone());
            result |= outlookContact.ApplyProperty(c => c.OtherTelephoneNumber, googleContact.Phonenumbers.FirstOrInstance(p => p.Rel == ContactsRelationships.IsFax).Value.FormatPhone());
            result |= outlookContact.ApplyProperty(c => c.MobileTelephoneNumber, googleContact.Phonenumbers.FirstOrInstance(p => p.Rel == ContactsRelationships.IsMobile).Value.FormatPhone());

            var primaryMailAddress = googleContact.PostalAddresses.FirstOrDefault(p => p.Primary) ?? googleContact.PostalAddresses.FirstOrInstance();
            result |= outlookContact.ApplyProperty(c => c.MailingAddressStreet, string.Format("{0} {1}", primaryMailAddress.Street, primaryMailAddress.Housename));
            result |= outlookContact.ApplyProperty(c => c.MailingAddressCity, primaryMailAddress.City);
            result |= outlookContact.ApplyProperty(c => c.MailingAddressCountry, primaryMailAddress.Country);
            result |= outlookContact.ApplyProperty(c => c.MailingAddressPostalCode, primaryMailAddress.Postcode);
            result |= outlookContact.ApplyProperty(c => c.MailingAddressPostOfficeBox, primaryMailAddress.Pobox);
            result |= outlookContact.ApplyProperty(c => c.MailingAddressState, primaryMailAddress.Region);

            DateTime birth;
            if (DateTime.TryParseExact(googleContact.ContactEntry.Birthday, DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out birth))
            {
                result |= outlookContact.ApplyProperty(c => c.Birthday, birth);
            }

            result |= outlookContact.ApplyProperty(c => c.BillingInformation, googleContact.ContactEntry.BillingInformation);
            result |= outlookContact.ApplyProperty(c => c.IMAddress, (googleContact.IMs.FirstOrDefault(i => i.Primary) ?? googleContact.IMs.FirstOrInstance()).Address);
            result |= outlookContact.ApplyProperty(c => c.Initials, googleContact.ContactEntry.Initials);
            result |= outlookContact.ApplyProperty(c => c.Language, googleContact.ContactEntry.Language);
            result |= outlookContact.ApplyProperty(c => c.Mileage, googleContact.ContactEntry.Mileage);
            result |= outlookContact.ApplyProperty(c => c.NickName, googleContact.ContactEntry.Nickname);
            result |= outlookContact.ApplyProperty(c => c.WebPage, (googleContact.ContactEntry.Websites.FirstOrDefault(w => w.Primary) ?? googleContact.ContactEntry.Websites.FirstOrInstance()).Href);
            result |= outlookContact.ApplyProperty(c => c.PersonalHomePage, googleContact.ContactEntry.Websites.FirstOrInstance(w => w.Rel == "home-page").Href);
            result |= outlookContact.ApplyProperty(c => c.BusinessHomePage, googleContact.ContactEntry.Websites.FirstOrInstance(w => w.Rel == "work").Href);
            
            var organization = googleContact.ContactEntry.Organizations.FirstOrDefault(o => o.Primary) ?? googleContact.ContactEntry.Organizations.FirstOrInstance();
            result |= outlookContact.ApplyProperty(c => c.CompanyName, organization.Name);
            result |= outlookContact.ApplyProperty(c => c.Department, organization.Department);
            result |= outlookContact.ApplyProperty(c => c.Profession, organization.Title);

            // Syncing Groups/Categories
            var contactGroups = googleContact.GroupMembership.Select(g => outlookGroups.FirstOrInstance(m => m.Id == g.HRef));
            result |= outlookContact.ApplyProperty(c => c.Categories, string.Join("; ", contactGroups.Select(g => (string.IsNullOrEmpty(g.SystemGroup) ? g.Title : g.SystemGroup))));

            return result;
        }

        /// <summary>
        /// Adds the picture.
        /// </summary>
        /// <param name="outlookContact">The outlook contact.</param>
        /// <param name="stream">The stream.</param>
        /// <returns>True if Changed.</returns>
        public static bool AddPicture(this ContactItem outlookContact, Stream stream)
        {
            if (stream == null || stream.Length == 0)
                return false;

            var imageDir = Environment.GetFolderPath(Environment.SpecialFolder.InternetCache, Environment.SpecialFolderOption.None);
            var imagePath = Path.Combine(imageDir, string.Format("contact_{0}", outlookContact.EntryID));
            if (!Directory.Exists(imageDir)) Directory.CreateDirectory(imageDir);

            var output = File.Create(imagePath);
            stream.Seek(0, SeekOrigin.Begin);
            stream.CopyTo(output);
            output.Close();

            outlookContact.AddPicture(imagePath);

            var attachment = outlookContact.Attachments.Cast<Attachment>().FirstOrDefault(a => a.FileName == "ContactPicture.jpg");
            if (attachment != null)
            {
                outlookContact.UserProperties.SetProperty("GooglePictureSize", attachment.Size.ToString(CultureInfo.InvariantCulture));
            }

            return true;
        }

        /// <summary>
        /// Gets the picture.
        /// </summary>
        /// <param name="request">The request.</param>
        /// <param name="contact">The contact.</param>
        /// <returns>The Picture stream.</returns>
        public static Stream GetPicture(this ContactsRequest request, Contact contact)
        {
            var stream = new MemoryStream();
            Stream retStream = null;

            try
            {
                if (contact.PhotoUri != null)
                {
                    retStream = request.Service.Query(contact.PhotoUri);
                }
            }
            catch (GDataRequestException ex)
            {
                var response = ex.Response as HttpWebResponse;
                if (response != null && response.StatusCode != HttpStatusCode.NotFound)
                {
                    throw;
                }
            }

            if (retStream != null)
            {
                retStream.CopyTo(stream);
            }

            return stream;
        }

        #endregion

        #region Outlook > Google

        /// <summary>
        /// Determines wether the contacts are mergable.
        /// </summary>
        /// <param name="googleContacts">The google contacts.</param>
        /// <param name="outlookContact">The outlook contact.</param>
        /// <returns>
        /// The mergeable Contacts.
        /// </returns>
        public static IEnumerable<Contact> Mergeable(this IEnumerable<Contact> googleContacts, ContactItem outlookContact)
        {
            return googleContacts.Where(c => Mergeable(outlookContact, c));
        }

        /// <summary>
        /// Merges the with.
        /// </summary>
        /// <param name="googleContact">The google contact.</param>
        /// <param name="outlookContact">The outlook contact.</param>
        /// <param name="googleGroups">The google groups.</param>
        /// <returns>
        /// True if Changed.
        /// </returns>
        public static bool MergeWith(this Contact googleContact, ContactItem outlookContact, IEnumerable<Group> googleGroups)
        {
            var result = false;

            result |= googleContact.Name.ApplyProperty(c => c.FullName, outlookContact.FullName);
            result |= googleContact.Name.ApplyProperty(c => c.GivenName, outlookContact.FirstName);
            result |= googleContact.Name.ApplyProperty(c => c.FamilyName, outlookContact.LastName);

            if (outlookContact.Email1AddressType != "EX") result |= googleContact.Emails.Merge(new EMail { Address = outlookContact.Email1Address, Rel = ContactsRelationships.IsOther });
            if (outlookContact.Email2AddressType != "EX") result |= googleContact.Emails.Merge(new EMail { Address = outlookContact.Email2Address, Rel = ContactsRelationships.IsOther });
            if (outlookContact.Email3AddressType != "EX") result |= googleContact.Emails.Merge(new EMail { Address = outlookContact.Email3Address, Rel = ContactsRelationships.IsOther });

            result |= googleContact.Phonenumbers.Merge(new PhoneNumber { Value = outlookContact.HomeTelephoneNumber, Rel = ContactsRelationships.IsHome });
            result |= googleContact.Phonenumbers.Merge(new PhoneNumber { Value = outlookContact.Home2TelephoneNumber, Rel = ContactsRelationships.IsHome });
            result |= googleContact.Phonenumbers.Merge(new PhoneNumber { Value = outlookContact.HomeFaxNumber, Rel = ContactsRelationships.IsHomeFax });
            result |= googleContact.Phonenumbers.Merge(new PhoneNumber { Value = outlookContact.BusinessTelephoneNumber, Rel = ContactsRelationships.IsWork });
            result |= googleContact.Phonenumbers.Merge(new PhoneNumber { Value = outlookContact.Business2TelephoneNumber, Rel = ContactsRelationships.IsWork });
            result |= googleContact.Phonenumbers.Merge(new PhoneNumber { Value = outlookContact.BusinessFaxNumber, Rel = ContactsRelationships.IsWorkFax });
            result |= googleContact.Phonenumbers.Merge(new PhoneNumber { Value = outlookContact.OtherTelephoneNumber, Rel = ContactsRelationships.IsOther });
            result |= googleContact.Phonenumbers.Merge(new PhoneNumber { Value = outlookContact.OtherFaxNumber, Rel = ContactsRelationships.IsFax });
            result |= googleContact.Phonenumbers.Merge(new PhoneNumber { Value = outlookContact.MobileTelephoneNumber, Rel = ContactsRelationships.IsMobile });
            if (!googleContact.Phonenumbers.Any(p => p.Primary)) googleContact.Phonenumbers.FirstOrInstance(p => p.Value == outlookContact.PrimaryTelephoneNumber).Primary = true;

            result |= googleContact.ContactEntry.PostalAddresses.Merge(new StructuredPostalAddress { Street = outlookContact.HomeAddressStreet, City = outlookContact.HomeAddressCity, Country = outlookContact.HomeAddressCountry, Pobox = outlookContact.HomeAddressPostOfficeBox, Postcode = outlookContact.HomeAddressPostalCode, Region = outlookContact.HomeAddressState, Rel = ContactsRelationships.IsHome });
            result |= googleContact.ContactEntry.PostalAddresses.Merge(new StructuredPostalAddress { Street = outlookContact.BusinessAddressStreet, City = outlookContact.BusinessAddressCity, Country = outlookContact.BusinessAddressCountry, Pobox = outlookContact.BusinessAddressPostOfficeBox, Postcode = outlookContact.BusinessAddressPostalCode, Region = outlookContact.BusinessAddressState, Rel = ContactsRelationships.IsWork });
            result |= googleContact.ContactEntry.PostalAddresses.Merge(new StructuredPostalAddress { Street = outlookContact.OtherAddressStreet, City = outlookContact.OtherAddressCity, Country = outlookContact.OtherAddressCountry, Pobox = outlookContact.OtherAddressPostOfficeBox, Postcode = outlookContact.OtherAddressPostalCode, Region = outlookContact.OtherAddressState, Rel = ContactsRelationships.IsOther });
            if (!googleContact.PostalAddresses.Any(p => p.Primary))
            {
                googleContact.PostalAddresses.FirstOrInstance(p => outlookContact.MailingAddressStreet.StartsWith(p.Street) && p.City == outlookContact.MailingAddressCity && p.Country == outlookContact.MailingAddressCountry && p.Pobox == outlookContact.MailingAddressPostOfficeBox && p.Postcode == outlookContact.MailingAddressPostalCode && p.Region == outlookContact.MailingAddressState).Primary = true;
            }

            if (outlookContact.Birthday != default(DateTime))
            {
                var birth = outlookContact.Birthday.Year == default(DateTime).Year
                                ? outlookContact.Birthday.ToString(DateFormats[1], CultureInfo.InvariantCulture)
                                : outlookContact.Birthday.ToString(DateFormats[0], CultureInfo.InvariantCulture);
                result |= googleContact.ContactEntry.ApplyProperty(c => c.Birthday, birth);
            }

            result |= googleContact.ContactEntry.ApplyProperty(c => c.BillingInformation, outlookContact.BillingInformation);
            result |= googleContact.IMs.Merge(new IMAddress { Address = outlookContact.IMAddress, Rel = ContactsRelationships.IsOther });
            result |= googleContact.ContactEntry.ApplyProperty(c => c.Initials, outlookContact.Initials);
            result |= googleContact.ContactEntry.ApplyProperty(c => c.Language, outlookContact.Language);
            result |= googleContact.ContactEntry.ApplyProperty(c => c.Mileage, outlookContact.Mileage);
            result |= googleContact.ContactEntry.ApplyProperty(c => c.Nickname, outlookContact.NickName);
            result |= googleContact.ContactEntry.Websites.Merge(new Website { Href = outlookContact.PersonalHomePage, Rel = "home-page" });
            result |= googleContact.ContactEntry.Websites.Merge(new Website { Href = outlookContact.BusinessHomePage, Rel = "work" });
            result |= googleContact.ContactEntry.Websites.Merge(new Website { Href = outlookContact.WebPage, Rel = "other" });

            result |= googleContact.ContactEntry.Organizations.Merge(new Organization { Name = outlookContact.CompanyName, Department = outlookContact.Department, Title = outlookContact.Profession, Rel = ContactsRelationships.IsWork });

            // Syncing Groups/Categories
            result |= googleContact.GroupMembership.Merge(outlookContact.Categories.Split(';').Select(c => c.Trim()), googleGroups);

            return result;
        }

        /// <summary>
        /// Gets the picture.
        /// </summary>
        /// <param name="outlookContact">The outlook contact.</param>
        /// <returns>The Picture Stream.</returns>
        public static Stream GetPicture(this ContactItem outlookContact)
        {
            var stream = new MemoryStream();
            var imagePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "/Temp/", string.Format("export_{0}", outlookContact.EntryID));

            var attachment = outlookContact.Attachments.Cast<Attachment>().FirstOrDefault(a => a.FileName == "ContactPicture.jpg");
            if (attachment != null)
            {
                outlookContact.UserProperties.SetProperty("GooglePictureSize", attachment.Size.ToString(CultureInfo.InvariantCulture));
                attachment.SaveAsFile(imagePath);

                var tempStream = new FileStream(imagePath, FileMode.Open);
                tempStream.CopyTo(stream);
                tempStream.Close();
                tempStream.Dispose();
                File.Delete(imagePath);

                stream.Seek(0, SeekOrigin.Begin);
            }
            return stream;
        }

        /// <summary>
        /// Merges the specified emails.
        /// </summary>
        /// <param name="emails">The emails.</param>
        /// <param name="mail">The mail.</param>
        /// <returns>True if Changed.</returns>
        public static bool Merge(this ExtensionCollection<EMail> emails, EMail mail)
        {
            if (!string.IsNullOrWhiteSpace(mail.Address) && !emails.Any(e => e.Address == mail.Address))
            {
                emails.Add(mail);

                if (emails.Any() && !emails.Any(e => e.Primary))
                {
                    emails.First().Primary = true;
                }

                return true;
            }

            return false;
        }

        /// <summary>
        /// Merges the specified ims.
        /// </summary>
        /// <param name="ims">The ims.</param>
        /// <param name="im">The im.</param>
        /// <returns>True if Changed.</returns>
        public static bool Merge(this ExtensionCollection<IMAddress> ims, IMAddress im)
        {
            if (!string.IsNullOrWhiteSpace(im.Address) && !ims.Any(e => e.Address == im.Address))
            {
                ims.Add(im);

                if (ims.Any() && !ims.Any(e => e.Primary))
                {
                    ims.First().Primary = true;
                }

                return true;
            }

            return false;
        }

        /// <summary>
        /// Merges the specified websites.
        /// </summary>
        /// <param name="websites">The websites.</param>
        /// <param name="website">The website.</param>
        /// <returns>True if Changed.</returns>
        public static bool Merge(this ExtensionCollection<Website> websites, Website website)
        {
            if (!string.IsNullOrWhiteSpace(website.Href) && !websites.Any(e => e.Href == website.Href))
            {
                websites.Add(website);

                if (websites.Any() && !websites.Any(e => e.Primary))
                {
                    websites.First().Primary = true;
                }

                return true;
            }

            return false;
        }

        /// <summary>
        /// Merges the specified organizations.
        /// </summary>
        /// <param name="organizations">The organizations.</param>
        /// <param name="organization">The organization.</param>
        /// <returns>True if Changed.</returns>
        public static bool Merge(this ExtensionCollection<Organization> organizations, Organization organization)
        {
            if ((!string.IsNullOrWhiteSpace(organization.Name) || !string.IsNullOrWhiteSpace(organization.Department)))
            {
                var result = false;

                if (organizations.Any(e => e.Name == organization.Name))
                {
                    var org = organizations.First(e => e.Name == organization.Name);
                    result |= org.ApplyProperty(o => o.Department, organization.Department);
                    result |= org.ApplyProperty(o => o.Title, organization.Title);
                }
                else
                {
                    organizations.Add(organization);
                    result = true;
                }

                if (organizations.Any() && !organizations.Any(e => e.Primary))
                {
                    organizations.First().Primary = true;
                }

                return result;
            }

            return false;
        }

        /// <summary>
        /// Merges the specified addresses.
        /// </summary>
        /// <param name="addresses">The addresses.</param>
        /// <param name="address">The address.</param>
        /// <returns>True if Changed.</returns>
        public static bool Merge(this ExtensionCollection<StructuredPostalAddress> addresses, StructuredPostalAddress address)
        {
            if (!string.IsNullOrWhiteSpace(address.Street) && !string.IsNullOrWhiteSpace(address.City) && !string.IsNullOrWhiteSpace(address.Postcode) && !string.IsNullOrWhiteSpace(address.Country) && 
                !addresses.Any(e => (address.Street != null && address.Street.StartsWith(e.Street)) && e.City == address.City && e.Postcode == address.Postcode && e.Country == address.Country))
            {
                addresses.Add(address);

                return true;
            }

            return false;
        }

        /// <summary>
        /// Merges the specified phone numbers.
        /// </summary>
        /// <param name="phoneNumbers">The phone numbers.</param>
        /// <param name="phone">The phone.</param>
        /// <returns>True if Changed.</returns>
        public static bool Merge(this ExtensionCollection<PhoneNumber> phoneNumbers, PhoneNumber phone)
        {
            if (!string.IsNullOrWhiteSpace(phone.Value) && !phoneNumbers.Any(e => e.Value == phone.Value))
            {
                phoneNumbers.Add(phone);

                return true;
            }

            return false;
        }

        /// <summary>
        /// Merges the specified google groups.
        /// </summary>
        /// <param name="groups">The groups.</param>
        /// <param name="outlookGroups">The outlook groups.</param>
        /// <param name="googleGroups">All google groups.</param>
        /// <returns>
        /// True if Changed.
        /// </returns>
        public static bool Merge(this ExtensionCollection<GroupMembership> groups, IEnumerable<string> outlookGroups, IEnumerable<Group> googleGroups)
        {
            return outlookGroups.Aggregate(false, (current, outlookGroup) => current | Merge(groups, outlookGroup, googleGroups));
        }

        /// <summary>
        /// Merges the specified google groups.
        /// </summary>
        /// <param name="groups">The groups.</param>
        /// <param name="outlookGroup">The outlook group.</param>
        /// <param name="googleGroups">All google groups.</param>
        /// <returns>
        /// True if Changed.
        /// </returns>
        public static bool Merge(this ExtensionCollection<GroupMembership> groups, string outlookGroup, IEnumerable<Group> googleGroups)
        {
            if (!string.IsNullOrWhiteSpace(outlookGroup))
            {
                var outlookGroupAsGoogleGroup = googleGroups.FirstOrInstance(g => (string.IsNullOrEmpty(g.SystemGroup) ? g.Title : g.SystemGroup) == outlookGroup);
                if (!string.IsNullOrWhiteSpace(outlookGroupAsGoogleGroup.Id))
                {
                    groups.Add(new GroupMembership { HRef = outlookGroupAsGoogleGroup.Id });

                    return true;
                }
            }

            return false;
        }

        #endregion

        /// <summary>
        /// Determines whether [has new picture] [the specified contact].
        /// </summary>
        /// <param name="contact">The contact.</param>
        /// <returns>
        ///  <c>true</c> if [has new picture] [the specified contact]; otherwise, <c>false</c>.
        /// </returns>
        public static bool HasNewPicture(this ContactItem contact)
        {
            if (contact.HasPicture)
            {
                var attachment = contact.Attachments.Cast<Attachment>().FirstOrDefault(a => a.FileName == "ContactPicture.jpg");
                return attachment != null && attachment.Size.ToString(CultureInfo.InvariantCulture) != contact.UserProperties.GetProperty("GooglePictureSize");
            }

            return false;
        }
    }
}
