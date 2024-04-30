using System;
using System.Linq;
using System.Threading;
using System.IO;
using System.Diagnostics;
using System.Net;
using System.Windows.Forms;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.PeopleService.v1;
using Google.Apis.PeopleService.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Newtonsoft.Json;

namespace googleSync

{
    public class ContactData
    {
        public ContactData(string token, DateTime expire)
        {
            this.Token = token;
            this.Expire = expire;
        }

        public string Token { get; set; }
        public DateTime Expire { get; set; }
    }


    public class SyncClass
    {
        #region Common objects

        UserCredential credential;
        readonly string[] scopes = { CalendarService.Scope.Calendar, PeopleServiceService.Scope.Contacts };
        readonly string ApplicationName = "Google Account - Microsoft Outlook Sync";
        readonly string varPath = @"%AppData%\Google Account - Microsoft Outlook Sync";

        #endregion

        #region Common methods

        public SyncClass()
        {
            try
            {
                string path = Environment.ExpandEnvironmentVariables(varPath);
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                Directory.SetCurrentDirectory(path);
                if (!File.Exists("credentials.json"))
                {
                    File.Copy(@"C:\Program Files (x86)\Google Account - Microsoft Outlook Sync\credentials.json", "credentials.json");
                    Globals.ThisAddIn.Application.Session.AddStore("googleSync.pst");
                    foreach (Store store in Globals.ThisAddIn.Application.Session.Stores)
                    {
                        if (store.FilePath == (Environment.ExpandEnvironmentVariables(varPath) + "\\googleSync.pst"))
                        {
                            store.GetRootFolder().Name = "Google Sync";
                            store.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems).UserDefinedProperties.Add("gId", OlUserPropertyType.olText, OlUserPropertyType.olText);
                        }
                    }
                }
            }
            catch (System.Exception error)
            {
                WriteLog("ERROR: " + error);
                MessageBox.Show("Si è verificato un errore. Consulta il log per ulteriori dettagli.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public UserCredential Login()
        {
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                WriteLog("Credential file saved to: " + credPath);
            }
            return credential;
        }

        public Store GetOStore()
        {
            Stores stores = Globals.ThisAddIn.Application.Session.Stores;
            foreach (Store store in stores)
            {
                if (store.FilePath == (Environment.ExpandEnvironmentVariables(varPath) + "\\googleSync.pst"))
                {
                    return store;
                }
            }
            return null;
        }

        public void WriteLog(string log, params object[] arg)
        {
            using (StreamWriter sw = File.AppendText("log.txt"))
            {
                string format = "[" + DateTime.Now.ToString() + "] - " + log;
                sw.WriteLine(format, arg);
            }
        }

        public void ShowLog()
        {
            Process.Start("notepad", "log.txt");
        }

        #endregion

        #region Methods for calendar sync

        public CalendarService CalendarInit(UserCredential credential)
        {
            CalendarService calendarService = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return calendarService;
        }

        public Dictionary<string, GCalendar> GetCalendarsDictionary()
        {
            if (File.Exists("calendarData.json"))
            {
                string calendarData = File.ReadAllText("calendarData.json");
                return JsonConvert.DeserializeObject<Dictionary<string, GCalendar>>(calendarData);
            }
            else
            {
                return new Dictionary<string, GCalendar>();
            }
        }

        public void CalendarSync(CalendarService calendarService, Store oStore, Dictionary<string, GCalendar> calendars)
        {
            WriteLog("Syncing calendars");

            Folder oCalMain = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            Folder oCalBin = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);

            Folder newCal;

            GCalendar jCalendar;

            List<int> remindersList = new List<int>();

            CalendarListResource.ListRequest gCalRequest = calendarService.CalendarList.List();
            gCalRequest.ShowDeleted = true;
            CalendarList gCalList = gCalRequest.Execute();

            foreach (CalendarListEntry gCalItem in gCalList.Items)
            {
                WriteLog("Calendar {0}, ID {1}", gCalItem.Summary, gCalItem.Id);

                bool gCalDeleted = gCalItem.Deleted != null && (bool)gCalItem.Deleted;
                calendars.TryGetValue(gCalItem.Id, out jCalendar);

                if (gCalDeleted && jCalendar != null)
                {
                    string jCalName = jCalendar.GName;
                    WriteLog("Deleting calendar {0}, ID {1}", jCalName, gCalItem.Id);

                    foreach (Folder personalCalendar in oCalMain.Folders)
                    {
                        if (personalCalendar.Name.Contains(calendars[gCalItem.Id].GName))
                        {
                            personalCalendar.Delete();
                            break;
                        }
                    }

                    foreach (Folder personalCalendar in oCalBin.Folders)
                    {
                        if (personalCalendar.Name.Contains(calendars[gCalItem.Id].GName))
                        {
                            personalCalendar.Delete();
                        }
                    }

                    calendars.Remove(gCalItem.Id);
                }
                else if (!gCalDeleted)
                {
                    bool isOCalPresent = false;

                    foreach (Folder personalCalendar in oCalMain.Folders)
                    {
                        if (personalCalendar.Name.Contains(gCalItem.Summary))
                        {
                            isOCalPresent = true;
                            break;
                        }
                    }
                    if (!isOCalPresent)
                    {
                        newCal = (Folder)oCalMain.Folders.Add(gCalItem.Summary, OlDefaultFolders.olFolderCalendar);
                        newCal.UserDefinedProperties.Add("gId", OlUserPropertyType.olText, OlUserPropertyType.olText);
                        newCal.UserDefinedProperties.Add("gCalId", OlUserPropertyType.olText, OlUserPropertyType.olText);
                    }

                    if (!calendars.ContainsKey(gCalItem.Id))
                    {
                        remindersList.Clear();
                        foreach (EventReminder reminder in gCalItem.DefaultReminders)
                        {
                            if (reminder.Method.Contains("popup"))
                            {
                                remindersList.Add((int)reminder.Minutes);
                            }
                        }

                        calendars.Add(gCalItem.Id, new GCalendar(gCalItem.Summary, gCalItem.AccessRole, "", remindersList.ToArray()));
                    }
                }
            }

            EventsResource.ListRequest gEventRequest;
            Events gEvents;
            Folder oCalendar = null;
            foreach (KeyValuePair<string, GCalendar> calendar in calendars)
            {
                foreach (Folder oCal in oCalMain.Folders)
                {
                    if (oCal.Name.Contains(calendar.Value.GName))
                    {
                        oCalendar = oCal;
                        break;
                    }
                }

                if (calendar.Value.Token == "")
                {
                    WriteLog("Sync token for calendar {0} not present", calendar.Value.GName);
                    gEventRequest = calendarService.Events.List(calendar.Key);
                    gEventRequest.ShowDeleted = false;
                    gEventRequest.SingleEvents = true;
                    gEvents = gEventRequest.Execute();
                }
                else
                {
                    WriteLog("Sync token for calendar {0} present: {1}", calendar.Value.GName, calendar.Value.Token);
                    gEventRequest = calendarService.Events.List(calendar.Key);
                    gEventRequest.SyncToken = calendar.Value.Token;
                    gEventRequest.SingleEvents = true;
                    try
                    {
                        gEvents = gEventRequest.Execute();
                    }
                    catch (System.Exception)
                    {
                        WriteLog("Sync token for calendar {0} expired", calendar.Value.GName);
                        gEventRequest = calendarService.Events.List(calendar.Key);
                        gEventRequest.ShowDeleted = false;
                        gEventRequest.SingleEvents = true;
                        gEvents = gEventRequest.Execute();
                    }
                    
                }

                EventsSync(gEvents, oCalendar, oStore, calendar.Key, calendar.Value.Reminders);

                while (gEvents.NextPageToken != null)
                {
                    gEventRequest = calendarService.Events.List(calendar.Key);
                    gEventRequest.ShowDeleted = false;
                    gEventRequest.SingleEvents = true;
                    gEventRequest.PageToken = gEvents.NextPageToken;
                    gEvents = gEventRequest.Execute();
                    EventsSync(gEvents, oCalendar, oStore, calendar.Key, calendar.Value.Reminders);
                }

                WriteLog("Writing sync token for calendar {0}: {1}", calendar.Value.GName, gEvents.NextSyncToken);
                calendar.Value.Token = gEvents.NextSyncToken;

            }

            File.WriteAllText("calendarData.json", JsonConvert.SerializeObject(calendars, Formatting.Indented));

        }

        private void EventsSync(Events gEvents, Folder oCal, Store oStore, string gCalId, int[] gCalReminders)
        {
            WriteLog("Number of events: {0}", gEvents.Items.Count());
            Folder oCalBin = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
            string filter;
            AppointmentItem oEventItem;
            Items restricted;
            foreach (Google.Apis.Calendar.v3.Data.Event gEventItem in gEvents.Items)
            {
                filter = "[gId] = '" + gEventItem.Id + "'";
                restricted = oCal.Items.Restrict(filter);
                oEventItem = (restricted.Count != 0) ? restricted[1] : null;
                if (oEventItem != null)
                {
                    WriteLog("Event found");
                    if (gEventItem.Status != "cancelled")
                    {
                        WriteLog("Updating appointment with id {0}", gEventItem.Id);
                        
                        oEventItem = PopulateAppointmentFields(oEventItem, gEventItem, gCalReminders);

                        oEventItem.Save();
                    }
                    else if (gEventItem.Status == "cancelled")
                    {
                        WriteLog("Deleting appointment with id {0}", gEventItem.Id);
                        oEventItem.Delete();
                        oEventItem = oCalBin.Items.Find(filter);
                        oEventItem?.Delete();
                    }
                }
                else
                {
                    if (gEventItem.Status != "cancelled")
                    {
                        oEventItem = oCal.Items.Add(OlItemType.olAppointmentItem) as AppointmentItem;
                        WriteLog("Creating appointment with id {0}", gEventItem.Id);

                        oEventItem.UserProperties.Add("gId", OlUserPropertyType.olText, true, OlUserPropertyType.olText).Value = gEventItem.Id;
                        oEventItem.UserProperties.Add("gCalId", OlUserPropertyType.olText, true, OlUserPropertyType.olText).Value = gCalId;

                        oEventItem = PopulateAppointmentFields(oEventItem, gEventItem, gCalReminders);

                        oEventItem.Save();
                    }
                }
            }
        }

        private AppointmentItem PopulateAppointmentFields (AppointmentItem oEvent, Google.Apis.Calendar.v3.Data.Event gEvent, int[] gCalReminders)
        {
            oEvent.Subject = gEvent.Summary;
            oEvent.Location = gEvent.Location;
            oEvent.Body = gEvent.Description;
            oEvent.BusyStatus = (gEvent.Transparency == null || gEvent.Transparency.Contains("opaque")) ? OlBusyStatus.olBusy : OlBusyStatus.olFree;
            if (gEvent.Start.DateTimeDateTimeOffset is DateTimeOffset)
            {
                oEvent.Start = DateTime.Parse(gEvent.Start.DateTimeRaw);
                oEvent.End = DateTime.Parse(gEvent.End.DateTimeRaw);
                oEvent.AllDayEvent = false;
            }
            else
            {
                oEvent.Start = DateTime.Parse(gEvent.Start.Date);
                oEvent.End = DateTime.Parse(gEvent.Start.Date).AddHours(23).AddMinutes(59).AddSeconds(59);
                oEvent.AllDayEvent = true;
            }
            oEvent.ReminderSet = false;
            if (gEvent.Reminders != null && DateTime.Compare(oEvent.Start.AddDays(1.0), DateTime.Now) >= 0)
            {
                if ((bool)gEvent.Reminders.UseDefault && gCalReminders != null)
                {
                    if (gCalReminders.Length != 0)
                    {
                        oEvent.ReminderSet = true;
                        oEvent.ReminderMinutesBeforeStart = gCalReminders[0];
                    }
                }
                else if (!(bool)gEvent.Reminders.UseDefault && gEvent.Reminders.Overrides != null)
                {
                    foreach (EventReminder reminder in gEvent.Reminders.Overrides)
                    {
                        if (reminder.Method.Contains("popup"))
                        {
                            oEvent.ReminderSet = true;
                            oEvent.ReminderMinutesBeforeStart = (int)reminder.Minutes;
                            break;
                        }
                    }
                }
            }

            return oEvent;
        }

        public void CreateOEvent(CalendarService calendarService, Store oStore, Dictionary<string, GCalendar> calendars, string gId)
        {
            Google.Apis.Calendar.v3.Data.Event gEventRequest = new Google.Apis.Calendar.v3.Data.Event();
            EventDateTime gEventStart = new EventDateTime();
            EventDateTime gEventEnd = new EventDateTime();
            AppointmentItem oEvent = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;

            gEventRequest.Summary = oEvent.Subject;
            gEventRequest.Location = oEvent.Location;
            gEventRequest.Description = oEvent.Body;
            gEventRequest.Transparency = (oEvent.BusyStatus == OlBusyStatus.olBusy) ? "opaque" : "transparent";
            if (oEvent.AllDayEvent)
            {
                gEventStart.Date = (oEvent.Start.Date.Year + "-" + oEvent.Start.Date.Month + "-" + oEvent.Start.Date.Day);
                gEventRequest.Start = gEventStart;
                gEventEnd.Date = (oEvent.End.Date.Year + "-" + oEvent.End.Date.Month + "-" + oEvent.End.Date.Day);
                gEventRequest.End = gEventEnd;
            }
            else if (!oEvent.AllDayEvent)
            {
                gEventStart.DateTimeDateTimeOffset = oEvent.Start;
                gEventRequest.Start = gEventStart;
                gEventEnd.DateTimeDateTimeOffset = oEvent.End;
                gEventRequest.End = gEventEnd;
            }


            if (oEvent.UserProperties.Count != 0)
            {
                WriteLog("Event with id {0} updated. Syncing with Google...", oEvent.UserProperties["gId"].Value);
                _ = calendarService.Events.Update(gEventRequest, oEvent.UserProperties["gCalId"].Value, oEvent.UserProperties["gId"].Value).Execute();
            }
            else
            {
                WriteLog("New event created. Syncing with Google...");
                _ = calendarService.Events.Insert(gEventRequest, gId).Execute();
            }
            Globals.ThisAddIn.Application.ActiveInspector().Close(OlInspectorClose.olDiscard);
            CalendarSync(calendarService, oStore, calendars);
        }

        public void DeleteOEvent(CalendarService calendarService, Store oStore, Dictionary<string, GCalendar> calendars)
        {
            AppointmentItem oEventItem = null;
            if (Globals.ThisAddIn.Application.ActiveInspector() != null)
            {
                oEventItem = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;
            }
            else if (Globals.ThisAddIn.Application.ActiveExplorer() != null)
            {
                if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 1)
                {
                    MessageBox.Show("Non puoi eliminare più di un evento per volta.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                oEventItem = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
            }
            string id = oEventItem.UserProperties["gId"].Value;
            string calId = oEventItem.UserProperties["gCalId"].Value;
            WriteLog("Event with id {0} deleted. Syncing with Google...", id);
            calendarService.Events.Delete(calId, id).Execute();
            CalendarSync(calendarService, oStore, calendars);
        }

        public void CalendarReset(CalendarService calendarService, Store oStore, Dictionary<string, GCalendar> calendars)
        {
            DialogResult result = MessageBox.Show("Vuoi rimuovere tutti i calendari sincronizzati ed eseguire una sincronizzazione completa? Tutti gli eventi non sincronizzati su Google Calendar andranno persi.", "Attenzione", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result == DialogResult.OK)
            {
                Folder oCalMain = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                Folder oCalBin = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
                WriteLog("Removing synced calendars and resyncing");
                foreach (KeyValuePair<string, GCalendar> calendar in calendars)
                {
                    foreach (Folder personalCalendar in oCalMain.Folders)
                    {
                        if (personalCalendar.Name.Contains(calendar.Value.GName))
                        {
                            personalCalendar.Delete();
                        }
                    }
                }
                foreach (KeyValuePair<string, GCalendar> calendar in calendars)
                {
                    foreach (Folder personalCalendar in oCalBin.Folders)
                    {
                        if (personalCalendar.Name.Contains(calendar.Value.GName))
                        {
                            personalCalendar.Delete();
                        }
                    }
                }
                File.Delete("calendarData.json");
                calendars = GetCalendarsDictionary();
                Thread.Sleep(1000);
                CalendarSync(calendarService, oStore, calendars);
                MessageBox.Show("Processo completato.", "Processo completato", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #endregion

        #region Methods for address book sync

        public PeopleServiceService AddressBookInit(UserCredential credential)
        {
            PeopleServiceService addressBookService = new PeopleServiceService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return addressBookService;
        }

        public void AddressBookSync(PeopleServiceService addressBookService, Store oStore)
        {
            WriteLog("Syncing address book");

            Folder oAbMain = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            Folder newAb = null;

            PeopleResource peopleResource = new PeopleResource(addressBookService);
            PeopleResource.ConnectionsResource.ListRequest listRequest;
            ListConnectionsResponse listResponse;

            ContactData tokenData;
            bool isOAbPresent = false;

            foreach (Folder personalAddressBook in oAbMain.Folders)
            {
                if (personalAddressBook.Name.Contains("Google Contacts"))
                {
                    newAb = personalAddressBook;
                    isOAbPresent = true;
                    break;
                }
            }
            if (!isOAbPresent)
            {
                newAb = (Folder)oAbMain.Folders.Add("Google Contacts", OlDefaultFolders.olFolderContacts);
                newAb.UserDefinedProperties.Add("gId", OlUserPropertyType.olText, OlUserPropertyType.olText);
            }

            if (File.Exists("contactData.json"))
            {
                string contactData = File.ReadAllText("contactData.json");
                tokenData = JsonConvert.DeserializeObject<ContactData>(contactData);
                if (DateTime.Compare(tokenData.Expire, DateTime.Now) < 0)
                {
                    tokenData = new ContactData("", DateTime.Now.AddDays(7.0));
                }
            }
            else
            {
                tokenData = new ContactData("", DateTime.Now.AddDays(7.0));
            }

            if (tokenData.Token == "")
            {
                WriteLog("Sync token for address book not present or expired");
                listRequest = peopleResource.Connections.List("people/me");
                listRequest.PersonFields = "addresses,emailAddresses,genders,locales,metadata,names,nicknames,occupations,organizations,phoneNumbers,photos,relations";
                listRequest.PageSize = 1000;
                listRequest.RequestSyncToken = true;
                listResponse = listRequest.Execute();
            }
            else
            {
                WriteLog("Sync token for address book present: {0}", tokenData.Token);
                listRequest = peopleResource.Connections.List("people/me");
                listRequest.PersonFields = "addresses,emailAddresses,genders,locales,metadata,names,nicknames,occupations,organizations,phoneNumbers,photos,relations";
                listRequest.PageSize = 1000;
                listRequest.RequestSyncToken = true;
                listRequest.SyncToken = tokenData.Token;
                try
                {
                    listResponse = listRequest.Execute();
                }
                catch (System.Exception)
                {
                    WriteLog("Sync token for address book expired");
                    listRequest = peopleResource.Connections.List("people/me");
                    listRequest.PersonFields = "addresses,emailAddresses,genders,locales,metadata,names,nicknames,occupations,organizations,phoneNumbers,photos,relations";
                    listRequest.PageSize = 1000;
                    listRequest.RequestSyncToken = true;
                    listResponse = listRequest.Execute();
                }
            }

            ContactsSync(listResponse.Connections, newAb, oStore);

            while (listResponse.NextPageToken != null)
            {
                listRequest = peopleResource.Connections.List("people/me");
                listRequest.PersonFields = "addresses,emailAddresses,genders,locales,metadata,names,nicknames,occupations,organizations,phoneNumbers,photos,relations";
                listRequest.PageSize = 1000;
                listRequest.RequestSyncToken = true;
                listRequest.PageToken = listResponse.NextPageToken;
                ContactsSync(listResponse.Connections, newAb, oStore);
            }

            WriteLog("Writing sync token for address book: {0}", listResponse.NextSyncToken);
            tokenData.Token = listResponse.NextSyncToken;
            tokenData.Expire = DateTime.Now.AddDays(7.0);
            File.WriteAllText("contactData.json", JsonConvert.SerializeObject(tokenData, Newtonsoft.Json.Formatting.Indented));
        }

        private void ContactsSync(IList<Person> gContacts, Folder oAb, Store oStore, ContactItem contactItem = null)
        {
            if (gContacts != null)
            {
                WriteLog("Number of contacts: {0}", gContacts.Count);
                Folder oAbBin = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
                ContactItem oContactItem = null;
                Items restricted;
                string filter;
                bool deleted = false;

                foreach (Person person in gContacts)
                {
                    filter = "[gId] = '" + person.ResourceName + "'";
                    if (contactItem == null)
                    {
                        restricted = oAb.Items.Restrict(filter);
                        oContactItem = (restricted.Count != 0) ? restricted[1] : null; 
                    }
                    if (oContactItem != null)
                    {
                        WriteLog("Contact found");
                        if (person.Metadata.Deleted != null)
                        {
                            deleted = (bool)person.Metadata.Deleted;
                        }

                        if (deleted)
                        {
                            WriteLog("Deleting contact with id {0}", person.ResourceName);
                            oContactItem.Delete();
                            oContactItem = oAbBin.Items.Find(filter);
                            oContactItem?.Delete();
                        }
                        else
                        {
                            WriteLog("Updating contact with id {0}", person.ResourceName);
                            oContactItem = PopulateContactFields(oContactItem, person);
                            oContactItem.Save();
                        }
                    }
                    else
                    {
                        if (person.Metadata.Deleted != null)
                        {
                            deleted = (bool)person.Metadata.Deleted;
                        }
                        if (!deleted)
                        {
                            WriteLog("Adding contact with id {0}", person.ResourceName);
                            oContactItem = contactItem ?? (ContactItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olContactItem);
                            oContactItem.UserProperties.Add("gId", OlUserPropertyType.olText, true, OlUserPropertyType.olText);
                            oContactItem.UserProperties.Add("gETag", OlUserPropertyType.olText, true, OlUserPropertyType.olText);
                            oContactItem.UserProperties["gId"].Value = person.ResourceName;
                            oContactItem = PopulateContactFields(oContactItem, person);
                            if (contactItem == null)
                            {
                                oContactItem.Move(oAb);
                            }
                            else
                            {
                                oContactItem.Save();
                            }
                        }
                    }
                }
            }
            else
            {
                WriteLog("Number of contacts: 0");
            }
        }

        private ContactItem PopulateContactFields(ContactItem oContact, Person person)
        {
            oContact.UserProperties["gETag"].Value = person.ETag;

            oContact.LastName = person.Names != null ? person.Names[0].FamilyName : "";
            oContact.FirstName = person.Names != null ? person.Names[0].GivenName : "";
            oContact.MiddleName = person.Names != null ? person.Names[0].MiddleName : "";
            oContact.Title = person.Names != null ? person.Names[0].HonorificPrefix : "";
            oContact.Suffix = person.Names != null ? person.Names[0].HonorificSuffix : "";

            oContact.CompanyName = person.Organizations != null ? person.Organizations[0].Name : "";
            oContact.Department = person.Organizations != null ? person.Organizations[0].Department : "";
            oContact.OfficeLocation = person.Organizations != null ? person.Organizations[0].Location : "";
            oContact.Profession = person.Occupations != null ? person.Occupations[0].Value : "";

            if (person.EmailAddresses != null)
            {
                bool firstSet = false, secondSet = false, thirdSet = false;
                foreach (EmailAddress emailAddress in person.EmailAddresses)
                {
                    if (!firstSet && !secondSet && !thirdSet)
                    {
                        oContact.Email1Address = emailAddress.Value;
                        oContact.Email1DisplayName = emailAddress.DisplayName;
                        firstSet = true;
                    }
                    else if (firstSet && !secondSet && !thirdSet)
                    {
                        oContact.Email2Address = emailAddress.Value;
                        oContact.Email2DisplayName = emailAddress.DisplayName;
                        secondSet = true;
                    }
                    else if (firstSet && secondSet && !thirdSet)
                    {
                        oContact.Email3Address = emailAddress.Value;
                        oContact.Email3DisplayName = emailAddress.DisplayName;
                        thirdSet = true;
                    }
                    else if (firstSet && secondSet && thirdSet)
                    {
                        break;
                    }
                }
            }

            if (person.Addresses != null)
            {
                foreach (Address address in person.Addresses)
                {
                    switch (address.Type)
                    {
                        case "work":
                            oContact.BusinessAddressCity = address.City;
                            oContact.BusinessAddressCountry = address.Country;
                            oContact.BusinessAddressPostalCode = address.PostalCode;
                            oContact.BusinessAddressStreet = address.StreetAddress;
                            oContact.BusinessAddressPostOfficeBox = address.PoBox;
                            break;

                        case "home":
                            oContact.HomeAddressCity = address.City;
                            oContact.HomeAddressCountry = address.Country;
                            oContact.HomeAddressPostalCode = address.PostalCode;
                            oContact.HomeAddressStreet = address.StreetAddress;
                            oContact.HomeAddressPostOfficeBox = address.PoBox;
                            break;

                        case "other":
                            oContact.OtherAddressCity = address.City;
                            oContact.OtherAddressCountry = address.Country;
                            oContact.OtherAddressPostalCode = address.PostalCode;
                            oContact.OtherAddressStreet = address.StreetAddress;
                            oContact.OtherAddressPostOfficeBox = address.PoBox;
                            break;
                    }
                }
            }

            if (person.PhoneNumbers != null)
            {
                bool homeFirstSet = false, workFirstSet = false;
                foreach (PhoneNumber phoneNumber in person.PhoneNumbers)
                {
                    switch (phoneNumber.Type)
                    {
                        case "home":
                            if (!homeFirstSet)
                            {
                                oContact.HomeTelephoneNumber = phoneNumber.CanonicalForm ?? phoneNumber.Value;
                                homeFirstSet = true;
                            }
                            else
                            {
                                oContact.Home2TelephoneNumber = phoneNumber.CanonicalForm ?? phoneNumber.Value;
                            }
                            break;

                        case "work":
                            if (!workFirstSet)
                            {
                                oContact.BusinessTelephoneNumber = phoneNumber.CanonicalForm ?? phoneNumber.Value;
                                workFirstSet = true;
                            }
                            else
                            {
                                oContact.Business2TelephoneNumber = phoneNumber.CanonicalForm ?? phoneNumber.Value;
                            }
                            break;

                        case "mobile":
                            oContact.MobileTelephoneNumber = phoneNumber.CanonicalForm ?? phoneNumber.Value;
                            break;

                        case "homeFax":
                            oContact.HomeFaxNumber = phoneNumber.CanonicalForm ?? phoneNumber.Value;
                            break;

                        case "workFax":
                            oContact.BusinessFaxNumber = phoneNumber.CanonicalForm ?? phoneNumber.Value;
                            break;

                        case "otherFax":
                            oContact.OtherFaxNumber = phoneNumber.CanonicalForm ?? phoneNumber.Value;
                            break;

                        case "pager":
                            oContact.PagerNumber = phoneNumber.CanonicalForm ?? phoneNumber.Value;
                            break;

                        case "other":
                            oContact.OtherTelephoneNumber = phoneNumber.CanonicalForm ?? phoneNumber.Value;
                            break;

                        case "main":
                            oContact.PrimaryTelephoneNumber = phoneNumber.CanonicalForm ?? phoneNumber.Value;
                            break;
                    }
                }
            }

            if (person.Relations != null)
            {
                foreach (Relation relation in person.Relations)
                {
                    switch (relation.Type)
                    {
                        case "spouse":
                            oContact.Spouse = relation.Person;
                            break;

                        case "assistant":
                            oContact.AssistantName = relation.Person;
                            break;

                        case "manager":
                            oContact.ManagerName = relation.Person;
                            break;
                    }
                }
            }

            oContact.FileAs = person.FileAses != null ? person.FileAses[0].Value : "";

            if (person.Genders != null)
            {
                switch (person.Genders[0].Value)
                {
                    case "male":
                        oContact.Gender = OlGender.olMale;
                        break;

                    case "female":
                        oContact.Gender = OlGender.olFemale;
                        break;

                    case "unspecified":
                        oContact.Gender = OlGender.olUnspecified;
                        break;
                }
            }

            if (person.Interests != null)
            {
                foreach (Interest interest in person.Interests)
                {
                    oContact.Hobby += interest.Value + "; ";
                }
            }

            oContact.NickName = person.Nicknames != null ? person.Nicknames[0].Value : "";

            if (person.Photos != null)
            {
                using (WebClient client = new WebClient())
                {
                    client.DownloadFile(new Uri(person.Photos[0].Url + "?sz=50"), "image.jpg");
                    oContact.AddPicture("image.jpg");
                    File.Delete("image.jpg");
                }
            }
            return oContact;
        }

        public void CreateOContact(PeopleServiceService addressBookService, Store oStore)
        {
            Person gContactRequest = new Person();
            Person gContactResponse;
            ContactItem oContact = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;
            Folder oAb = null;


            gContactRequest.Names = new List<Name>
            {
                new Name
                {
                    FamilyName = oContact.LastName,
                    GivenName = oContact.FirstName,
                    MiddleName = oContact.MiddleName,
                    HonorificPrefix = oContact.Title,
                    HonorificSuffix = oContact.Suffix
                }
            };

            gContactRequest.Organizations = new List<Organization>
            {
                new Organization
                {
                    Name = oContact.CompanyName,
                    Department = oContact.Department,
                    Location = oContact.OfficeLocation
                }
            };

            gContactRequest.Occupations = new List<Occupation> { new Occupation { Value = oContact.Profession } };

            gContactRequest.EmailAddresses = new List<EmailAddress>
            {
                new EmailAddress
                {
                    Value = oContact.Email1Address,
                    DisplayName = oContact.Email1DisplayName
                },
                new EmailAddress
                {
                    Value = oContact.Email2Address,
                    DisplayName = oContact.Email2DisplayName
                },
                new EmailAddress
                {
                    Value = oContact.Email3Address,
                    DisplayName = oContact.Email3DisplayName
                }
            };

            gContactRequest.Addresses = new List<Address>();
            if (oContact.HomeAddressCity != "" || oContact.HomeAddressCountry != "" || oContact.HomeAddressPostalCode != "" || oContact.HomeAddressStreet != "" || oContact.HomeAddressPostOfficeBox != "")
            {
                gContactRequest.Addresses.Add(new Address
                {
                    City = oContact.HomeAddressCity,
                    Country = oContact.HomeAddressCountry,
                    PostalCode = oContact.HomeAddressPostalCode,
                    StreetAddress = oContact.HomeAddressStreet,
                    PoBox = oContact.HomeAddressPostOfficeBox,
                    Type = "home"
                });
            }
            if (oContact.BusinessAddressCity != "" || oContact.BusinessAddressCountry != "" || oContact.BusinessAddressPostalCode != "" || oContact.BusinessAddressStreet != "" || oContact.BusinessAddressPostOfficeBox != "")
            {
                gContactRequest.Addresses.Add(new Address
                {
                    City = oContact.BusinessAddressCity,
                    Country = oContact.BusinessAddressCountry,
                    PostalCode = oContact.BusinessAddressPostalCode,
                    StreetAddress = oContact.BusinessAddressStreet,
                    PoBox = oContact.BusinessAddressPostOfficeBox,
                    Type = "work"
                });
            }
            if (oContact.OtherAddressCity != "" || oContact.OtherAddressCountry != "" || oContact.OtherAddressPostalCode != "" || oContact.OtherAddressStreet != "" || oContact.OtherAddressPostOfficeBox != "")
            {
                gContactRequest.Addresses.Add(new Address
                {
                    City = oContact.OtherAddressCity,
                    Country = oContact.OtherAddressCountry,
                    PostalCode = oContact.OtherAddressPostalCode,
                    StreetAddress = oContact.OtherAddressStreet,
                    PoBox = oContact.OtherAddressPostOfficeBox,
                    Type = "other"
                });
            }

            gContactRequest.PhoneNumbers = new List<PhoneNumber>();
            if (oContact.HomeTelephoneNumber != "")
            {
                gContactRequest.PhoneNumbers.Add(new PhoneNumber
                {
                    Value = oContact.HomeTelephoneNumber,
                    Type = "home"
                });
            }
            if (oContact.Home2TelephoneNumber != "")
            {
                gContactRequest.PhoneNumbers.Add(new PhoneNumber
                {
                    Value = oContact.Home2TelephoneNumber,
                    Type = "home"
                });
            }
            if (oContact.BusinessTelephoneNumber != "")
            {
                gContactRequest.PhoneNumbers.Add(new PhoneNumber
                {
                    Value = oContact.BusinessTelephoneNumber,
                    Type = "work",
                });
            }
            if (oContact.Business2TelephoneNumber != "")
            {
                gContactRequest.PhoneNumbers.Add(new PhoneNumber
                {
                    Value = oContact.Business2TelephoneNumber,
                    Type = "work"
                });
            }
            if (oContact.MobileTelephoneNumber != "")
            {
                gContactRequest.PhoneNumbers.Add(new PhoneNumber
                {
                    Value = oContact.MobileTelephoneNumber,
                    Type = "mobile"
                });
            }
            if (oContact.HomeFaxNumber != "")
            {
                gContactRequest.PhoneNumbers.Add(new PhoneNumber
                {
                    Value = oContact.HomeFaxNumber,
                    Type = "homeFax"
                });
            }
            if (oContact.BusinessFaxNumber != "")
            {
                gContactRequest.PhoneNumbers.Add(new PhoneNumber
                {
                    Value = oContact.BusinessFaxNumber,
                    Type = "workFax"
                });
            }
            if (oContact.OtherFaxNumber != "")
            {
                gContactRequest.PhoneNumbers.Add(new PhoneNumber
                {
                    Value = oContact.OtherFaxNumber,
                    Type = "otherFax"
                });
            }
            if (oContact.PagerNumber != "")
            {
                gContactRequest.PhoneNumbers.Add(new PhoneNumber
                {
                    Value = oContact.PagerNumber,
                    Type = "pager"
                });
            }
            if (oContact.OtherTelephoneNumber != "")
            {
                gContactRequest.PhoneNumbers.Add(new PhoneNumber
                {
                    Value = oContact.OtherTelephoneNumber,
                    Type = "other"
                });
            }
            if (oContact.PrimaryTelephoneNumber != "")
            {
                gContactRequest.PhoneNumbers.Add(new PhoneNumber
                {
                    Value = oContact.PrimaryTelephoneNumber,
                    Type = "main"
                });
            }

            gContactRequest.Relations = new List<Relation>();
            if (oContact.Spouse != "")
            {
                gContactRequest.Relations.Add(new Relation
                {
                    Person = oContact.Spouse,
                    Type = "spouse"
                });
            }
            if (oContact.AssistantName != "")
            {
                gContactRequest.Relations.Add(new Relation
                {
                    Person = oContact.AssistantName,
                    Type = "assistant"
                });
            }
            if (oContact.ManagerName != "")
            {
                gContactRequest.Relations.Add(new Relation
                {
                    Person = oContact.ManagerName,
                    Type = "manager"
                });
            }

            gContactRequest.FileAses = new List<FileAs> { new FileAs { Value = oContact.FileAs } };

            Gender gender = new Gender();
            switch (oContact.Gender)
            {
                case OlGender.olMale:
                    gender.Value = "male";
                    break;

                case OlGender.olFemale:
                    gender.Value = "female";
                    break;

                case OlGender.olUnspecified:
                    gender.Value = "unspecified";
                    break;
            }
            gContactRequest.Genders = new List<Gender> { gender };

            gContactRequest.Interests = new List<Interest> { new Interest { Value = oContact.Hobby } };

            gContactRequest.Nicknames = new List<Nickname> { new Nickname { Value = oContact.NickName } };


            if (oContact.UserProperties.Count != 0)
            {
                WriteLog("Contact with id {0} updated. Syncing with Google...", oContact.UserProperties["gId"].Value);
                gContactRequest.ETag = oContact.UserProperties["gETag"].Value;
                PeopleResource.UpdateContactRequest updateContactRequest = addressBookService.People.UpdateContact(gContactRequest, oContact.UserProperties["gId"].Value);
                updateContactRequest.UpdatePersonFields = "addresses,emailAddresses,genders,locales,names,nicknames,occupations,organizations,phoneNumbers,relations";
                gContactResponse = updateContactRequest.Execute();
            }
            else
            {
                WriteLog("New contact created. Syncing with Google...");
                gContactResponse = addressBookService.People.CreateContact(gContactRequest).Execute();
            }

            Globals.ThisAddIn.Application.ActiveInspector().Close(OlInspectorClose.olSave);

            foreach (Folder personalAddressBook in oStore.GetDefaultFolder(OlDefaultFolders.olFolderContacts).Folders)
            {
                if (personalAddressBook.Name.Contains("Google Contacts"))
                {
                    oAb = personalAddressBook;
                    break;
                }
            }
            List<Person> personWrap = new List<Person>{gContactResponse};

            ContactsSync(personWrap, oAb, oStore, oContact);
        }

        public void DeleteOContact(PeopleServiceService addressBookService, Store oStore)
        {
            ContactItem oContact;
            Folder oAb = null;

            List<string> resourceNames = null;
            BatchDeleteContactsRequest batchDeleteContactsRequest = new BatchDeleteContactsRequest();
            PeopleResource.GetBatchGetRequest batchGetRequest = addressBookService.People.GetBatchGet();
            batchGetRequest.PersonFields = "metadata";

            if (Globals.ThisAddIn.Application.ActiveInspector() != null)
            {
                oContact = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;
                resourceNames = new List<string>{oContact.UserProperties["gId"].Value};
                batchDeleteContactsRequest.ResourceNames = resourceNames;
                batchGetRequest.ResourceNames = resourceNames;
            }
            else if (Globals.ThisAddIn.Application.ActiveExplorer() != null)
            {
                if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count >= 200)
                {
                    MessageBox.Show("Non puoi eliminare più di 200 contatti per volta.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                resourceNames = new List<string>(Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count);
                foreach (ContactItem oContactItem in Globals.ThisAddIn.Application.ActiveExplorer().Selection)
                {
                    resourceNames.Add(oContactItem.UserProperties["gId"].Value);
                }
                batchDeleteContactsRequest.ResourceNames = resourceNames;
                batchGetRequest.ResourceNames = resourceNames;
            }
            WriteLog("{0} contact(s) deleted. Syncing with Google...", resourceNames.Count);
            GetPeopleResponse batchGetResponse = batchGetRequest.Execute();
            addressBookService.People.BatchDeleteContacts(batchDeleteContactsRequest).Execute();

            List<Person> people = new List<Person>();
            foreach (PersonResponse personResponse in batchGetResponse.Responses)
            {
                personResponse.Person.Metadata.Deleted = true;
                people.Add(personResponse.Person);
            }
            foreach (Folder personalAddressBook in oStore.GetDefaultFolder(OlDefaultFolders.olFolderContacts).Folders)
            {
                if (personalAddressBook.Name.Contains("Google Contacts"))
                {
                    oAb = personalAddressBook;
                    break;
                }
            }

            ContactsSync(people, oAb, oStore);
        }

        public void AddressBookReset(PeopleServiceService addressBookService, Store oStore)
        {
            DialogResult result = MessageBox.Show("Vuoi rimuovere tutti i contatti ed eseguire una sincronizzazione completa? Tutti i contatti non sincronizzati con Google andranno persi.", "Attenzione", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result == DialogResult.OK)
            {
                Folder oAbMain = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                Folder oAbBin = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
                WriteLog("Removing synced contacts and resyncing");
                foreach (Folder personalAddressBook in oAbMain.Folders)
                {
                    if (personalAddressBook.Name.Contains("Google Contacts"))
                    {
                        personalAddressBook.Delete();
                    }
                }
                foreach (Folder personalAddressBook in oAbBin.Folders)
                {
                    if (personalAddressBook.Name.Contains("Google Contacts"))
                    {
                        personalAddressBook.Delete();
                    }
                }
                File.Delete("contactData.json");
                Thread.Sleep(1000);
                AddressBookSync(addressBookService, oStore);
                MessageBox.Show("Processo completato.", "Processo completato", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #endregion
    }
}