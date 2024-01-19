using System;
using System.Linq;
using System.Threading;
using System.IO;
using System.Diagnostics;
using System.Net;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Xml;
using Microsoft.Office.Interop.Outlook;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.PeopleService.v1;
using Google.Apis.PeopleService.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace googleSync
{
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
					File.Copy(@"C:\Program Files (x86)\Google Account - Microsoft Outlook Sync\calendarData.xml", "calendarData.xml");
					Globals.ThisAddIn.Application.Session.AddStore("googleSync.pst");
				}
			}
			catch(System.Exception error)
            {
				WriteLog("ERROR: " + error);
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
					GoogleClientSecrets.Load(stream).Secrets,
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

		public CalendarService CalendarInit(UserCredential credential, CalendarService calendarService)
		{
			calendarService = new CalendarService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = ApplicationName,
			});
			return calendarService;
		}

		public void CalendarSync(CalendarService calendarService, Store oStore)
		{
			WriteLog("Syncing calendars");
			XmlDocument calendarData = new XmlDocument();
			calendarData.Load("calendarData.xml");
			XmlNode xmlRoot =  calendarData.FirstChild;
			Folder oCalMain = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
			CalendarListResource.ListRequest gCalRequest = calendarService.CalendarList.List();
			CalendarList gCalList = gCalRequest.Execute();
			int gCalNumber = gCalList.Items.Count();
			int n = 0;
			string[] oCalNames = new string[gCalNumber];

			foreach (CalendarListEntry gCalItem in gCalList.Items)
			{
				bool isOCalPresent = false, isXCalPresent = false;
				WriteLog("Calendar ID: " + gCalItem.Id);

				foreach (Folder personalCalendar in oCalMain.Folders)
				{
					if (personalCalendar.Name.Contains(gCalItem.Summary))
					{
						isOCalPresent = true;
					}
				}
				if (!isOCalPresent)
				{
					oCalMain.Folders.Add(gCalItem.Summary, OlDefaultFolders.olFolderCalendar);
				}
				oCalNames[n] = gCalItem.Summary;
				n++;

				foreach (XmlNode xCalendar in xmlRoot.ChildNodes)
                {
					if (xCalendar.Attributes[1].Value == gCalItem.Id)
                    {
						isXCalPresent = true;
                    }
                }
				if (!isXCalPresent)
				{
					XmlElement xCal = calendarData.CreateElement("calendar");
					XmlAttribute attribute = calendarData.CreateAttribute("id");
					attribute.Value = n.ToString();
					xCal.Attributes.Append(attribute);
					attribute = calendarData.CreateAttribute("gId");
					attribute.Value = gCalItem.Id;
					xCal.Attributes.Append(attribute);
					attribute = calendarData.CreateAttribute("gName");
					attribute.Value = gCalItem.Summary;
					xCal.Attributes.Append(attribute);
					attribute = calendarData.CreateAttribute("accessRole");
					attribute.Value = gCalItem.AccessRole;
					xCal.Attributes.Append(attribute);
					attribute = calendarData.CreateAttribute("token");
					attribute.Value = "";
					xCal.Attributes.Append(attribute);
					xmlRoot.AppendChild(xCal);
				}
			}
			calendarData.Save("calendarData.xml");


			if (xmlRoot.FirstChild.Attributes[4].Value == "")
			{
				WriteLog("Sync tokens not present");
				n = 0;
                foreach (CalendarListEntry gCalItem in gCalList.Items)
                {
                    WriteLog("Syncing calendar n. {0}", n);
                    EventsResource.ListRequest gEventRequest = calendarService.Events.List(gCalItem.Id);
                    gEventRequest.ShowDeleted = false;
                    gEventRequest.SingleEvents = true;
                    Events gEvents = gEventRequest.Execute();

                    foreach (Folder oCal in oCalMain.Folders)
                    {
                        if (oCal.Name.Contains(oCalNames[n]))
                        {
                            EventsFirstSync(gEvents, oCal, gCalItem.Id);
                            break;
                        }
                    }

                    while (gEvents.NextPageToken != null)
                    {
                        gEventRequest = calendarService.Events.List(gCalItem.Id);
                        gEventRequest.ShowDeleted = false;
                        gEventRequest.SingleEvents = true;
                        gEventRequest.PageToken = gEvents.NextPageToken;
                        gEvents = gEventRequest.Execute();
                        foreach (Folder oCal in oCalMain.Folders)
                        {
                            if (oCal.Name.Contains(oCalNames[n]))
                            {
								EventsFirstSync(gEvents, oCal, gCalItem.Id);
                                break;
                            }
                        }
                    }


                    WriteLog("First time writing sync token n.{0}: {1}", n, gEvents.NextSyncToken);
                    xmlRoot.ChildNodes.Item(n).Attributes[4].Value = gEvents.NextSyncToken;
                    n++;
                }
				calendarData.Save("calendarData.xml");
			}

			else
			{
				WriteLog("Sync tokens present");
				string s;
				n = 0;
                foreach (CalendarListEntry gCalItem in gCalList.Items)
                {
                    WriteLog("Syncing calendar n. {0}", n);
                    s = xmlRoot.ChildNodes.Item(n).Attributes[4].Value;
                    WriteLog("Reading sync token n.{0}: {1}", n, s);
                    EventsResource.ListRequest gEventRequest = calendarService.Events.List(gCalItem.Id);
                    gEventRequest.SyncToken = s;
                    gEventRequest.SingleEvents = true;
                    Events gEvents = gEventRequest.Execute();

                    foreach (Folder oCal in oCalMain.Folders)
                    {
                        if (oCal.Name.Contains(oCalNames[n]))
                        {
                            EventsSync(gEvents, oCal, oStore, gCalItem.Id);
                            break;
                        }
                    }

                    while (gEvents.NextPageToken != null)
                    {
                        gEventRequest = calendarService.Events.List(gCalItem.Id);
                        gEventRequest.ShowDeleted = false;
                        gEventRequest.SingleEvents = true;
                        gEventRequest.PageToken = gEvents.NextPageToken;
                        gEvents = gEventRequest.Execute();
                        foreach (Folder oCal in oCalMain.Folders)
                        {
                            if (oCal.Name.Contains(oCalNames[n]))
                            {
                                EventsSync(gEvents, oCal, oStore, gCalItem.Id);
                                break;
                            }
                        }
                    }


                    WriteLog("Writing sync token n.{0}: {1}", n, gEvents.NextSyncToken);
                    xmlRoot.ChildNodes.Item(n).Attributes[4].Value = gEvents.NextSyncToken;
                    n++;
                }
            }
			calendarData.Save("calendarData.xml");
		}

		private void EventsSync(Events gEvents, Folder oCal, Store oStore, string gCalId)
		{
			WriteLog("Number of events: {0}", gEvents.Items.Count());
			Folder oCalBin = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
			foreach (Google.Apis.Calendar.v3.Data.Event gEventItem in gEvents.Items)
			{
				if (gEventItem.Status != "cancelled")
				{
					foreach (AppointmentItem oEventItem in oCal.Items)
					{
						if (oEventItem.UserProperties[1].Value == gEventItem.Id)
						{
							WriteLog("Updating appointment with id " + gEventItem.Id);
							
							oEventItem.Subject = gEventItem.Summary;
							oEventItem.Location = gEventItem.Location;
							oEventItem.Body = gEventItem.Description;
							if (gEventItem.Start.DateTime is DateTime)
							{
								oEventItem.Start = DateTime.Parse(gEventItem.Start.DateTimeRaw);
								oEventItem.End = DateTime.Parse(gEventItem.End.DateTimeRaw);
								oEventItem.AllDayEvent = false;
								oEventItem.ReminderOverrideDefault = false;
							}
							else
							{
								oEventItem.Start = DateTime.Parse(gEventItem.Start.Date);
								oEventItem.End = DateTime.Parse(gEventItem.Start.Date).AddHours(23).AddMinutes(59).AddSeconds(59);
								oEventItem.AllDayEvent = true;
								oEventItem.ReminderSet = false;
							}
							
							oEventItem.Save();
						}
					}
				}
				else if (gEventItem.Status == "cancelled")
				{
					foreach (AppointmentItem oEventItem in oCal.Items)
					{
						if (oEventItem.UserProperties[1].Value == gEventItem.Id)
						{
							WriteLog("Deleting appointment with id " + gEventItem.Id);
							oEventItem.Delete();
							break;
						}
					}
					foreach (AppointmentItem oEventItem in oCalBin.Items)
					{
						if (oEventItem.UserProperties[1].Value == gEventItem.Id)
						{
							oEventItem.Delete();
							break;
						}
					}
				}
			}
		}

		private void EventsFirstSync(Events gEvents, Folder oCal, string gCalId)
		{
			WriteLog("Number of events: {0}", gEvents.Items.Count());
			foreach (Google.Apis.Calendar.v3.Data.Event gEventItem in gEvents.Items)
			{
                AppointmentItem oEvent = oCal.Items.Add(OlItemType.olAppointmentItem) as AppointmentItem;
                WriteLog("Creating appointment with id " + gEventItem.Id);

                oEvent.UserProperties.Add("gId", OlUserPropertyType.olText, false, OlUserPropertyType.olText);
                oEvent.UserProperties.Add("gCalId", OlUserPropertyType.olText, false, OlUserPropertyType.olText);
                oEvent.UserProperties[1].Value = gEventItem.Id;
                oEvent.UserProperties[2].Value = gCalId;
                oEvent.Subject = gEventItem.Summary;
                oEvent.Location = gEventItem.Location;
                oEvent.Body = gEventItem.Description;
                if (gEventItem.Start.DateTime is DateTime)
                {
                    oEvent.Start = DateTime.Parse(gEventItem.Start.DateTimeRaw);
                    oEvent.End = DateTime.Parse(gEventItem.End.DateTimeRaw);
                    oEvent.AllDayEvent = false;
                    oEvent.ReminderOverrideDefault = false;
                }
                else
                {
                    oEvent.Start = DateTime.Parse(gEventItem.Start.Date);
                    oEvent.End = DateTime.Parse(gEventItem.Start.Date).AddHours(23).AddMinutes(59).AddSeconds(59);
                    oEvent.AllDayEvent = true;
                    oEvent.ReminderSet = false;
                }
                
				oEvent.Save();
			}
		}

		public void CreateOEvent(CalendarService calendarService, Store oStore, string gId)
		{
			Google.Apis.Calendar.v3.Data.Event gEventRequest = new Google.Apis.Calendar.v3.Data.Event();
			EventDateTime gEventStart = new EventDateTime();
			EventDateTime gEventEnd = new EventDateTime();
			AppointmentItem oEvent = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;

			gEventRequest.Summary = oEvent.Subject;
			gEventRequest.Description = oEvent.Body;
			gEventRequest.Location = oEvent.Location;
			if (oEvent.AllDayEvent)
			{
				gEventStart.Date = (oEvent.Start.Date.Year + "-" + oEvent.Start.Date.Month + "-" + oEvent.Start.Date.Day);
				gEventRequest.Start = gEventStart;
				gEventEnd.Date = (oEvent.End.Date.Year + "-" + oEvent.End.Date.Month + "-" + oEvent.End.Date.Day);
				gEventRequest.End = gEventEnd;
			}
			else if (!oEvent.AllDayEvent)
			{
				gEventStart.DateTime = oEvent.Start;
				gEventRequest.Start = gEventStart;
				gEventEnd.DateTime = oEvent.End;
				gEventRequest.End = gEventEnd;
			}


			if(oEvent.UserProperties.Count != 0)
            {
				WriteLog("Event with id {0} updated. Syncing with Google...", oEvent.UserProperties[1].Value);
				_ = calendarService.Events.Update(gEventRequest, oEvent.UserProperties[2].Value, oEvent.UserProperties[1].Value).Execute();
            }
			else
            {
				WriteLog("New event created. Syncing with Google...");
				_ = calendarService.Events.Insert(gEventRequest, gId).Execute();
			}
			Globals.ThisAddIn.Application.ActiveInspector().Close(OlInspectorClose.olDiscard);
			CalendarSync(calendarService, oStore);
		}

		public void DeleteOEvent(CalendarService calendarService, Store oStore)
		{
			Folder oCalBin = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
			Items oCalDeleted = oCalBin.Items;
			AppointmentItem oEvent = null;
			if(Globals.ThisAddIn.Application.ActiveInspector() != null)
            {
				oEvent = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;
			}
			else if (Globals.ThisAddIn.Application.ActiveExplorer() != null)
			{
				oEvent = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
			}
			string id = oEvent.UserProperties[1].Value;
			string calId = oEvent.UserProperties[2].Value;
			WriteLog("Deleting appointment with id " + id);
			_ = calendarService.Events.Delete(calId, id).Execute();
			oEvent.Delete();
			foreach (AppointmentItem oEventItem in oCalDeleted)
            {
                if (oEventItem.UserProperties[1].Value == id)
                {
                    oEventItem.Delete();
					break;
                }
            }
            CalendarSync(calendarService, oStore);
		}

		public void CalendarReset(CalendarService calendarService, Store oStore)
		{
			DialogResult result = MessageBox.Show("Vuoi rimuovere tutti i calendari sincronizzati ed eseguire una sincronizzazione completa? Tutti gli eventi non sincronizzati su Google Calendar andranno persi.", "Attenzione", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
			if (result == DialogResult.OK)
			{
				Folder oCalMain = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
				Folder oCalBin = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
				XmlDocument calendarData = new XmlDocument();
				calendarData.Load("calendarData.xml");
				XmlNode xmlRoot = calendarData.FirstChild;
				CalendarListResource.ListRequest gCalRequest = calendarService.CalendarList.List();
				CalendarList gCalList = gCalRequest.Execute();
				WriteLog("Removing synced calendars and resyncing");
				foreach (CalendarListEntry gCalItem in gCalList.Items)
				{
					foreach (Folder personalCalendar in oCalMain.Folders)
					{
						if (personalCalendar.Name.Contains(gCalItem.Summary))
						{
							personalCalendar.Delete();
						}
					}
					foreach (Folder personalCalendar in oCalBin.Folders)
					{
						personalCalendar.Delete();
					}
				}
				foreach (XmlNode xCal in xmlRoot.ChildNodes)
                {
					xCal.Attributes[4].Value = "";
                }
				calendarData.Save("calendarData.xml");
				Thread.Sleep(1000);
				CalendarSync(calendarService, oStore);
				MessageBox.Show("Processo completato.", "Processo completato", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		#endregion

		#region Methods for address book sync

		public PeopleServiceService AddressBookInit(UserCredential credential, PeopleServiceService addressBookService)
		{
			addressBookService = new PeopleServiceService(new BaseClientService.Initializer()
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
			Items oAb = oAbMain.Items;
			PeopleResource peopleResource = new PeopleResource(addressBookService);
			PeopleResource.ConnectionsResource.ListRequest listRequest = peopleResource.Connections.List("people/me");
			listRequest.PersonFields = "addresses,emailAddresses,genders,locales,metadata,names,nicknames,occupations,organizations,phoneNumbers,photos,relations";
			listRequest.PageSize = 1000;
			listRequest.RequestSyncToken = true;

			bool completeSync = false;
			if (File.Exists("addressBookSyncToken.txt"))
            {
				if (TimeSpan.Compare((DateTime.Now - File.GetLastWriteTime("addressBookSyncToken.txt")), TimeSpan.FromDays(6.0)) != -1)
				{
					completeSync = true;
					AddressBookReset(addressBookService, oStore);
                }
            }
            else
            {
				completeSync = true;
            }

			if (completeSync)
            {
				WriteLog("addressBookSyncToken.txt not present");
				using (StreamWriter sw = File.CreateText("addressBookSyncToken.txt"))
                {
					ListConnectionsResponse listResponse = listRequest.Execute();
					WriteLog("Number of contacts: {0}", listResponse.Connections.Count);
					foreach (Person person in listResponse.Connections)
                    {
						ContactSync(person, oStore);
                    }

					while (listResponse.NextPageToken != null)
                    {
						listRequest.PageToken = listResponse.NextPageToken;
						listResponse = listRequest.Execute();
						WriteLog("Number of contacts: {0}", listResponse.Connections.Count);
						foreach (Person person in listResponse.Connections)
						{
							ContactSync(person, oStore);
						}
					}

					WriteLog("First time writing sync token: {0}", listResponse.NextSyncToken);
					sw.WriteLine(listResponse.NextSyncToken);
				}
            }
            else
            {
				WriteLog("addressBookSyncToken.txt present");
				string s;
				using (StreamReader sr = File.OpenText("addressBookSyncToken.txt"))
                {
					while ((s = sr.ReadLine()) != null)
					{
						listRequest.SyncToken = s;
						WriteLog("Reading sync token: {0}", s);
					}
				}
				using (StreamWriter sw = File.CreateText("addressBookSyncToken.txt"))
				{
					ListConnectionsResponse listResponse = listRequest.Execute();
					if (listResponse.Connections != null)
                    {
						WriteLog("Number of contacts: {0}", listResponse.Connections.Count);
						foreach (Person person in listResponse.Connections)
						{
							ContactSync(person, oStore);
						}
					}
                    else
                    {
						WriteLog("Number of contacts: 0");
					}

					while (listResponse.NextPageToken != null)
					{
						listRequest.PageToken = listResponse.NextPageToken;
						listResponse = listRequest.Execute();
						if (listResponse.Connections != null)
						{
							WriteLog("Number of contacts: {0}", listResponse.Connections.Count);
							foreach (Person person in listResponse.Connections)
							{
								ContactSync(person, oStore);
							}
						}
						else
						{
							WriteLog("Number of contacts: 0");
						}
					}

					WriteLog("Writing sync token: {0}", listResponse.NextSyncToken);
					sw.WriteLine(listResponse.NextSyncToken);
				}
			}
		}

		private void ContactSync(Person person, Store oStore)
        {
			Folder oAbMain = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
			Folder oAbBin = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
			Items oAb = oAbMain.Items;
			Items oAbDeleted = oAbBin.Items;
			ContactItem oContact = (ContactItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olContactItem);
			bool deleted = false;
			if (person.Metadata.Deleted != null)
            {
				deleted = (bool)person.Metadata.Deleted;
			}

			if (!deleted)
            {
				bool isUpdate = false;
				foreach (ContactItem oContactItem in oAb)
				{
					if (oContactItem.UserProperties[1].Value == person.ResourceName)
					{
						WriteLog("Updating contact with ID: " + person.ResourceName);
						isUpdate = true;
						ContactItem oContactUpdate = oContactItem;
						oContactUpdate = PopulateContactFields(oContactUpdate, person);
						oContactUpdate.Save();
						break;
					}
				}
				if(!isUpdate)
                {
					WriteLog("Adding contact with ID: " + person.ResourceName);

					oContact.UserProperties.Add("gResourceName", OlUserPropertyType.olText, false, OlUserPropertyType.olText);
					oContact.UserProperties.Add("gETag", OlUserPropertyType.olText, false, OlUserPropertyType.olText);
					oContact.UserProperties[1].Value = person.ResourceName;

					oContact = PopulateContactFields(oContact, person);

					_ = oContact.Move(oAbMain);
				}
			}
            else
            {
				WriteLog("Deleting contact with ID: " + person.ResourceName);
				foreach (ContactItem oContactItem in oAb)
				{
					if (oContactItem.UserProperties[1].Value == person.ResourceName)
					{
						oContactItem.Delete();
						break;
					}
				}
				foreach (ContactItem oContactItem in oAbDeleted)
				{
					if (oContactItem.UserProperties[1].Value == person.ResourceName)
					{
						oContactItem.Delete();
						break;
					}
				}
			}
		}

		private ContactItem PopulateContactFields(ContactItem oContact, Person person)
        {
			oContact.UserProperties[2].Value = person.ETag;

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
								oContact.HomeTelephoneNumber = phoneNumber.CanonicalForm != null ? phoneNumber.CanonicalForm : phoneNumber.Value;
								homeFirstSet = true;
							}
							else
							{
								oContact.Home2TelephoneNumber = phoneNumber.CanonicalForm != null ? phoneNumber.CanonicalForm : phoneNumber.Value;
							}
							break;

						case "work":
							if (!workFirstSet)
							{
								oContact.BusinessTelephoneNumber = phoneNumber.CanonicalForm != null ? phoneNumber.CanonicalForm : phoneNumber.Value;
								workFirstSet = true;
							}
							else
							{
								oContact.Business2TelephoneNumber = phoneNumber.CanonicalForm != null ? phoneNumber.CanonicalForm : phoneNumber.Value;
							}
							break;

						case "mobile":
							oContact.MobileTelephoneNumber = phoneNumber.CanonicalForm != null ? phoneNumber.CanonicalForm : phoneNumber.Value;
							break;

						case "homeFax":
							oContact.HomeFaxNumber = phoneNumber.CanonicalForm != null ? phoneNumber.CanonicalForm : phoneNumber.Value;
							break;

						case "workFax":
							oContact.BusinessFaxNumber = phoneNumber.CanonicalForm != null ? phoneNumber.CanonicalForm : phoneNumber.Value;
							break;

						case "otherFax":
							oContact.OtherFaxNumber = phoneNumber.CanonicalForm != null ? phoneNumber.CanonicalForm : phoneNumber.Value;
							break;

						case "pager":
							oContact.PagerNumber = phoneNumber.CanonicalForm != null ? phoneNumber.CanonicalForm : phoneNumber.Value;
							break;

						case "other":
							oContact.OtherTelephoneNumber = phoneNumber.CanonicalForm != null ? phoneNumber.CanonicalForm : phoneNumber.Value;
							break;

						case "main":
							oContact.PrimaryTelephoneNumber = phoneNumber.CanonicalForm != null ? phoneNumber.CanonicalForm : phoneNumber.Value;
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
			ContactItem oContact = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;



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

			gContactRequest.Occupations = new List<Occupation>{new Occupation{Value = oContact.Profession}};

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

			gContactRequest.FileAses = new List<FileAs>{new FileAs{Value = oContact.FileAs}};

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
			gContactRequest.Genders = new List<Gender>{gender};

			gContactRequest.Interests = new List<Interest>{new Interest{Value = oContact.Hobby}};

			gContactRequest.Nicknames = new List<Nickname>{new Nickname{Value = oContact.NickName}};



			if (oContact.UserProperties.Count != 0)
			{
                WriteLog("Contact with id {0} updated. Syncing with Google...", oContact.UserProperties[1].Value);
                gContactRequest.ETag = oContact.UserProperties[2].Value;
                PeopleResource.UpdateContactRequest updateContactRequest = addressBookService.People.UpdateContact(gContactRequest, oContact.UserProperties[1].Value);
                updateContactRequest.UpdatePersonFields = "addresses,emailAddresses,genders,locales,names,nicknames,occupations,organizations,phoneNumbers,relations";
				_ = updateContactRequest.Execute();
			}
			else
            {
				WriteLog("New contact created. Syncing with Google...");
				_ = addressBookService.People.CreateContact(gContactRequest).Execute();
			}
			Globals.ThisAddIn.Application.ActiveInspector().Close(OlInspectorClose.olDiscard);
			AddressBookSync(addressBookService, oStore);
		}

		public void DeleteOContact(PeopleServiceService addressBookService, Store oStore)
		{
			Folder oAbBin = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
			Items oAbDeleted = oAbBin.Items;
			ContactItem oContact = null;
			if (Globals.ThisAddIn.Application.ActiveInspector() != null)
			{
				oContact = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;
			}
			else if (Globals.ThisAddIn.Application.ActiveExplorer() != null)
			{
				oContact = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
			}
			string resourceName = oContact.UserProperties[1].Value;
			WriteLog("Deleting contact with ID " + resourceName);
			_ = addressBookService.People.DeleteContact(resourceName).Execute();
			oContact.Delete();
			foreach (ContactItem oContactItem in oAbDeleted)
			{
				if (oContactItem.UserProperties[1].Value == resourceName)
				{
					oContactItem.Delete();
					break;
				}
			}
		}

		public void AddressBookReset(PeopleServiceService addressBookService, Store oStore)
        {
			Folder oAbMain = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
			Folder oAbBin = (Folder)oStore.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);
			Items oAb = oAbMain.Items;
			Items oAbDeleted = oAbBin.Items;
			WriteLog("Removing synced contacts and resyncing");
			foreach (ContactItem oContactItem in oAb)
			{
				oContactItem.Delete();
			}
			foreach (ContactItem oContactItem in oAbDeleted)
			{
				oContactItem.Delete();
			}
			File.Delete("addressBookSyncToken.txt");
		}

		#endregion
	}
}