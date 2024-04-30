using Microsoft.Office.Interop.Outlook;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.PeopleService.v1;
using System.Collections.Generic;
using System.Windows.Forms;


namespace googleSync
{

    public class GCalendar
    {
        public GCalendar(string gName, string accessRole, string token, int[] reminders)
        {
            this.GName = gName;
            this.AccessRole = accessRole;
            this.Token = token;
            this.Reminders = reminders;
        }

        public string GName { get; set; }
        public string AccessRole { get; set; }
        public string Token { get; set; }
		public int[] Reminders { get; set; }
    }

    public partial class ThisAddIn
	{
		UserCredential credential;
		public static CalendarService calendarService;
		public static Store oStore;
		public static PeopleServiceService addressBookService;
        public static Dictionary<string, GCalendar> calendars;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new SyncRibbon();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			SyncClass syncClass = new SyncClass();
			syncClass.WriteLog("Starting Outlook and loading the add-in");
			try
			{
				oStore = syncClass.GetOStore();
				credential = syncClass.Login();
				calendars = syncClass.GetCalendarsDictionary();
				calendarService = syncClass.CalendarInit(credential);
				addressBookService = syncClass.AddressBookInit(credential);
				syncClass.CalendarSync(calendarService, oStore, calendars);
				syncClass.AddressBookSync(addressBookService, oStore);
            }
            catch (System.Exception error)
            {
				syncClass.WriteLog("ERROR: " + error);
                MessageBox.Show("Si è verificato un errore. Consulta il log per ulteriori dettagli.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
			// Nota: Outlook non genera più questo evento. Se è presente codice che 
			//    deve essere eseguito all'arresto di Outlook, vedere https://go.microsoft.com/fwlink/?LinkId=506785
		}

		#region Codice generato da VSTO

		/// <summary>
		/// Metodo richiesto per il supporto della finestra di progettazione. Non modificare
		/// il contenuto del metodo con l'editor di codice.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}
		
		#endregion
	}
}