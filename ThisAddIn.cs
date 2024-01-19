using Microsoft.Office.Interop.Outlook;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.PeopleService.v1;


namespace googleSync
{
	public partial class ThisAddIn
	{
		UserCredential credential;
		public static CalendarService calendarService;
		public static Store oStore;
		public static PeopleServiceService addressBookService;

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
				calendarService = syncClass.CalendarInit(credential, calendarService);
				addressBookService = syncClass.AddressBookInit(credential, addressBookService);
				syncClass.CalendarSync(calendarService, oStore);
				syncClass.AddressBookSync(addressBookService, oStore);
            }
            catch (System.Exception error)
            {
				syncClass.WriteLog("ERROR: " + error);
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