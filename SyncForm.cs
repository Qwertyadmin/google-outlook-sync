using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace googleSync
{
    public partial class SyncForm : Form
    {
        public SyncForm()
        {
            InitializeComponent();
            string varPath = @"%AppData%\Google Account - Microsoft Outlook Sync";
            string path = Environment.ExpandEnvironmentVariables(varPath);
            Directory.SetCurrentDirectory(path);
            ArrayList comboBoxItems = new ArrayList();
            foreach (KeyValuePair<string, GCalendar> calendar in ThisAddIn.calendars)
            {
                if (calendar.Value.AccessRole != "reader")
                {
                    comboBoxItems.Add(new ComboBoxItem(calendar));
                }
            }
            this.calendarComboBox.DataSource = comboBoxItems;
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            SyncClass syncClass = new SyncClass();
            try
            {
                syncClass.CreateOEvent(ThisAddIn.calendarService, ThisAddIn.oStore, ThisAddIn.calendars, calendarComboBox.SelectedValue.ToString());
            }
            catch (Exception error)
            {
                syncClass.WriteLog("ERROR: " + error);
                MessageBox.Show("Si è verificato un errore. Consulta il log per ulteriori dettagli.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }
    }

    public class ComboBoxItem
    {
        private KeyValuePair<string, GCalendar> item;

        public ComboBoxItem(KeyValuePair<string, GCalendar> item)
        {
            this.item = item;
        }

        public string Name
        {
            get
            {
                return item.Value.GName;
            }
        }

        public string Value
        {
            get
            {
                return item.Key;
            }
        }
    }
}
