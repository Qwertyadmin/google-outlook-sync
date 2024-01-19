using System;
using System.IO;
using System.Windows.Forms;
using System.Xml;

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
            XmlDocument calendarData = new XmlDocument();
            calendarData.Load("calendarData.xml");
            XmlNode xmlRoot = calendarData.FirstChild;
            foreach (XmlNode xCal in xmlRoot)
            {
                if (xCal.Attributes[3].Value != "reader")
                {
                    calendarComboBox.Items.Add(xCal.Attributes[2].Value);
                }
            }
            calendarComboBox.Text = calendarComboBox.Items[0].ToString();
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            SyncClass syncClass = new SyncClass();
            XmlDocument calendarData = new XmlDocument();
            calendarData.Load("calendarData.xml");
            XmlNode xmlRoot = calendarData.FirstChild;
            foreach (XmlNode xCal in xmlRoot)
            {
                if (xCal.Attributes[2].Value == calendarComboBox.SelectedItem.ToString())
                {
                    try
                    {
                        syncClass.CreateOEvent(ThisAddIn.calendarService, ThisAddIn.oStore, xCal.Attributes[1].Value);
                    }
                    catch (Exception error)
                    {
                        syncClass.WriteLog("ERROR: " + error);
                    }
                }
            }
            Close();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
