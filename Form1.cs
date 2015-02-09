using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Net;
using System.IO;
using System.Management;
using System.Printing;

using Microsoft.Win32;

namespace ReimagePCInfo
{
    public partial class Form1 : Form
    {
        PrintDocument document = new PrintDocument();
        PrintDialog dialog = new PrintDialog();

        public Form1()
        {
            InitializeComponent();
            document.PrintPage += new PrintPageEventHandler(document_PrintPage);
        }
        void document_PrintPage(object sender, PrintPageEventArgs e)
        {
            e.Graphics.DrawString(textBox1.Text, new Font("Arial", 12, FontStyle.Regular), Brushes.Black, 20, 20);
        }

        private void PopulateInstalledPrintersCombo()
        {
            // Add list of installed printers found to the combo box.
            // The pkInstalledPrinters string will be used to provide the display string.
            
        }
        private PrintDocument printDoc = new PrintDocument();

        public string GetDefaultDomainName()
        {
            const string userRoot = "HKEY_LOCAL_MACHINE";
            const string subkey = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon";
            const string keyName = userRoot + "\\" + subkey;
            string adminValue = Registry.GetValue(keyName, "DefaultDomainName", "Undefined").ToString();
            return adminValue;
        }

        public string GetDefaultUserName()
        {
            const string userRoot = "HKEY_LOCAL_MACHINE";
            const string subkey = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon";
            const string keyName = userRoot + "\\" + subkey;
            string adminValue = Registry.GetValue(keyName, "DefaultUserName", "Undefined").ToString();
            return adminValue;
        }

        public string GetDefaultPassword()
        {
            const string userRoot = "HKEY_LOCAL_MACHINE";
            const string subkey = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon";
            const string keyName = userRoot + "\\" + subkey;
            string adminValue = Registry.GetValue(keyName, "DefaultPassword", "Undefined").ToString();
            return adminValue;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Get IP Address of Pc
            IPAddress[] localIPs = Dns.GetHostAddresses(Dns.GetHostName());
            //String computerIP = "ip";           // have to work on this later
            String domainName = GetDefaultDomainName();
            String defaultUserName = GetDefaultUserName();
            String defaultPassword = GetDefaultPassword();
            String pkInstalledPrinters,
                    printerPort,
                    isDefault,
                    driverName;

            textBox1.AppendText(("Re-Image PC Deployment") + Environment.NewLine);
            textBox1.AppendText("HostName: " + System.Environment.MachineName + Environment.NewLine);
            textBox1.AppendText("Date/Time: " + DateTime.Now.ToString("MM.dd.yyyy hh:mm:ss tt") + Environment.NewLine + Environment.NewLine);

            //textBox1.AppendText("Computer IP Address: " + computerIP + Environment.NewLine);
            //auto logon information
            textBox1.AppendText("Auto Logon Information:" + Environment.NewLine + Environment.NewLine);
            textBox1.AppendText("Default Domain:" + domainName + Environment.NewLine);
            textBox1.AppendText("Default UserName:" + defaultUserName + Environment.NewLine);
            textBox1.AppendText("Default Password:" + defaultPassword + Environment.NewLine + Environment.NewLine);
            //printer information
            textBox1.AppendText("Printer Information:" + Environment.NewLine + Environment.NewLine);
                        
            for (int i = 0; i < PrinterSettings.InstalledPrinters.Count; i++)
            {
                pkInstalledPrinters = PrinterSettings.InstalledPrinters[i];
                String printerName = pkInstalledPrinters.ToString();
                textBox1.AppendText("Printer Name: " + pkInstalledPrinters + Environment.NewLine);

                string query = string.Format("SELECT * from Win32_Printer WHERE Name LIKE '%{0}'", printerName);
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(query);
                ManagementObjectCollection coll = searcher.Get();
                
                foreach (ManagementObject printer in coll)
                {
                    foreach (PropertyData property in printer.Properties)
                    {
                        // IP Address
                        if (property.Name == "PortName")
                        {
                            printerPort = property.Value.ToString();
                            textBox1.AppendText("Port Address: " + printerPort + Environment.NewLine + Environment.NewLine);
                        }
                        // Is Default?
                        if (property.Name == "Default")
                        {
                            isDefault = property.Value.ToString();
                            textBox1.AppendText("Is Default?: " + isDefault + Environment.NewLine);
                        }
                        // Driver used
                        if (property.Name == "DriverName")
                        {
                            driverName = property.Value.ToString();
                            textBox1.AppendText("Driver Name: " + driverName + Environment.NewLine);
                        }
                    }
                }
           
            }
        }
        private void menuStrip2_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dialog.Document = document;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                document.Print();
            }
        }
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Exit
            Application.Exit();
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Save as text
            //Create time stamp
            String timeStamp = DateTime.Now.ToString("MMddyyyy hhmmsstt");

            File.WriteAllText((System.Environment.MachineName + timeStamp + ".txt"), textBox1.Text);
            //Creating popup to tell user that the data is saved
            MessageBox.Show(((System.Environment.MachineName + timeStamp + ".txt") + " is saved in " + (Environment.CurrentDirectory)), "File saved successfully.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
