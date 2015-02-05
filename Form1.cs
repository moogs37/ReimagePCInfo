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
        public Form1()
        {
            InitializeComponent();

        }

        private void PopulateInstalledPrintersCombo()
        {
            // Add list of installed printers found to the combo box.
            // The pkInstalledPrinters string will be used to provide the display string.
            
        }
        private PrintDocument printDoc = new PrintDocument();

        private void Form1_Load(object sender, EventArgs e)
        {
            // Get IP Address of Pc
            IPAddress[] localIPs = Dns.GetHostAddresses(Dns.GetHostName());
            String computerIP;

            richTextBox1.AppendText("Re-Image PC Deployment\n\n");
            richTextBox1.AppendText("HostName: " + System.Environment.MachineName + "\n");
            richTextBox1.AppendText("Computer IP Address: " + computerIP + "\n");
            richTextBox1.AppendText("\nPrinter Information:\n\n");

            String pkInstalledPrinters, 
                    printerPort,
                    isDefault,
                    driverName;




            for (int i = 0; i < PrinterSettings.InstalledPrinters.Count; i++)
            {
                pkInstalledPrinters = PrinterSettings.InstalledPrinters[i];
                String printerName = pkInstalledPrinters.ToString();
                richTextBox1.AppendText("Printer Name: " + pkInstalledPrinters + "\n");

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
                            richTextBox1.AppendText("Port Address: " + printerPort + "\n\n");
                        }
                        // Is Default?
                        if (property.Name == "Default")
                        {
                            isDefault = property.Value.ToString();
                            richTextBox1.AppendText("Is Default?: " + isDefault + "\n");
                        }
                        // Driver used
                        if (property.Name == "DriverName")
                        {
                            driverName = property.Value.ToString();
                            richTextBox1.AppendText("Driver Name: " + driverName + "\n");
                        }
                    }
                }
           
            }
        }
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void menuStrip2_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //print on click
            //PrintCommand();

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Exit
            Application.Exit();
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Save as text
            File.WriteAllText((System.Environment.MachineName + ".txt"), richTextBox1.Text);
            //Creating popup to tell user that the data is saved
            MessageBox.Show(((System.Environment.MachineName + ".txt") + " is saved in " + (Environment.CurrentDirectory)), "File saved successfully.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }
}
