/******************************************************************************/
/*  A/T CONTROL SOFTWARE                                                      */
/*  Copyright (C) 2017 AISIN AW CO.,Ltd.                                 */
/*  Licensed material of AISIN AW CO.,Ltd.                                    */
/*----------------------------------------------------------------------------*/
/*  OBDLink SX Data Retrieval & Parsing Tool                                                     */
/*----------------------------------------------------------------------------*/
/*  $Author:: S.McWilliams                                                  $ */
/*  $Modtime::                                                              $ */
/*  $Revision::                                                             $ */
/******************************************************************************/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using Microsoft.Win32;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace AWTC_OBDSerialPort_01
{

/******************************************************************************************************************************/
/* Main Program Initialization ************************************************************************************************/
/******************************************************************************************************************************/

    public partial class AWTC_OBDLinkSX : Form
    {
        int portOBDUSB = 0;                         //Set variable to track number of OBD USB devices connected to PC
        string portName = string.Empty;             //Declare variable to store name of USB port device is connected to
        bool blinkingImage = false;                 //Boolean to set for flashing image while searching for device
        bool foundOBDLinkSX = false;                //Boolean for whether or not device is found
        bool registryFound = false;                 //Boolean to check if device drivers are installed
        Timer myTimer = new Timer();                //Declare timer to use for blinking image while device is searched for
        SerialPort OBDUSBPort = new SerialPort();   //Declare variable object for port where device is plugged in
        string modeCommandData = string.Empty;      //String to hold data received through OBD port
        string modeCommand = string.Empty;          //String to hold user input command
        private byte _terminator = 0x4;

        public AWTC_OBDLinkSX()
        {
            checkOBDUSBDeviceRegistry();            //Check computer registry to ensure device drivers installed
            if (registryFound)
            {
                InitializeComponent();              //If device drivers found in registry, launch application
                Text = "AWTC_OBDLinkSX - Created by S.McWilliams";            //Set title of application
                Icon = Properties.Resources.AWTC_AW_ICON;   //Set icon of application
            }
            else {
                Environment.Exit(0);                //Close application after message box displayed when drivers are not found
            }
        }



/******************************************************************************************************************************/
/*  Button click functions ****************************************************************************************************/
/******************************************************************************************************************************/

        private void button1_Click(object sender, EventArgs e)
        {
            var win32DeviceClassName = "WIN32_PnPEntity";       //Variable for plug and play device entities

            //Define SQL query to search for USB Serial Port where OBD device is plugged in
            var deviceQuery1 = string.Format("select Name from {0} where Name like '%USB Serial Port (COM%' AND PNPDeviceID like '%VID_0403%'", win32DeviceClassName);

            //Checks if device is not yet found, and starts timer to cause blinking image
            if (foundOBDLinkSX == false)
            {
                myTimer.Enabled = true;
                myTimer.Interval = 100;
                myTimer.Start();
                myTimer.Tick += new EventHandler(BlinkingTimer);
            }

            /*****************************************************************************************/
            /* Process executes using SQL query defined above to find object collection of current   */
            /* plug and play devices that are plugged into USB ports with OBDLink vendor ID          */
            /*****************************************************************************************/
            using (var searcher = new ManagementObjectSearcher(deviceQuery1))               //Sets searcher variable to search through PC Objects using query defined above
            {
                ManagementObjectCollection objectCollection = searcher.Get();               //Gets object collection returned from query

                /* Look at each object returned in object collection returned from query       */
                /* (should only be one object). Using the port name returned from the query,   */
                /* set portName variable equal to port identifier COM## in order to access     */
                /* the port where the device is plugged in to send ELM327 commands to OBD port */
                foreach (ManagementBaseObject managementBaseObject in objectCollection)     
                {
                    foreach (PropertyData propertyData in managementBaseObject.Properties)
                    {
                        portOBDUSB += 1;

                        string fullPortName = propertyData.Value.ToString();

                        //Depending on actual length of port name, select appropriate length of substring
                        if (fullPortName.Length == 22)
                        {
                            portName = fullPortName.Substring(17, 4);
                        }
                        else if (fullPortName.Length == 23) {
                            portName = fullPortName.Substring(17, 5);
                        }
                    }
                }
            }

            //If OBDLink device is not found, alert user to plug device into USB port
            if (portOBDUSB <= 0)
            {
                MessageBox.Show("OBD USB Device not found. Please plug device in and try again");
            }
            //If OBDLink device is found, enable button to connect to device
            else if (portOBDUSB == 1)
            {
                myTimer.Stop();
                foundOBDLinkSX = true;
                blinkingImage = false;
                portOBDUSB = 0;

                try
                {
                    OpenOBDUSBPort();
                    button5.Enabled = true;
                }
                catch
                {
                    MessageBox.Show("Unable to open port");
                    Environment.Exit(0);
                }
            }
            //If more than one device is detected, alert user to unplug additional devices
            else
            {
                MessageBox.Show("Please make sure only one OBD USB device is connected to the computer and try again");
                portOBDUSB = 0;
            }
        }

        //Function that reads user input command and writes to OBD device
        private void button3_Click(object sender, EventArgs e)
        {
            OBDUSBPort.DiscardInBuffer();
            OBDUSBPort.DiscardOutBuffer();
            modeCommandData = string.Empty;
            modeCommand = string.Empty;
            modeCommand = textBox1.Text;

            if (string.IsNullOrEmpty(modeCommand) || string.IsNullOrWhiteSpace(modeCommand)) {
                MessageBox.Show("Please enter a valid command");
            } else
            {
                OBDUSBPort.Write(modeCommand + "\r");
                label3.Text = "Sent: " + modeCommand;

                if (!label4.Enabled && !label4.Visible)
                {
                    label4.Enabled = true;
                    label4.Visible = true;
                }

                while (string.IsNullOrEmpty(modeCommandData) && !modeCommandData.Contains(modeCommand))
                {
                    modeCommandData = OBDUSBPort.ReadExisting();
                }

                label4.Text = modeCommandData;
            }
            
            //button4.Enabled = true;
        }

        //Function that exports received data to Excel file
        private void button4_Click(object sender, EventArgs e)
        {
            
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            modeCommandData = "First excel file test was good";

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed. Please be sure Excel is installed on this PC.");
                return;
            }

            xlApp.Visible = true;

            Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = (Worksheet)wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            // Select the Excel cells, in the range c1 to c7 in the worksheet.
            Range aRange = ws.get_Range("C1", "C7");

            if (aRange == null)
            {
                Console.WriteLine("Could not get a range. Check to be sure you have the correct versions of the office DLLs.");
            }

            char[] delimiterChars = { ' ', ',', '.', ':', '\t' };

            string[] dataReceived = modeCommandData.Split(delimiterChars);

            // Fill the cells in the C1 to C7 range of the worksheet with the number 6.
            /*for (int i = 0; i <= dataReceived.Length; i++)
            {

            }*/

            aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, dataReceived);

            /*if (!string.IsNullOrEmpty(modeCommandData) || !string.IsNullOrWhiteSpace(modeCommandData))
            {
                
            }*/

            button4.Enabled = false;
        }


        private void button5_Click(object sender, EventArgs e)
        {
            pictureBox2.Image = Properties.Resources.Red_Circle_XS;

            if (button3.Enabled)
            {
                button3.Enabled = false;
            }

            label3.Text = string.Empty;
            label4.Text = string.Empty;
            modeCommandData = String.Empty;
            modeCommand = string.Empty;
            textBox1.Text = string.Empty;

            if (OBDUSBPort.IsOpen)
            {
                OBDUSBPort.DiscardInBuffer();
                OBDUSBPort.DiscardOutBuffer();
                OBDUSBPort.Close();
            }
        }



/******************************************************************************************************************************/
/*  Functions initiated by button presses   ***********************************************************************************/
/******************************************************************************************************************************/

        //Function that checks to see if device from OBDLink vendor exists in Windows registry
        //If registry exists, check the driver value to ensure that the drivers have been installed
        private void checkOBDUSBDeviceRegistry()
        {
            string regKeyPath = "SYSTEM\\CurrentControlSet\\Enum\\USB\\VID_0403&PID_6015\\";

            RegistryKey topRegKey = Registry.LocalMachine.OpenSubKey("SYSTEM");
            RegistryKey regKey1 = topRegKey.OpenSubKey("CurrentControlSet");
            RegistryKey regKey2 = regKey1.OpenSubKey("Enum");
            RegistryKey regKey3 = regKey2.OpenSubKey("USB");
            string[] subKeyNames = regKey3.GetSubKeyNames();
            foreach (string subKey in subKeyNames) {
                if (subKey.Contains("VID_0403"))
                {
                    RegistryKey regKey4 = regKey3.OpenSubKey(subKey);
                    string[] usbRegKeys = regKey4.GetSubKeyNames();

                    if (usbRegKeys.Length == 1 && usbRegKeys != null)
                    {
                        regKeyPath += usbRegKeys[0];

                        string userRoot = "HKEY_LOCAL_MACHINE";
                        string keyName = userRoot + "\\" + regKeyPath;

                        var keyValue = Registry.GetValue(keyName, "Driver", null);

                        if (keyValue != null)
                        {
                            registryFound = true;
                            return;
                        }
                    }
                }
            }

            //If drivers are not found, run OBDLink SX Driver installer executable
            if (registryFound == false)
            {
                Process.Start(Properties.Resources.OBDLink_SX_Driver_Installer.ToString());
                return;
            }
        }

        //Function that alternates circle image to blink while searching for device or performing other processes
        private void BlinkingTimer(Object myObject, EventArgs myEventArgs)
        {
            myTimer.Stop();

            if (blinkingImage)
            {
                pictureBox2.BackgroundImage = Properties.Resources.Yellow_Circle_XS;
                blinkingImage = false;
                myTimer.Enabled = true;
            }
            else
            {
                pictureBox2.BackgroundImage = Properties.Resources.Green_Circle_XS;
                blinkingImage = true;
                myTimer.Enabled = true;
            }
        }

        //Function to set port properties and open corresponding port to allow OBD communication
        private void OpenOBDUSBPort() {
            string openData = string.Empty;
            OBDUSBPort.PortName = portName;
            OBDUSBPort.BaudRate = 115200;
            OBDUSBPort.StopBits = StopBits.One;
            OBDUSBPort.DataBits = 8;
            OBDUSBPort.Parity = Parity.None;
            OBDUSBPort.Handshake = Handshake.None;
            OBDUSBPort.RtsEnable = false;

            if (!OBDUSBPort.IsOpen) {
                OBDUSBPort.Open();
                OBDUSBPort.DiscardInBuffer();
                OBDUSBPort.DiscardOutBuffer();
                pictureBox2.Image = Properties.Resources.Green_Circle_XS;
                button3.Enabled = true;
                textBox1.Enabled = true;
                AcceptButton = button3;
            }
                        
            return;
        }
    }
}
