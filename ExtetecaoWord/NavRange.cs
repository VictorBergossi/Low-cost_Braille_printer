using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

using System.Management;
using System.IO.Ports;

namespace ExtetecaoWord
{
    public partial class NavRange
    {
        private void NavRange_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        private void PrintCaecus_Click(object sender, RibbonControlEventArgs e)
        {
            Send();
        }

        private string AutodetectArduinoPort()
        {
            ManagementScope connectionScope = new ManagementScope();
            SelectQuery serialQuery = new SelectQuery("SELECT * FROM Win32_SerialPort");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(connectionScope, serialQuery);

            try
            {
                foreach (ManagementObject item in searcher.Get())
                {
                    string desc = item["Description"].ToString();
                    string deviceId = item["DeviceID"].ToString();

                    if (desc.Contains("Arduino"))
                    {
                        return deviceId;
                    }
                }
            }
            catch (ManagementException e)
            {
                /* Do Nothing */
            }

            return null;
        }

        private void Send()
        {
            string portName = AutodetectArduinoPort();
            if (portName == null)
            {
                MessageBox.Show("Caecus não encontrada!");
                return;
            }
            SerialPort serialPort = new SerialPort();
            serialPort.PortName = portName;

            Braile braile = new Braile();
            int[] d = new int[] { 97 };
            byte[] buf = braile.getSerializationToArduino();
            int count = sizeof(byte) * buf.Length;

            serialPort.Open();
            serialPort.Write(buf, 0, count);

            MessageBox.Show("Enviado!");

            serialPort.Close();
        }
    }

}
