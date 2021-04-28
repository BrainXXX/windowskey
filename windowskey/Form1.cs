using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Collections;
using System.Management;


namespace windowskey
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            label2.Text = DecodeWinPkey();
            label8.Text = GetBoardMaker() + ' ' + GetBoardProductId() + ' ' + GetBoardSerial();
            label10.Text = GetCpuManufacturer();
            label12.Text = GetMemManufacturer();
            label14.Text = GetOsStructure();
            label16.Text = GetDrive();

            toolTip1.IsBalloon = true;
            toolTip1.SetToolTip(button1, "Скопировать ключ Windows в буфер обмена");
            toolTip2.IsBalloon = true;
            toolTip2.SetToolTip(button2, "Скопировать ключ BIOS в буфер обмена");
            toolTip3.IsBalloon = true;
            toolTip3.SetToolTip(button3, "Сделать скриншот экрана");

            byte[] buffer = null;
            if (checkMSDM(out buffer))
            {
                Encoding encoding = Encoding.GetEncoding(0x4e4);
                string oemid = encoding.GetString(buffer, 10, 6);
                string dmkey = encoding.GetString(buffer, 56, 29);
                label4.Text=dmkey;
            }
            else
            {
                label4.Text="Ключ OA3.0 в BIOS не был прошит!";
            }

            label5.Text = "*После прошивки ключа в BIOS необходимо перезагрузить компьютер, для его отображения.";

        }

        public static string DecodeWinPkey()
        {
            IList<byte> digitalProductId = null;
            {
                RegistryKey registry = null;
                bool is64 = Environment.Is64BitOperatingSystem;
                if (is64)
                {
                    // 64-bit
                    registry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", false);
                }
                else
                {
                    // 32-bit
                    registry = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", false);
                }
                if (registry != null)
                {
                    // TODO: For other products, key name maybe different.
                    digitalProductId = registry.GetValue("DigitalProductId")
                      as byte[];
                    registry.Close();
                }
                else return null;
            }

            int keyStartIndex = 52;
            int keyEndIndex = keyStartIndex + 15;

            const int numLetters = 24;
            // Possible alpha-numeric characters in product key.
            char[] digits = new[]
            {
                'B', 'C', 'D', 'F', 'G', 'H', 'J', 'K', 'M', 'P', 'Q', 'R',
                'T', 'V', 'W', 'X', 'Y', '2', '3', '4', '6', '7', '8', '9'
            };

            // Check if Windows 8/Office 2013 Style Key (Can contain the letter "N")
            int containsN = (digitalProductId[keyStartIndex + 14] >> 3) & 1;
            digitalProductId[keyStartIndex + 14] = (byte)((digitalProductId[keyStartIndex + 14] & 0xF7) | ((containsN & 2) << 2));

            // Length of decoded product key
            const int decodeLength = 29;

            // Length of decoded product key in byte-form.
            // Each byte represents 2 chars.
            const int decodeStringLength = 15;

            // Array of containing the decoded product key.
            char[] decodedChars = new char[decodeLength];

            // Extract byte 52 to 67 inclusive.
            List<byte> hexPid = new List<byte>();
            for (int i = keyStartIndex; i <= keyEndIndex; i++)
            {
                hexPid.Add(digitalProductId[i]);
            }
            for (int i = decodeLength - 1; i >= 0; i--)
            {
                // Every sixth char is a separator.
                if ((i + 1) % 6 == 0)
                {
                    decodedChars[i] = '-';
                }
                else
                {
                    // Do the actual decoding.
                    int digitMapIndex = 0;
                    for (int j = decodeStringLength - 1; j >= 0; j--)
                    {
                        int byteValue = (digitMapIndex << 8) | hexPid[j];
                        hexPid[j] = (byte)(byteValue / numLetters);
                        digitMapIndex = byteValue % numLetters;
                        decodedChars[i] = digits[digitMapIndex];
                    }
                }
            }
            // Remove first character and put N in the right place
            if (containsN != 0)
            {
                int firstLetterIndex = 0;
                for (int index = 0; index < numLetters; index++)
                {
                    if (decodedChars[0] != digits[index]) continue;
                    firstLetterIndex = index;
                    break;
                }
                string keyWithN = new string(decodedChars);

                keyWithN = keyWithN.Replace("-", string.Empty).Remove(0, 1);
                keyWithN = keyWithN.Substring(0, firstLetterIndex) + "N" + keyWithN.Remove(0, firstLetterIndex);
                keyWithN = keyWithN.Substring(0, 5) + "-" + keyWithN.Substring(5, 5) + "-" + keyWithN.Substring(10, 5) + "-" + keyWithN.Substring(15, 5) + "-" + keyWithN.Substring(20, 5);

                return keyWithN;
            }
            return new string(decodedChars);
        }

        [DllImport("kernel32")]
        private static extern uint EnumSystemFirmwareTables(uint FirmwareTableProviderSignature, IntPtr pFirmwareTableBuffer, uint BufferSize);
        [DllImport("kernel32")]
        private static extern uint GetSystemFirmwareTable(uint FirmwareTableProviderSignature, uint FirmwareTableID, IntPtr pFirmwareTableBuffer, uint BufferSize);

        private static bool checkMSDM(out byte[] buffer)
        {
            uint firmwareTableProviderSignature = 0x41435049; // 'ACPI' in Hexadecimal
            uint bufferSize = EnumSystemFirmwareTables(firmwareTableProviderSignature, IntPtr.Zero, 0);
            IntPtr pFirmwareTableBuffer = Marshal.AllocHGlobal((int)bufferSize);
            buffer = new byte[bufferSize];
            EnumSystemFirmwareTables(firmwareTableProviderSignature, pFirmwareTableBuffer, bufferSize);
            Marshal.Copy(pFirmwareTableBuffer, buffer, 0, buffer.Length);
            Marshal.FreeHGlobal(pFirmwareTableBuffer);
            if (Encoding.ASCII.GetString(buffer).Contains("MSDM"))
            {
                uint firmwareTableID = 0x4d44534d; // Reversed 'MSDM' in Hexadecimal
                bufferSize = GetSystemFirmwareTable(firmwareTableProviderSignature, firmwareTableID, IntPtr.Zero, 0);
                buffer = new byte[bufferSize];
                pFirmwareTableBuffer = Marshal.AllocHGlobal((int)bufferSize);
                GetSystemFirmwareTable(firmwareTableProviderSignature, firmwareTableID, pFirmwareTableBuffer, bufferSize);
                Marshal.Copy(pFirmwareTableBuffer, buffer, 0, buffer.Length);
                Marshal.FreeHGlobal(pFirmwareTableBuffer);
                return true;
            }
            return false;
        }

        public string GetBoardProductId()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BaseBoard");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    return wmi.GetPropertyValue("Product").ToString();
                }
                catch
                {
                }
            }
            return "Product: Unknown";
        }

        public string GetBoardMaker()
        {
            foreach (ManagementObject managementObject in new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BaseBoard").Get())
            {
                try
                {
                    return managementObject.GetPropertyValue("Manufacturer").ToString();
                }
                catch
                {
                }
            }
            return "MB: Unknown";
        }

        public static string GetCpuManufacturer()
        {
            try
            {
                using (ManagementObjectCollection.ManagementObjectEnumerator enumerator = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_Processor").Get().GetEnumerator())
                {
                    if (enumerator.MoveNext())
                        return enumerator.Current["name"].ToString();
                }
            }
            catch
            {
                return (string)null;
            }
            return (string)null;
        }

        public static string GetMemManufacturer()
        {
            try
            {
                using (ManagementObjectCollection.ManagementObjectEnumerator enumerator = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_ComputerSystem").Get().GetEnumerator())
                {
                    if (enumerator.MoveNext())
                        return Convert.ToInt32(Convert.ToDouble(enumerator.Current["TotalPhysicalMemory"]) / 1048576.0).ToString() + " MB";
                }
            }
            catch
            {
                return (string)null;
            }
            return (string)null;
        }

        public static string GetOsStructure()
        {
            try
            {
                using (ManagementObjectCollection.ManagementObjectEnumerator enumerator = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_OperatingSystem").Get().GetEnumerator())
                {
                    if (enumerator.MoveNext())
                    {
                        ManagementObject managementObject = (ManagementObject)enumerator.Current;
                        return managementObject["Caption"].ToString() + " " + managementObject["OSArchitecture"].ToString();
                    }
                }
            }
            catch
            {
                return (string)null;
            }
            return (string)null;
        }

        public string GetBoardSerial()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_BIOS");

            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    return wmi.GetPropertyValue("SerialNumber").ToString();
                }
                catch { }
            }
            return "Unknown Serial";
        }

        public static string GetDrive()
        {
            try
            {
                string hdd = "";
                string hddSerial = "";

                ManagementObjectSearcher mosDisks = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_DiskDrive WHERE InterfaceType = 'IDE' OR InterfaceType = 'SCSI'");
                foreach (ManagementObject moDisk in mosDisks.Get())
                {
                    try
                    {
                        hddSerial = "[" + moDisk["SerialNumber"].ToString().Trim() + "]";
                    }
                    catch
                    {
                        hddSerial = "";
                    }

                    hdd = hdd + moDisk["Model"].ToString() + " "
                        + Math.Round(Convert.ToDouble(moDisk["Size"]) / 1024 / 1024 / 1024, 2).ToString() + " GB " + hddSerial + "\n";
                }
                return hdd.Trim();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Source);
                return (string)null;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(this.label2.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(this.label4.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Bitmap bmp = CaptureScreen();
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "bmp|*.bmp";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                bmp.Save(saveFileDialog1.FileName);
            }
        }

        public static Bitmap CaptureScreen()
        {
            Bitmap BMP = new Bitmap(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width, System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            System.Drawing.Graphics GFX = System.Drawing.Graphics.FromImage(BMP);
            GFX.CopyFromScreen(System.Windows.Forms.Screen.PrimaryScreen.Bounds.X, System.Windows.Forms.Screen.PrimaryScreen.Bounds.Y, 0, 0,
                                System.Windows.Forms.Screen.PrimaryScreen.Bounds.Size, System.Drawing.CopyPixelOperation.SourceCopy);
            return BMP;
        }

    }
}