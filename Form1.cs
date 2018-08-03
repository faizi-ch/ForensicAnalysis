using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ForensicAnalysis
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string registry_key = "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RecentDocs";
            using (Microsoft.Win32.RegistryKey key = Registry.CurrentUser.OpenSubKey(registry_key))
            {
                foreach (string subkey_name in key.GetSubKeyNames())
                {
                    using (RegistryKey subkey = key.OpenSubKey(subkey_name))
                    {
                        string[] ss=subkey.Name.Split('\\');
                        //Console.WriteLine(ss[7]);
                        if (!ss[7].Contains("com") || !subkey_name.Contains("0"))
                        {
                            //Console.WriteLine(BytesToStringConverted(GetRecentlyOpenedFile(ss[7])));
                            //listBox1.Items.Add(GetRecentlyOpenedFile(ss[7]));
                            string s = Encoding.Unicode.GetString(GetRecentlyOpenedFile(ss[7]));
                            Encoding en = Enc.GetTextEncoding(GetRecentlyOpenedFile(ss[7]));

                            
                                richTextBox1.Text += "\n" + en.GetString(GetRecentlyOpenedFile(ss[7]));
                            
                            
                        }
                        
                        //using (RegistryKey sbkey_name = key.OpenSubKey(subkey.Name))
                        //{
                        //using (string exname = subkey.Name)
                        //{
                        //Console.WriteLine(sbkey_name.Name);
                        //}
                        //}

                    }
                }
            }
            //Console.WriteLine(GetRecentlyOpenedFile(".mp3"));
            //listBox1.Items.Add(BytesToStringConverted(GetRecentlyOpenedFile(".mp3")));
            //richTextBox1.Text = GetRecentlyOpenedFile(".mp3");
        }

        

        public byte[] GetRecentlyOpenedFile(string extention)
        {
            RegistryKey regKey = Registry.CurrentUser;
            byte[] recentFile = null;
            regKey = regKey.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RecentDocs");

            if (string.IsNullOrEmpty(extention))
                extention = ".docs";

            RegistryKey myKey = regKey.OpenSubKey(extention);

            if (myKey == null && regKey.GetSubKeyNames().Length > 0)
                myKey = regKey.OpenSubKey(
                               regKey.GetSubKeyNames()[regKey.GetSubKeyNames().Length - 2]);
            if (myKey != null)
            {
                string[] names = myKey.GetValueNames();
                if (names != null && names.Length > 0)
                {
                    recentFile = (byte[])myKey.GetValue(names[names.Length - 2]);

                }

            }

            return recentFile;
        }
        static string BytesToStringConverted(byte[] bytes)
        {
            using (var stream = new MemoryStream(bytes))
            {
                using (var streamReader = new StreamReader(stream))
                {
                    return streamReader.ReadToEnd();
                }
            }
        }
    }

    static class Enc
    {
        public static Encoding GetTextEncoding(byte[] bytes, int offset = 0, int? length = null)
        {
            if (bytes == null)
            {
                throw new ArgumentNullException("bytes");
            }
            length = length ?? bytes.Length;
            if (offset < 0 || offset > bytes.Length)
            {
                throw new ArgumentOutOfRangeException("offset", "Offset is out of range.");
            }
            if (length < 0 || length > bytes.Length)
            {
                throw new ArgumentOutOfRangeException("length", "Length is out of range.");
            }
            else if ((offset + length) > bytes.Length)
            {
                throw new ArgumentOutOfRangeException("offset", "The specified range is outside of the specified buffer.");
            }
            // Look for a byte order mark:
            if (length >= 4)
            {
                var one = bytes[offset];
                var two = bytes[offset + 1];
                var three = bytes[offset + 2];
                var four = bytes[offset + 3];
                if (one == 0x2B &&
                    two == 0x2F &&
                    three == 0x76 &&
                    (four == 0x38 || four == 0x39 || four == 0x2B || four == 0x2F))
                {
                    return Encoding.UTF7;
                }
                else if (one == 0xFE && two == 0xFF && three == 0x00 && four == 0x00)
                {
                    return Encoding.UTF32;
                }
                else if (four == 0xFE && three == 0xFF && two == 0x00 && one == 0x00)
                {
                    throw new NotSupportedException("The byte order mark specifies UTF-32 in big endian order, which is not supported by .NET.");
                }
            }
            else if (length >= 3)
            {
                var one = bytes[offset];
                var two = bytes[offset + 1];
                var three = bytes[offset + 2];
                if (one == 0xFF && two == 0xFE)
                {
                    return Encoding.Unicode;
                }
                else if (one == 0xFE && two == 0xFF)
                {
                    return Encoding.BigEndianUnicode;
                }
                else if (one == 0xEF && two == 0xBB && three == 0xBF)
                {
                    return Encoding.UTF8;
                }
            }
            if (length > 1)
            {
                // Look for a leading < sign:
                if (bytes[offset] == 0x3C)
                {
                    if (bytes[offset + 1] == 0x00)
                    {
                        return Encoding.Unicode;
                    }
                    else
                    {
                        return Encoding.UTF8;
                    }

                }
                else if (bytes[offset] == 0x00 && bytes[offset + 1] == 0x3C)
                {
                    return Encoding.BigEndianUnicode;
                }
            }
            if (bytes.IsUtf8())
            {
                return Encoding.UTF8;
            }
            else
            {
                // Impossible to tell.
                return null;
            }
        }
        public static bool IsUtf8(this byte[] bytes, int offset = 0, int? length = null)
        {
            if (bytes == null)
            {
                throw new ArgumentNullException("bytes");
            }
            length = length ?? (bytes.Length - offset);
            if (offset < 0 || offset > bytes.Length)
            {
                throw new ArgumentOutOfRangeException("offset", "Offset is out of range.");
            }
            else if (length < 0)
            {
                throw new ArgumentOutOfRangeException("length");
            }
            else if ((offset + length) > bytes.Length)
            {
                throw new ArgumentOutOfRangeException("offset", "The specified range is outside of the specified buffer.");
            }
            var bytesRemaining = length.Value;
            while (bytesRemaining > 0)
            {
                var rank = bytes.GetUtf8MultibyteRank(offset, Math.Min(4, bytesRemaining));
                if (rank == MultibyteRank.None)
                {
                    return false;
                }
                else
                {
                    var charsRead = (int)rank;
                    offset += charsRead;
                    bytesRemaining -= charsRead;
                }
            }
            return true;
        }
        public static MultibyteRank GetUtf8MultibyteRank(this byte[] bytes, int offset = 0, int length = 4)
        {
            if (bytes == null)
            {
                throw new ArgumentNullException("bytes");
            }
            if (offset < 0 || offset > bytes.Length)
            {
                throw new ArgumentOutOfRangeException("offset", "Offset is out of range.");
            }
            else if (length < 0 || length > 4)
            {
                throw new ArgumentOutOfRangeException("length", "Only values 1-4 are valid.");
            }
            else if ((offset + length) > bytes.Length)
            {
                throw new ArgumentOutOfRangeException("offset", "The specified range is outside of the specified buffer.");
            }
            // Possible 4 byte sequence
            if (length > 3 && IsLead4(bytes[offset]))
            {
                if (IsExtendedByte(bytes[offset + 1]) && IsExtendedByte(bytes[offset + 2]) && IsExtendedByte(bytes[offset + 3]))
                {
                    return MultibyteRank.Four;
                }
            }
            // Possible 3 byte sequence
            else if (length > 2 && IsLead3(bytes[offset]))
            {
                if (IsExtendedByte(bytes[offset + 1]) && IsExtendedByte(bytes[offset + 2]))
                {
                    return MultibyteRank.Three;
                }
            }
            // Possible 2 byte sequence
            else if (length > 1 && IsLead2(bytes[offset]) && IsExtendedByte(bytes[offset + 1]))
            {
                return MultibyteRank.Two;
            }
            if (bytes[offset] < 0x80)
            {
                return MultibyteRank.One;
            }
            else
            {
                return MultibyteRank.None;
            }
        }
        private static bool IsLead4(byte b)
        {
            return b >= 0xF0 && b < 0xF8;
        }
        private static bool IsLead3(byte b)
        {
            return b >= 0xE0 && b < 0xF0;
        }
        private static bool IsLead2(byte b)
        {
            return b >= 0xC0 && b < 0xE0;
        }
        private static bool IsExtendedByte(byte b)
        {
            return b > 0x80 && b < 0xC0;
        }


        public enum MultibyteRank
        {
            None = 0,
            One = 1,
            Two = 2,
            Three = 3,
            Four = 4
        }
    }
}
