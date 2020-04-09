using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OpenMSKey
{
    class Program
    {
        static void Main(string[] args)
        {
            //System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("en-US");
            //System.Threading.Thread.CurrentThread.CurrentCulture = ci;
            //System.Threading.Thread.CurrentThread.CurrentUICulture = ci;

            string thumbprint = "EE50DC0A625FC266A312FA0A10E7171D0AFF152C";
            string hash = "dlV5V3dZNGN3TmdoQVFyeU45NEFpTVhVYW9rZGs5Uk9neGt1L0MwSEt2cz04VFlNUA=="; //length = 68
            string key = "GT8NF-4YBGK-CBJJM-C49WG-8TYMP"; //3273066966471   24004084    vUyWwY4cwNghAQryN94AiMXUaokdk9ROgxku/C0HKvs=8TYMP

            //string openhash = Regex.Match(hash, ".{5}$").Value;
            //string openhash = MDOSCryptography.GetEncryptProductKey(key, "1mic");
            //string openhash = MDOSCryptography.GetProductKey(hash, "1");
            //string openhash = Encoding.ASCII.GetString(Convert.FromBase64String(hash));
            //string openhash = Convert.ToBase64String(Encoding.ASCII.GetBytes(MDOSCryptography.EncryptStringAES("GT8NF-4YBGK-CBJJM-C49WG", "oa30", "mic")+ "8TYMP"));
            //string openhash = Encoding.ASCII.GetString(Convert.FromBase64String(hash));
            string openhash = MDOSCryptography.DecryptStringAES("vUyWwY4cwNghAQryN94AiMXUaokdk9ROgxku/C0HKvs=", thumbprint, thumbprint) 
                + "8TYMP";
            Console.WriteLine(hash);
            Console.WriteLine(key);
            Console.WriteLine(openhash);
            //Console.ReadKey();

            /*for (int i = 0; i < 1000000; i++)
            {
                if (hash == MDOSCryptography.GetEncryptProductKey("GT8NF-4YBGK-CBJJM-C49WG-8TYMP", i.ToString()))
                {
                    Console.WriteLine(i.ToString());
                }
            }*/

            Console.WriteLine("The End");
            Console.ReadKey();
        }

        private static class MDOSCryptography
        {
            private static string salt = "OA30";
            public static string _cryptoThumbprint = string.Empty;

            public static string CryptoThumbprint
            {
                get
                {
                    if (string.IsNullOrEmpty(MDOSCryptography._cryptoThumbprint))
                    {
                        X509Store x509Store = new X509Store(StoreName.My, StoreLocation.LocalMachine);
                        x509Store.Open(OpenFlags.ReadOnly);
                        X509Certificate2Collection certificate2Collection = x509Store.Certificates.Find(X509FindType.FindByThumbprint, 
                            (object)"EE50DC0A625FC266A312FA0A10E7171D0AFF152C", false);
                        if (certificate2Collection != null && certificate2Collection.Count > 0)
                            MDOSCryptography._cryptoThumbprint = certificate2Collection[0].Thumbprint;
                    }
                    //return MDOSCryptography._cryptoThumbprint;
                    return "EE50DC0A625FC266A312FA0A10E7171D0AFF152C";
                }
            }

            public static string GetDecryptedString(string encryptedString)
            {
                string empty = string.Empty;
                return MDOSCryptography.DecryptStringAES(encryptedString, MDOSCryptography.CryptoThumbprint, MDOSCryptography.CryptoThumbprint);
            }

            public static string GetEncryptedString(string stringToEncrypt)
            {
                string empty = string.Empty;
                return MDOSCryptography.EncryptStringAES(stringToEncrypt, MDOSCryptography.CryptoThumbprint, MDOSCryptography.CryptoThumbprint);
            }

            public static string GetProductKey(string key, string thumbprint)
            {
                if (string.IsNullOrEmpty(key) || string.IsNullOrEmpty(thumbprint))
                    return (string)null;
                string str = Encoding.ASCII.GetString(Convert.FromBase64String(key));
                return MDOSCryptography.DecryptStringAES(str.Substring(0, str.Length - 5), thumbprint, thumbprint) + str.Substring(str.Length - 5);
            }

            public static string GetEncryptProductKey(string key, string thumbprint)
            {
                return Convert.ToBase64String(Encoding.ASCII.GetBytes(MDOSCryptography.EncryptStringAES
                    (key.Substring(0, key.Length - 5), thumbprint, thumbprint) + key.Substring(key.Length - 5)));
            }

            public static string DecryptStringAES(string stringToBeDecrypted, string IV, string key)
            {
                MemoryStream memoryStream = new MemoryStream(Convert.FromBase64String(stringToBeDecrypted));
                CryptoStream cryptoStream = (CryptoStream)null;
                PasswordDeriveBytes passwordDeriveBytes1 = (PasswordDeriveBytes)null;
                PasswordDeriveBytes passwordDeriveBytes2 = (PasswordDeriveBytes)null;
                try
                {
                    if (stringToBeDecrypted == null || IV == null || key == null)
                        throw new NullReferenceException();
                    using (AesCryptoServiceProvider cryptoServiceProvider = new AesCryptoServiceProvider())
                    {
                        string str1 = "Microsoft";
                        string str2 = IV.Length < 4 ? IV + str1.Substring(0, 4 - IV.Length).ToLowerInvariant() : IV.Substring(0, 4).ToLowerInvariant();
                        passwordDeriveBytes1 = new PasswordDeriveBytes(key, new byte[13]
                        {
                            (byte) 73,
                            (byte) 118,
                            (byte) 97,
                            (byte) 110,
                            (byte) 32,
                            (byte) 77,
                            (byte) 101,
                            (byte) 100,
                            (byte) 118,
                            (byte) 101,
                            (byte) 100,
                            (byte) 101,
                            (byte) 118
                        });
                        passwordDeriveBytes2 = new PasswordDeriveBytes(MDOSCryptography.salt + str2, new byte[13]
                        {
                            (byte) 73,
                            (byte) 118,
                            (byte) 97,
                            (byte) 110,
                            (byte) 32,
                            (byte) 77,
                            (byte) 101,
                            (byte) 100,
                            (byte) 118,
                            (byte) 97,
                            (byte) 98,
                            (byte) 99,
                            (byte) 100
                        });
                        cryptoStream = new CryptoStream((Stream)memoryStream, 
                            cryptoServiceProvider.CreateDecryptor(passwordDeriveBytes1.GetBytes(32), 
                            passwordDeriveBytes2.GetBytes(16)), CryptoStreamMode.Read);
                        using (StreamReader streamReader = new StreamReader((Stream)cryptoStream))
                            return streamReader.ReadToEnd();
                    }
                }
                finally
                {
                    cryptoStream?.Dispose();
                    passwordDeriveBytes1?.Dispose();
                    passwordDeriveBytes2?.Dispose();
                    memoryStream?.Dispose();
                }
            }

            public static string EncryptStringAES(string stringToBeEncrypted, string IV, string key)
            {
                MemoryStream memoryStream = new MemoryStream();
                CryptoStream cryptoStream = (CryptoStream)null;
                StreamWriter streamWriter = (StreamWriter)null;
                PasswordDeriveBytes passwordDeriveBytes1 = (PasswordDeriveBytes)null;
                PasswordDeriveBytes passwordDeriveBytes2 = (PasswordDeriveBytes)null;
                try
                {
                    if (stringToBeEncrypted == null || IV == null || key == null)
                        throw new NullReferenceException();
                    using (AesCryptoServiceProvider cryptoServiceProvider = new AesCryptoServiceProvider())
                    {
                        string str1 = "Microsoft";
                        string str2 = IV.Length < 4 ? IV + str1.Substring(0, 4 - IV.Length).ToLowerInvariant() : IV.Substring(0, 4).ToLowerInvariant();
                        passwordDeriveBytes1 = new PasswordDeriveBytes(key, new byte[13]
                        {
                            (byte) 73,
                            (byte) 118,
                            (byte) 97,
                            (byte) 110,
                            (byte) 32,
                            (byte) 77,
                            (byte) 101,
                            (byte) 100,
                            (byte) 118,
                            (byte) 101,
                            (byte) 100,
                            (byte) 101,
                            (byte) 118
                        });
                        passwordDeriveBytes2 = new PasswordDeriveBytes(MDOSCryptography.salt + str2, new byte[13]
                        {
                            (byte) 73,
                            (byte) 118,
                            (byte) 97,
                            (byte) 110,
                            (byte) 32,
                            (byte) 77,
                            (byte) 101,
                            (byte) 100,
                            (byte) 118,
                            (byte) 97,
                            (byte) 98,
                            (byte) 99,
                            (byte) 100
                        });
                        cryptoStream = new CryptoStream((Stream)memoryStream, cryptoServiceProvider.CreateEncryptor
                            (passwordDeriveBytes1.GetBytes(32), passwordDeriveBytes2.GetBytes(16)), CryptoStreamMode.Write);
                        streamWriter = new StreamWriter((Stream)cryptoStream);
                        streamWriter.Write(stringToBeEncrypted);
                        streamWriter.Flush();
                        cryptoStream.FlushFinalBlock();
                        return Convert.ToBase64String(memoryStream.ToArray());
                    }
                }
                finally
                {
                    memoryStream?.Dispose();
                    cryptoStream?.Dispose();
                    passwordDeriveBytes1?.Dispose();
                    passwordDeriveBytes2?.Dispose();
                    streamWriter?.Dispose();
                }
            }

            public static string GetDecryptedStringForKPSService(string encryptedString)
            {
                string empty = string.Empty;
                try
                {
                    return MDOSCryptography.DecryptStringAES(encryptedString, MDOSCryptography.CryptoThumbprint, MDOSCryptography.CryptoThumbprint);
                }
                catch
                {
                    return encryptedString;
                }
            }
        }
    }
}