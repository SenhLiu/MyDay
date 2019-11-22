using Microsoft.Win32;
using System;

namespace MyDay
{
    class RegistryTools
    {

        public static string ReadHKCUString(string subKey, string KeyName)
        {
            try
            {
                // Opening the registry key
                RegistryKey rk = Registry.CurrentUser;

                // Open a subKey as read-only
                RegistryKey sk1 = rk.OpenSubKey(subKey);
                // If the RegistrySubKey doesn't exist -> (null)
                if (sk1 == null)
                    return "";

                return Convert.ToString(sk1.GetValue(KeyName));
            }
            catch (Exception)
            {                
                return "";
            }
        }

        public static bool WriteHKCUString(string subKey, string keyName, string value)
        {
            return WriteHKCUVal(subKey, keyName, value);
        }

        public static bool WriteHKCUVal(string subKey, string keyName, object value)
        {
            try
            {
                // Setting
                RegistryKey rk = Registry.CurrentUser;

                // I have to use CreateSubKey 
                // (create or open it if already exits), 
                // 'cause OpenSubKey open a subKey as read-only
                RegistryKey sk1 = rk.CreateSubKey(subKey);
                // Save the value
                sk1.SetValue(keyName, value);

                return true;
            }
            catch (Exception)
            {                
                return false;
            }
        }
    }
}
