using System.Collections.Generic;
using Microsoft.Win32;

namespace ExcelToCsv.Helpers
{
    public class RegistryManager
    {
        private const string ApplicationName = "ExcelToCsv";
        private const string QuickDirectories = "QuickDirectories";

        public static void SaveQuickDirectories(IEnumerable<string> quickDirectories)
        {
            var regPath = GetRegPath(ApplicationName);
            var key = Registry.CurrentUser.CreateSubKey(regPath);

            key?.SetValue(QuickDirectories, string.Join(";", quickDirectories));
        }

        public static IEnumerable<string> LoadQuickDirectories()
        {
            var regPath = GetRegPath(ApplicationName);
            var key = Registry.CurrentUser.OpenSubKey(regPath);

            if (key == null)
                return new List<string>();

            return key.GetValue(QuickDirectories).ToString().Split(';');
        }

        private static string GetRegPath(string applicationName)
        {
            return $@"Software\{applicationName}\WindowBounds\";
        }
    }
}