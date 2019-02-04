using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace ExcelToCsv.Helpers
{
    public class FileManager
    {
        public static string CreateTemplateFile(string file)
        {
            var templatePath = Environment.GetFolderPath(Environment.SpecialFolder.Templates);
            var fileName = Path.GetFileName(file);
            var templateFile = Path.Combine(templatePath, fileName ?? throw new InvalidOperationException());

            if (File.Exists(templateFile) &&
                new FileInfo(templateFile).IsReadOnly)
                new FileInfo(templateFile).IsReadOnly = false;

            File.Copy(file, templateFile, overwrite: true);
            return templateFile;
        }

        public static bool GetAccessFile(string file, bool showMessage = true)
        {
            try
            {
                var fileInfo = new FileInfo(file);

                if (fileInfo.IsReadOnly)
                {
                    if (showMessage)
                        MessageBox.Show($"\"{file}\"\n{Resources.Dictionary.FileIsReadOnly}",
                            Resources.Dictionary.FileNotEditable, MessageBoxButton.OK, MessageBoxImage.Warning);

                    return false;
                }

                using (new FileStream(file, FileMode.Open, FileAccess.ReadWrite))
                {
                    return true;
                }
            }
            catch (Exception e)
            {
                if (showMessage)
                    MessageBox.Show(e.Message, Resources.Dictionary.FileNotEditable, MessageBoxButton.OK, MessageBoxImage.Warning);

                return false;
            }
        }

        public static bool GetAccessDirectory(string directory, bool showMessage = true)
        {
            var files = GetAllFiles(directory);
            var filesInUse = files.Where(n => !GetAccessFile(n, showMessage: false)).ToList();

            if (!filesInUse.Any() ||
                !showMessage)
                return true;

            MessageBox.Show(
                $"{string.Join("\n", filesInUse)}",
                Resources.Dictionary.FilesNotEditable, MessageBoxButton.OK, MessageBoxImage.Warning);

            return false;

        }

        public static List<string> GetAllFiles(string directory)
        {
            if (string.IsNullOrEmpty(directory) ||
                !Directory.Exists(directory))
                return new List<string>();

            var allFiles = Directory.GetFiles(directory).ToList();
            var directories = Directory.GetDirectories(directory).ToList();

            while (directories.Any())
            {
                var currentDirectory = directories.ElementAt(0);
                directories.RemoveAt(0);
                directories.AddRange(Directory.GetDirectories(currentDirectory));
                allFiles.AddRange(Directory.GetFiles(currentDirectory));
            }

            return allFiles;
        }
    }
}