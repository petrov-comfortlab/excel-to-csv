using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using WixSharp;
using WixSharp.Forms;

namespace Installer
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                Compiler.WixLocation = Path.Combine(SolutionDirectoryManager.GetSolutionDirectoryInfo().FullName, @"packages\WiX.3.11.1\tools");

                AutoElements.SupportEmptyDirectories = CompilerSupportState.Enabled;

                var projects = new List<Project>
                {
                    ExcelToCsvProject(
                        guid: new Guid("B6731803-190B-4FA7-B654-0F929DFF565F"),
                        upgradeGuid: new Guid("CC92BCBD-CEB7-4D6A-B494-2B6A1D4DE78F"),
                        productGuid: new Guid("318E1848-A260-4F06-BC53-B0B1A6BFF428")), //TODO
                };

                projects.ForEach(project => Compiler.BuildMsi(project));

                Console.WriteLine("Complete! Press any key to exit...");
                Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine($"{e.Message}");
                Console.ReadLine();
            }
        }

        private static readonly string SourceFiles = Path.Combine(SolutionDirectoryManager.GetProjectDirectoryInfo().FullName, @"Source\");
        private static readonly string InstallerSourceFiles = Path.Combine(SolutionDirectoryManager.GetProjectDirectoryInfo().FullName, @"InstallerSource\");

        private static Project ExcelToCsvProject(Guid guid, Guid upgradeGuid, Guid productGuid)
        {
            var fileExe = Directory.GetFiles(SourceFiles).FirstOrDefault(n => Path.GetExtension(n).ToLower() == ".exe");

            if (string.IsNullOrEmpty(fileExe))
            {
                Console.WriteLine($"Source *.exe file not found");
                Console.ReadLine();
                return null;
            }

            var fileVersionInfo = FileVersionInfo.GetVersionInfo(fileExe);
            var version = new Version(fileVersionInfo.ProductVersion);
            var appName = fileVersionInfo.ProductName;
            var targetPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), fileVersionInfo.CompanyName, appName);

            var project = new ManagedProject(appName,
                new Dir(targetPath,
                    new Files(Path.Combine(SourceFiles, "*.*")),
                    new Dir(@"%Desktop%",
                        new ExeFileShortcut(appName, Path.Combine("[INSTALLDIR]", $"{Path.GetFileName(fileExe)}"), arguments: "") { WorkingDirectory = "[INSTALLDIR]" }),
                    new Dir(@"%ProgramMenu%",
                        new ExeFileShortcut(appName, Path.Combine("[INSTALLDIR]", $"{Path.GetFileName(fileExe)}"), arguments: "") { WorkingDirectory = "[INSTALLDIR]" })))
            {
                UI = WUI.WixUI_FeatureTree,
                GUID = guid,
                UpgradeCode = upgradeGuid,
                ProductId = productGuid,
                OutFileName = $"SetupExcelToCsv-v.{version}",
                InstallerVersion = 200,
                Platform = Platform.x64,
                Version = version,
                Language = "ru-RU",
                //LicenceFile = Path.Combine(InstallerSourceFiles, "license.rtf"),
                BannerImage = Path.Combine(InstallerSourceFiles, "BunnerImage.png"),
                BackgroundImage = Path.Combine(InstallerSourceFiles, "BackgroundImage1.png"),
                ControlPanelInfo =
                {
                    Manufacturer = "LK-Proekt LTD",
                },
                MajorUpgrade = new MajorUpgrade()
                {
                    AllowDowngrades = false,
                    AllowSameVersionUpgrades = false,
                    Disallow = false,
                    DowngradeErrorMessage = "A later version of [ProductName] is already installed. Setup will now exit.",
                    MigrateFeatures = true,
                    Schedule = UpgradeSchedule.afterInstallInitialize,
                },
                ManagedUI = new ManagedUI()
            };

            project.ManagedUI.InstallDialogs.Add(Dialogs.Welcome)
                .Add(Dialogs.InstallDir)
                .Add(Dialogs.Progress)
                .Add(Dialogs.Exit);

            project.ManagedUI.ModifyDialogs.Add(Dialogs.MaintenanceType)
                .Add(Dialogs.Features)
                .Add(Dialogs.Progress)
                .Add(Dialogs.Exit);

            return project;
        }
    }
}
