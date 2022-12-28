

namespace OpenInTerminal
{
    using EnvDTE;
    using EnvDTE80;
    using Microsoft.VisualStudio.Shell;
    using System;
    using System.ComponentModel.Design;
    using System.IO;
    using System.Windows.Forms;
    using Microsoft.Win32;
    using Microsoft;
    using OpenInTerminal.Helpers;
    using OpenInTerminal;

    internal sealed class OpenTerminalCommand
    {
        private readonly Package _package;
        private readonly Options _options;

        private OpenTerminalCommand(Package package, Options options)
        {
            _package = package;
            _options = options;

            var commandService = (OleMenuCommandService)ServiceProvider.GetService(typeof(IMenuCommandService));

            if (commandService != null)
            {
                var menuCommandID = new CommandID(PackageGuids.guidOpenInTerminalCmdSet, PackageIds.OpenInTerminal);
                var menuItem = new MenuCommand(OpenFolderInTerminal, menuCommandID);
                commandService.AddCommand(menuItem);

                var currentCommandID = new CommandID(PackageGuids.guidOpenCurrentInTerminalCmdSet, PackageIds.OpenCurrentInTerminal);
                var currentItem = new MenuCommand(OpenCurrentFileInTerminal, currentCommandID);
                commandService.AddCommand(currentItem);
            }
        }

        public static OpenTerminalCommand Instance { get; private set; }

        private IServiceProvider ServiceProvider
        {
            get { return _package; }
        }

        public static void Initialize(Package package, Options options)
        {
            Instance = new OpenTerminalCommand(package, options);
        }

        private void OpenCurrentFileInTerminal(object sender, EventArgs e)
        {
            try
            {
                var dte = (DTE2)ServiceProvider.GetService(typeof(DTE));
                Assumes.Present(dte);

                var activeDocument = dte.ActiveDocument;

                if (activeDocument != null)
                {
                    var path = activeDocument.FullName;

                    if (!string.IsNullOrEmpty(path))
                    {
                        int line = 0;

                        if (activeDocument.Selection is TextSelection selection)
                        {
                            line = selection.ActivePoint.Line;
                        }

                        OpenVsCode(path, line);
                    }
                    else
                    {
                        MessageBox.Show("Couldn't resolve the folder");
                    }
                }
                else
                {
                    MessageBox.Show("Couldn't find active document");
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
        }

        private void OpenFolderInTerminal(object sender, EventArgs e)
        {
            try
            {
                var dte = (DTE2)ServiceProvider.GetService(typeof(DTE));
                Assumes.Present(dte);

                string path = ProjectHelper.GetSelectedPath(dte, _options.OpenSolutionProjectAsRegularFile);

                if (!string.IsNullOrEmpty(path))
                {
                    int line = 0;

                    if (dte.ActiveDocument?.Selection is TextSelection selection)
                    {
                        line = selection.ActivePoint.Line;
                    }

                    OpenVsCode(path, line);
                }
                else
                {
                    MessageBox.Show("Couldn't resolve the folder");
                }
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
            }
        }

        private void OpenVsCode(string path, int line = 0)
        {
            EnsurePathExist();
            bool isDirectory = Directory.Exists(path);

            var args = isDirectory ? "." : line > 0 ? $"-g {path}:{line}" : $"{path}";
            if (!string.IsNullOrEmpty(_options.CommandLineArguments))
            {
                args = $"{args} {_options.CommandLineArguments}";
            }

            var start = new System.Diagnostics.ProcessStartInfo()
            {
                FileName = $"\"{_options.PathToExe}\"",
                Arguments = args,
                CreateNoWindow = true,
                UseShellExecute = false,
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden,
            };

            if (isDirectory)
            {
                start.WorkingDirectory = path;
            }

            using (System.Diagnostics.Process.Start(start))
            {

            }
        }

        private void EnsurePathExist()
        {
            if (File.Exists(_options.PathToExe))
                return;

            if (!string.IsNullOrEmpty(WindowsTerminalDetect.InRegistry()))
            {
                SaveOptions(_options, WindowsTerminalDetect.InRegistry());
            }
            else if (!string.IsNullOrEmpty(WindowsTerminalDetect.InEnvVarPath()))
            {
                SaveOptions(_options, WindowsTerminalDetect.InEnvVarPath());
            }
            else if (!string.IsNullOrEmpty(WindowsTerminalDetect.InLocalAppData()))
            {
                SaveOptions(_options, WindowsTerminalDetect.InLocalAppData());
            }
            else
            {
                var box = MessageBox.Show(
                    "I can't find Visual Studio Code (Code.exe). Would you like to help me find it?", Vsix.Name,
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (box == DialogResult.No)
                    return;

                var dialog = new OpenFileDialog
                {
                    DefaultExt = ".exe",
                    FileName = "Code.exe",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    CheckFileExists = true
                };

                var result = dialog.ShowDialog();

                if (result == DialogResult.OK)
                {
                    SaveOptions(_options, dialog.FileName);
                }
            }
        }

        private void SaveOptions(Options options, string path)
        {
            options.PathToExe = path;
            options.SaveSettingsToStorage();
        }

    }

    internal static class WindowsTerminalDetect
    {
        internal static string InRegistry()
        {
            var key = Registry.CurrentUser;
            var name = "Icon";
            try
            {
                var subKey = key.OpenSubKey(@"SOFTWARE\Classes\*\shell\wt\");
                var value = subKey.GetValue(name).ToString();
                if (File.Exists(value))
                {
                    return value;
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        internal static string InLocalAppData()
        {
            var localAppData = Environment.GetEnvironmentVariable("LOCALAPPDATA");

            var codePartDir = @"Programs\Microsoft VS Code";
            var codeDir = Path.Combine(localAppData, codePartDir);
            var drives = DriveInfo.GetDrives();

            foreach (var drive in drives)
            {
                if (drive.DriveType == DriveType.Fixed)
                {
                    var path = Path.Combine(drive.Name[0] + codeDir.Substring(1), "code.exe");
                    if (File.Exists(path))
                    {
                        return path;
                    }
                }
            }

            return null;
        }

        internal static string InEnvVarPath()
        {
            var envPath = Environment.GetEnvironmentVariable("Path");
            var paths = envPath.Split(';');
            var parentDir = "Microsoft VS Code";
            foreach (var path in paths)
            {
                if (path.ToLower().Contains("code"))
                {
                    var temp = Path.Combine(path.Substring(0, path.IndexOf(parentDir)),
                        parentDir, "code.exe");
                    if (File.Exists(temp))
                    {
                        return temp;
                    }
                }
            }
            return null;
        }
    }
}




