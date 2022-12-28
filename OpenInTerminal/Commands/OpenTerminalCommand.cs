// ***********************************************************************
// <copyright file="OpenTerminalCommand.cs" company="Debanjan Paul">
//     Copyright (c) . All rights reserved.
// </copyright>
// ***********************************************************************

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
    using OpenInTerminal.Commands;

    /// <summary>
    /// The Open In Windows Terminal Command Class
    /// </summary>
    internal sealed class OpenTerminalCommand
    {
        /// <summary>
        /// The Package property
        /// </summary>
        private readonly Package _package;

        /// <summary>
        /// The Options property
        /// </summary>
        private readonly Options _options;

        /// <summary>
        /// Create a new instance of <see cref="OpenTerminalCommand"/> class
        /// </summary>
        /// <param name="package">The Package</param>
        /// <param name="options">The Options</param>
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

        /// <summary>
        /// Creates an instance of <see cref="OpenTerminalCommand"/>
        /// </summary>
        public static OpenTerminalCommand Instance { get; private set; }

        /// <summary>
        /// The IService Provider
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get { return _package; }
        }

        /// <summary>
        /// Initializes the class
        /// </summary>
        /// <param name="package">The Package parameters</param>
        /// <param name="options">The Options parameters</param>
        public static void Initialize(Package package, Options options)
        {
            Instance = new OpenTerminalCommand(package, options);
        }

        /// <summary>
        /// The Open Current File in Terminal Command
        /// </summary>
        /// <param name="sender">The sender object</param>
        /// <param name="e">The Event trigger</param>
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

                        OpenWTerminal(path, line);
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

        /// <summary>
        /// The Open Folder in Terminal
        /// </summary>
        /// <param name="sender">The sender object</param>
        /// <param name="e">The Event Trigger</param>
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

                    OpenWTerminal(path, line);
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

        /// <summary>
        /// Opens the Windows Terminal
        /// </summary>
        /// <param name="path">The path to the file</param>
        /// <param name="line">default line integer variable</param>
        private void OpenWTerminal(string path, int line = 0)
        {
            EnsurePathExist();
            bool isDirectory = Directory.Exists(path);

            var args = "";
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

        /// <summary>
        /// Ensures the path to the file exists
        /// </summary>
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

        /// <summary>
        /// Saving options
        /// </summary>
        /// <param name="options">The Options parameter</param>
        /// <param name="path">The Path variable</param>
        private void SaveOptions(Options options, string path)
        {
            options.PathToExe = path;
            options.SaveSettingsToStorage();
        }

    }
}




