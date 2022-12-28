// ***********************************************************************
// <copyright file="Options.cs" company="Debanjan Paul">
//     Copyright (c) . All rights reserved.
// </copyright>
// ***********************************************************************

namespace OpenInTerminal
{
    using System;
    using System.ComponentModel;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using Microsoft.VisualStudio.Shell;

    /// <summary>
    /// The Options Class
    /// </summary>
    [ComVisible(true)]
    public class Options : DialogPage
    {
        /// <summary>
        /// The Command Line Arguments
        /// </summary>
        [Category("General")]
        [DisplayName("Command line arguments")]
        [Description("Command line arguments to pass to wt.exe")]
        public string CommandLineArguments { get; set; }

        /// <summary>
        /// The Path to Windows Terminal exe file
        /// </summary>
        [Category("General")]
        [DisplayName("Path to wt.exe")]
        [Description("Specify the path to wt.exe.")]
        public string PathToExe { get; set; } = Environment.ExpandEnvironmentVariables(@"%localappdata%\Microsoft\WindowsApps\wt.exe");

        [Category("General")]
        [DisplayName("Open solution/project as regular file")]
        [Description("When true, opens solutions/projects as regular files and does not load folder path into Windows Terminal.")]
        public bool OpenSolutionProjectAsRegularFile { get; set; }

        protected override void OnApply(PageApplyEventArgs e)
        {
            if (!File.Exists(PathToExe))
            {
                e.ApplyBehavior = ApplyKind.Cancel;
                MessageBox.Show($"The file \"{PathToExe}\" doesn't exist.", Vsix.Name, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            base.OnApply(e);
        }
    }
}
