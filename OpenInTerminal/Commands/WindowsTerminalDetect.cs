// ***********************************************************************
// <copyright file="WindowsTerminalDetect.cs" company="Debanjan Paul">
//     Copyright (c) . All rights reserved.
// </copyright>
// ***********************************************************************

namespace OpenInTerminal.Commands
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Win32;

    /// <summary>
    /// The Windows Terminal Detect class
    /// </summary>
    internal static class WindowsTerminalDetect
    {
        /// <summary>
        /// The Registry method
        /// </summary>
        /// <returns>The string of the registry</returns>
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

        /// <summary>
        /// The Local App data method
        /// </summary>
        /// <returns>The path of the exe in local app data</returns>
        internal static string InLocalAppData()
        {
            var localAppData = Environment.GetEnvironmentVariable("LOCALAPPDATA");

            var codePartDir = @"Microsoft\WindowsApps";
            var codeDir = Path.Combine(localAppData, codePartDir);
            var drives = DriveInfo.GetDrives();

            foreach (var drive in drives)
            {
                if (drive.DriveType == DriveType.Fixed)
                {
                    var path = Path.Combine(drive.Name[0] + codeDir.Substring(1), "wt.exe");
                    if (File.Exists(path))
                    {
                        return path;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Returns the presence of the exe file in Environment Variable in Path
        /// </summary>
        /// <returns>The Path variable</returns>
        internal static string InEnvVarPath()
        {
            var envPath = Environment.GetEnvironmentVariable("Path");
            var paths = envPath.Split(';');
            var parentDir = "WindowsApps";
            foreach (var path in paths)
            {
                if (path.ToLower().Contains("wt"))
                {
                    var temp = Path.Combine(path.Substring(0, path.IndexOf(parentDir)),
                        parentDir, "wt.exe");
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
