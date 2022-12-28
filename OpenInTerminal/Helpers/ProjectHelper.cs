// ***********************************************************************
// <copyright file="ProjectHelper.cs" company="Debanjan Paul">
//     Copyright (c) . All rights reserved.
// </copyright>
// ***********************************************************************

namespace OpenInTerminal.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Windows.Forms;
    using EnvDTE;
    using EnvDTE80;

    /// <summary>
    /// The Project Helper Class
    /// </summary>
    internal static class ProjectHelper
    {
        /// <summary>
        /// Gets the selected Path
        /// </summary>
        /// <param name="dte">The DTE</param>
        /// <param name="openSolutionProjectAsRegularFile">The open solution as a regular file boolean</param>
        /// <returns>The string of the selected path</returns>
        public static string GetSelectedPath(DTE2 dte, bool openSolutionProjectAsRegularFile)
        {
            var items = (Array)dte.ToolWindows.SolutionExplorer.SelectedItems;
            var files = new List<string>();

            foreach (UIHierarchyItem selItem in items)
            {
                ProjectItem item = selItem.Object as ProjectItem;

                if (item != null)
                    files.Add(item.GetFilePath());

                Project proj = selItem.Object as Project;

                if (proj != null)
                    return openSolutionProjectAsRegularFile ? $"\"{proj.FileName}\"" : proj.GetRootFolder();

                Solution sol = selItem.Object as Solution;

                if (sol != null)
                    return openSolutionProjectAsRegularFile ? $"\"{sol.FullName}\"" : Path.GetDirectoryName(sol.FileName);
            }

            return files.Count > 0 ? String.Join(" ", files) : null;
        }

        /// <summary>
        /// Gets the file path
        /// </summary>
        /// <param name="item">The Project Item</param>
        /// <returns>The String of the file path</returns>
        public static string GetFilePath(this ProjectItem item)
        {
            return $"\"{item.FileNames[1]}\""; // Indexing starts from 1
        }

        /// <summary>
        /// Gets the root folder
        /// </summary>
        /// <param name="project">The Project details</param>
        /// <returns>The string path of the root folder</returns>
        public static string GetRootFolder(this Project project)
        {
            if (string.IsNullOrEmpty(project.FullName))
                return null;

            string fullPath;

            try
            {
                fullPath = project.Properties.Item("FullPath").Value as string;
            }
            catch (ArgumentException)
            {
                try
                {
                    // MFC projects don't have FullPath, and there seems to be no way to query existence
                    fullPath = project.Properties.Item("ProjectDirectory").Value as string;
                }
                catch (ArgumentException)
                {
                    // Installer projects have a ProjectPath.
                    fullPath = project.Properties.Item("ProjectPath").Value as string;
                }
            }

            if (string.IsNullOrEmpty(fullPath))
            {
                return File.Exists(project.FullName) ? Path.GetDirectoryName(project.FullName) : null;
            }

            if (Directory.Exists(fullPath))
            {
                return fullPath;
            }

            if (File.Exists(fullPath))
            {
                return Path.GetDirectoryName(fullPath);
            }

            return null;
        }
    }
}
