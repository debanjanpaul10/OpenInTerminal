// ***********************************************************************
// <copyright file="Logger.cs" company="Debanjan Paul">
//     Copyright (c) . All rights reserved.
// </copyright>
// ***********************************************************************

namespace OpenInTerminal.Helpers
{
    using System;
    using System.Diagnostics.CodeAnalysis;
    using Microsoft;
    using Microsoft.VisualStudio.Shell;
    using Microsoft.VisualStudio.Shell.Interop;

    /// <summary>
    /// The Logger Class
    /// </summary>
    public static class Logger
    {
        /// <summary>
        /// The Visual Studio Output Window Pane
        /// </summary>
        private static IVsOutputWindowPane pane;

        /// <summary>
        /// The IService Provider
        /// </summary>
        private static IServiceProvider _provider;

        /// <summary>
        /// The Name string
        /// </summary>
        private static string _name;

        /// <summary>
        /// Initialize the class
        /// </summary>
        /// <param name="provider">The Package provider</param>
        /// <param name="name">The name string</param>
        public static void Initialize(Package provider, string name)
        {
            _provider = provider;
            _name = name;
        }

        /// <summary>
        /// The Log method
        /// </summary>
        /// <param name="message">The message to be logged</param>
        public static void Log(string message)
        {
            if (string.IsNullOrEmpty(message))
                return;

            try
            {
                if (EnsurePane())
                {
                    pane.OutputString(DateTime.Now.ToString() + ": " + message + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex);
            }
        }

        /// <summary>
        /// The exception logger
        /// </summary>
        /// <param name="ex">The exception to be logged</param>
        public static void Log(Exception ex)
        {
            if (ex != null)
            {
                Log(ex.ToString());
            }
        }

        /// <summary>
        /// The Ensure Pane in VS Terminal
        /// </summary>
        /// <returns>The boolean flag</returns>
        private static bool EnsurePane()
        {
            if (pane == null)
            {
                Guid guid = Guid.NewGuid();
                IVsOutputWindow output = (IVsOutputWindow)_provider.GetService(typeof(SVsOutputWindow));
                Assumes.Present(output);

                output.CreatePane(ref guid, _name, 1, 1);
                output.GetPane(ref guid, out pane);
            }

            return pane != null;
        }
    }
}
