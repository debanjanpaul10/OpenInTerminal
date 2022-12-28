// ***********************************************************************
// <copyright file="OpenInTerminalPackage.cs" company="Debanjan Paul">
//     Copyright (c) . All rights reserved.
// </copyright>
// ***********************************************************************


namespace OpenInTerminal
{
    using Community.VisualStudio.Toolkit;
    using Microsoft.VisualStudio.Shell;
    using System;
    using System.Threading.Tasks;
    using OpenInTerminal.Helpers;
    using System.Runtime.InteropServices;
    using System.Threading;

    /// <summary>
    /// The OpenInTerminal Package
    /// </summary>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration(Vsix.Name, Vsix.Description, Vsix.Version)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideOptionPage(typeof(Options), "Web", Vsix.Name, 101, 102, true, new string[0], ProvidesLocalizedCategoryName = false)]
    [Guid(PackageGuids.guidPackageString)]
    public sealed class OpenInTerminalPackage : AsyncPackage
    {
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync();

            var options = (Options)GetDialogPage(typeof(Options));

            Logger.Initialize(this, Vsix.Name);
            OpenTerminalCommand.Initialize(this, options);
        }
    }
}