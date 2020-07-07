﻿using System;
using System.Runtime.InteropServices;
using System.Threading;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace AutoCloseFolder
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration(Vsix.Name, Vsix.Description, Vsix.Version)]
    [ProvideAutoLoad(UIContextGuids.SolutionHasSingleProject, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(UIContextGuids.SolutionHasMultipleProjects, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideOptionPage(typeof(Options), "Environment", Vsix.Name, 101, 102, true, new string[] { }, ProvidesLocalizedCategoryName = false)]
    [Guid(Vsix.Id)]
    public sealed class AutoCloseFolderPackage : AsyncPackage
    {
        private AutoCloseFolder _autoCloseFolder;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            var dte = await GetServiceAsync(typeof(DTE)) as DTE2;
            var options = (Options)GetDialogPage(typeof(Options));

            _autoCloseFolder = new AutoCloseFolder(this, dte, options);
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (disposing)
            {
                _autoCloseFolder?.Dispose();
            }
        }
    }
}
