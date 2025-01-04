using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using EnvDTE;
using EnvDTE80;
using Microsoft.Internal.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace AutoCloseFolder
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration(Vsix.Name, Vsix.Description, Vsix.Version)]
    [ProvideAutoLoad(UIContextGuids.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
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

            if (await GetServiceAsync(typeof(SVsUIShell)) is IVsUIShell shell)
            {
                // major hack in action here
                // I don't know how to get expanded/collapsed events with an official VS api
                var solutionExplorer = new Guid(ToolWindowGuids80.SolutionExplorer);
                shell.FindToolWindow(0, solutionExplorer, out var frame);
                var view = GetFrameView(frame);
                var content = GetContent(GetContent(GetContent(GetContent(view))));

                // tree is Microsoft.Internal.VisualStudio.PlatformUI.VirtualizingTreeView, note this is a ListBox object
                var tree = GetTreeView(content);
                if (tree != null)
                {
                    var collapsed = tree.GetType().GetEvent("NodeCollapsed");
                    if (collapsed != null)
                    {
                        collapsed.AddEventHandler(tree, new EventHandler<TreeNodeEventArgs>(OnNodeCollapsed));
                    }

                    var expanded = tree.GetType().GetEvent("NodeExpanded");
                    if (expanded != null)
                    {
                        expanded.AddEventHandler(tree, new EventHandler<TreeNodeEventArgs>(OnNodeExpanded));
                    }
                }
            }
        }

        private void OnNodeExpanded(object sender, TreeNodeEventArgs e) => _autoCloseFolder?.RestartTimer(_autoCloseFolder.Options.FinalPeriodOnExpansion);
        private void OnNodeCollapsed(object sender, TreeNodeEventArgs e) => _autoCloseFolder?.RestartTimer();

        private static object GetTreeView(object instance) => GetPropertyValue(instance, "TreeView");
        private static object GetFrameView(object instance) => GetPropertyValue(instance, "FrameView");
        private static object GetContent(object instance) => GetPropertyValue(instance, "Content");
        private static object GetPropertyValue(object instance, string propertyName)
        {
            if (instance == null)
                return null;

            var pi = instance.GetType().GetProperties().FirstOrDefault(p => p.Name == propertyName);
            if (pi == null || !pi.CanRead || pi.GetIndexParameters().Length > 0)
                return null;

            return pi.GetValue(instance);
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (disposing)
            {
                Interlocked.Exchange(ref _autoCloseFolder, null)?.Dispose();
            }
        }
    }
}
