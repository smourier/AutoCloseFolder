using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Events;

namespace AutoCloseFolder
{
    public sealed class AutoCloseFolder : IDisposable
    {
        private readonly DTE2 _dte;
        private readonly Options _options;
        private readonly RunningDocumentTable _table;
        private readonly RunningDocumentTableEventListener _documentslistener;
        private readonly HierarchyEventListener _hierarchieslistener;
        private readonly Timer _timer;

        public AutoCloseFolder(IServiceProvider serviceProvider, DTE2 dte, Options options)
        {
            _dte = dte;
            _options = options;
            _table = new RunningDocumentTable(serviceProvider);
            _documentslistener = new RunningDocumentTableEventListener(_table);
            _documentslistener.Change += (s, e) => UpdateTimer();

            _hierarchieslistener = new HierarchyEventListener();
            _hierarchieslistener.Change += (s, e) => UpdateTimer();

            _timer = new Timer((state) => ExecuteCloseFolderWithoutRunningDocuments(), null, _options.FinalPeriod, Timeout.Infinite);

            Microsoft.VisualStudio.Shell.Events.SolutionEvents.OnAfterOpenProject += OnAfterOpenProject;
            Microsoft.VisualStudio.Shell.Events.SolutionEvents.OnBeforeCloseProject += OnBeforeCloseProject;
            Microsoft.VisualStudio.Shell.Events.SolutionEvents.OnBeforeCloseSolution += (s, e) =>
            {
                _timer.Change(Timeout.Infinite, Timeout.Infinite);
                ExecuteCloseSolution();
            };
        }

        private void OnAfterOpenProject(object sender, OpenProjectEventArgs e)
        {
            _hierarchieslistener.AddHierarchy(e.Hierarchy);
            UpdateTimer();
        }

        private void OnBeforeCloseProject(object sender, CloseProjectEventArgs e)
        {
            _hierarchieslistener.RemoveHierarchy(e.Hierarchy);
            UpdateTimer();
        }

        private void UpdateTimer() => _timer.Change(_options.FinalPeriod, Timeout.Infinite);

        private void ExecuteCloseFolderWithoutRunningDocuments()
        {
            if (!_options.CollapseFolders)
                return;

            try
            {
                var docs = new List<string>();
                foreach (var doc in _dte.Documents.Cast<Document>())
                {
                    docs.Add(doc.FullName);
                }

                var hierarchy = _dte.ToolWindows.SolutionExplorer.UIHierarchyItems;
                try
                {
                    _dte.SuppressUI = true;
                    CloseFolderWithoutRunningDocuments(docs, hierarchy);
                }
                finally
                {
                    _dte.SuppressUI = false;
                }
            }
            catch
            {
                // do nothing
            }
            UpdateTimer();
        }

        private void CloseFolderWithoutRunningDocuments(List<string> docs, UIHierarchyItems items)
        {
            foreach (var item in items.Cast<UIHierarchyItem>().Where(item => item.UIHierarchyItems.Count > 0))
            {
                CloseFolderWithoutRunningDocuments(docs, item.UIHierarchyItems);

                if (ShouldCloseFolder(docs, item))
                {
                    item.UIHierarchyItems.Expanded = false;
                }
            }
        }

        private bool ContainsRunningDocument(List<string> docs, UIHierarchyItem item) => EnumerateAllChildItems(item).Any(i => IsRunningDocument(docs, i));
        private bool IsRunningDocument(List<string> docs, UIHierarchyItem item)
        {
            var pi = GetProjectItem(item);
            if (pi == null)
                return true;

            string fullPath = null;
            try
            {
                fullPath = (string)pi.Properties.Item("FullPath")?.Value;
            }
            catch
            {
                // do nothing
            }
            if (fullPath == null)
                return true;

            return docs.Any(d => string.Compare(d, fullPath, StringComparison.OrdinalIgnoreCase) == 0);
        }

        private ProjectItem GetProjectItem(UIHierarchyItem item)
        {
            if (item.Object is ProjectItem pi)
                return pi;

            var parent = item.Collection?.Parent as UIHierarchyItem;
            if (parent != null)
                return GetProjectItem(parent);

            return null;
        }

        private bool ShouldCloseFolder(List<string> docs, UIHierarchyItem item)
        {
            if (!item.UIHierarchyItems.Expanded)
                return false;

            if (item.Object is Project || item.Object is Solution)
                return false;

            return !ContainsRunningDocument(docs, item);
        }

        private IEnumerable<UIHierarchyItem> EnumerateAllChildItems(UIHierarchyItem item)
        {
            foreach (var child in item.UIHierarchyItems.Cast<UIHierarchyItem>())
            {
                yield return child;
                foreach (var grandChild in EnumerateAllChildItems(child))
                {
                    yield return grandChild;
                }
            }
        }

        private void ExecuteCloseSolution()
        {
            if (!_options.CollapseOnClose)
                return;

            var hierarchy = _dte.ToolWindows.SolutionExplorer.UIHierarchyItems;
            try
            {
                _dte.SuppressUI = true;
                CollapseHierarchy(hierarchy);
            }
            finally
            {
                _dte.SuppressUI = false;
            }
        }

        private void CollapseHierarchy(UIHierarchyItems items)
        {
            foreach (var item in items.Cast<UIHierarchyItem>().Where(item => item.UIHierarchyItems.Count > 0))
            {
                CollapseHierarchy(item.UIHierarchyItems);

                if (ShouldCollapse(item))
                {
                    item.UIHierarchyItems.Expanded = false;
                }
            }
        }

        private bool ShouldCollapse(UIHierarchyItem item)
        {
            if (!item.UIHierarchyItems.Expanded)
                return false;

            if (!(item.Object is Project project))
                return true;

            if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder && _options.CollapseSolutionFolders)
                return true;

            if (project.Kind != ProjectKinds.vsProjectKindSolutionFolder && _options.CollapseProjects)
                return true;

            return false;
        }

        public void Dispose()
        {
            _timer.Dispose();
            _documentslistener.Dispose();
        }
    }
}
