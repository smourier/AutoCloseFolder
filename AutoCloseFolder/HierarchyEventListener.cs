using System;
using System.Collections.Concurrent;
using Microsoft.VisualStudio.Shell.Interop;

namespace AutoCloseFolder
{
    public class HierarchyEventListener : IVsHierarchyEvents, IVsHierarchyEvents2
    {
        public event EventHandler Change;

        private readonly ConcurrentDictionary<IVsHierarchy, uint> _cookies = new ConcurrentDictionary<IVsHierarchy, uint>();

        public void AddHierarchy(IVsHierarchy hierarchy)
        {
            if (hierarchy == null)
                return;

            if (hierarchy.AdviseHierarchyEvents(this, out var cookie) == 0)
            {
                _cookies[hierarchy] = cookie;
            }
        }

        public void RemoveHierarchy(IVsHierarchy hierarchy)
        {
            if (hierarchy == null)
                return;

            if (_cookies.TryRemove(hierarchy, out var cookie))
            {
                hierarchy.UnadviseHierarchyEvents(cookie);
            }
        }

        private int OnChange()
        {
            Change?.Invoke(this, EventArgs.Empty);
            return 0;
        }

        public int OnItemAdded(uint itemidParent, uint itemidSiblingPrev, uint itemidAdded, bool ensureVisible) => OnChange();
        public int OnItemAdded(uint itemidParent, uint itemidSiblingPrev, uint itemidAdded) => OnChange();
        public int OnItemsAppended(uint itemidParent) => OnChange();
        public int OnItemDeleted(uint itemid) => OnChange();
        public int OnPropertyChanged(uint itemid, int propid, uint flags) => OnChange();
        public int OnInvalidateItems(uint itemidParent) => OnChange();
        public int OnInvalidateIcon(IntPtr hicon) => OnChange();
    }
}
