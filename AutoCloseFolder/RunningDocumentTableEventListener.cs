using System;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace AutoCloseFolder
{
    public sealed class RunningDocumentTableEventListener : IDisposable, IVsRunningDocTableEvents
    {
        public event EventHandler Change;

        private readonly uint _cookie;

        public RunningDocumentTableEventListener(RunningDocumentTable table)
        {
            if (table == null)
                throw new ArgumentNullException(nameof(table));

            Table = table;
            _cookie = table.Advise(this);
        }

        public RunningDocumentTable Table { get; }

        public void Dispose() => Table.Unadvise(_cookie);

        private int OnChange()
        {
            Change?.Invoke(this, EventArgs.Empty);
            return 0;
        }

        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining) => OnChange();
        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining) => OnChange();
        public int OnAfterSave(uint docCookie) => OnChange();
        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs) => OnChange();
        public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame) => OnChange();
        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame) => OnChange();
    }
}
