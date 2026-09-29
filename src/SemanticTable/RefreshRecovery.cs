using System;
using System.Runtime.InteropServices;

namespace SemanticTable
{
    internal static class RefreshRecovery
    {
        // Excel reports stale sessions using its generic refresh error. One recovery
        // attempt is safe even when that error has another cause; never loop.
        internal static bool Run(Func<bool> refresh, Action restorePrevious,
            Action reapplyRequested, Action<string> log)
        {
            try { return refresh(); }
            catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x800A03EC))
            {
                log("Initial refresh failed; attempting recovery with the previous query. " + ex);
            }

            restorePrevious();
            if (!refresh()) return false;
            log("Previous query refreshed successfully; retrying the requested query once.");
            reapplyRequested();
            return refresh();
        }
    }
}
