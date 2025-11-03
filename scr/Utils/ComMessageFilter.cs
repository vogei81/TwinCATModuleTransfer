using System;
using System.Runtime.InteropServices;

namespace TwinCATModuleTransfer.Utils
{
    public class ComMessageFilter : ComMessageFilter.IOleMessageFilter
    {
        public static void Register()
        {
            IOleMessageFilter oldFilter;
            CoRegisterMessageFilter(new ComMessageFilter(), out oldFilter);
        }

        public static void Revoke()
        {
            IOleMessageFilter oldFilter;
            CoRegisterMessageFilter(null, out oldFilter);
        }

        int ComMessageFilter.IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) { return 0; }
        int ComMessageFilter.IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            return dwRejectType == 2 ? 100 : -1; // SERVERCALL_RETRYLATER
        }
        int ComMessageFilter.IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) { return 2; }

        [DllImport("Ole32.dll")] private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
        [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IOleMessageFilter
        {
            [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
            [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
            [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
        }
    }
}