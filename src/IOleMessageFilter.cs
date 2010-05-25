using System;
using System.Runtime.InteropServices;

namespace MattBerther.Install
{
	[ComImport(), Guid("00000016-0000-0000-C000-000000000046"),    
	InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
	interface IOleMessageFilter // deliberately renamed to avoid confusion w/ System.Windows.Forms.IMessageFilter
	{
		[PreserveSig]
		int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

		[PreserveSig]
		int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

		[PreserveSig]
		int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
	}
}
