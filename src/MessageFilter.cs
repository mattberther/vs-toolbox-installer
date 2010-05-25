using System;
using System.Runtime.InteropServices;

namespace MattBerther.Install
{
	class MessageFilter : IOleMessageFilter
	{
		private const int SERVERCALL_ISHANDLED = 0;
		private const int SERVERCALL_RETRYLATER = 2;
		private const int PENDINGMSGS_WAITDEFPROCESS = 2;

		public static void Register()
		{
			IOleMessageFilter newFilter = new MessageFilter();
			IOleMessageFilter oldFilter = null;

			CoRegisterMessageFilter(newFilter, out oldFilter);
		}

		public static void Revoke()
		{
			IOleMessageFilter oldFilter = null;
			CoRegisterMessageFilter(null, out oldFilter);
		}

		#region IOleMessageFilter Members

		public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
		{
			return SERVERCALL_ISHANDLED;
		}

		public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
		{
			if (dwRejectType == SERVERCALL_RETRYLATER)
			{
				return 99; // retry immediately if return >= 0 or < 100;
			}

			return -1; // cancel call
		}

		public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
		{
			return PENDINGMSGS_WAITDEFPROCESS;
		}

		#endregion

		[DllImport("ole32.dll")]
		private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter outFilter);
	}
}
