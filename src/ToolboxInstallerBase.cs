using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Configuration.Install;

using Microsoft.Win32;

namespace MattBerther.Install
{
	/// <summary>
	/// Implements the basic functionality required to add a components controls to a tab in Visual Studio.NET.
	/// </summary>
	/// <remarks></remarks>
	/// <example>
	/// The follwing example derives a class to install the components installed from MattBerther.Components.dll to a
	/// new tab named MattBerther. This class will also remove the tab when uninstalling.
	/// <code>
	/// [
	/// RunInstaller(true),
	/// ToolboxItem(false)
	/// ]
	/// public class MyComponentToolboxInstaller : MattBerther.Install.ToolboxInstallerBase
	/// {
	///		protected override string ComponentName
	///		{
	///			get { return "MattBerther.Com Controls"; }
	///		}
	///		
	///		protected override string ComponentPath
	///		{
	///			get
	///			{
	///				return System.IO.Path.Combine(this.Context.Parameters["INSTALLDIR"], "MattBerther.Components.dll"); 
	///			}
	///		}
	///		
	///		protected override TabName
	///		{
	///			get { return "MattBerther"; }
	///		}
	///		
	///		protected override bool UninstallRemovesTab
	///		{
	///			get { return true; }
	///		}
	///		
	///		public override Install(System.Collections.IDictionary stateSaver)
	///		{
	///			// Perform check to make sure that VS.NET isnt running
	///			// This will allow us to be certain that our tab shows up properly.
	///			while (this.IsVisualStudioRunning)
	///			{
	///				"One or more instances of Visual Studio .NET are running.\r\n\r\nTo continue, please shutdown all running instances of VS.NET,\r\n and click 'Retry'.", 
	///				"MattBerther.Com Controls", MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning);
	///				if (dr == DialogResult.Cancel)
	///				{
	///					throw new InstallException("Unable to continue.");
	///				}
	///			}
	///			
	///			base.Install(stateSaver);
	///		}
	/// }
	/// </code>
	/// </example>
	public abstract class ToolboxInstallerBase : Installer
	{
		/// <summary>
		/// The full file system path to the component to add to the toolbox tab.
		/// </summary>
		protected abstract string ComponentPath { get; }
		
		/// <summary>
		/// The unique name of the component to add to the toolbox tab.
		/// </summary>
		protected abstract string ComponentName { get; }

		/// <summary>
		/// The name of the tab to be created.
		/// </summary>
		protected abstract string TabName { get; }

		/// <summary>
		///  Whether or not an uninstall removes the tab.
		/// </summary>
		/// <value><b>true</b> if the tab should be removed when uninstalling, otherwise <b>false</b>. The default is <b>false</b>.</value>
		protected virtual bool UninstallRemovesTab 
		{
			get { return false; }
		}

		/// <summary>
		/// Gets a value indicating whether or not Visual Studio.NET is running
		/// </summary>
		/// <remarks>
		/// This implementation relies on a registry key at HKLM\SOFTWARE\Microsoft\VisualStudio\7.1.
		/// Override this method if you need to test for a different version.
		/// </remarks>
		protected virtual bool IsVisualStudioRunning
		{
			get
			{
				RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\VisualStudio\7.1", false);
				string installDir = (string)key.GetValue("InstallDir");

				string fullPath = System.IO.Path.Combine(installDir, "devenv.exe");

				System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("devenv");
				foreach (System.Diagnostics.Process process in processes)
				{
					if (String.Compare(process.MainModule.FileName, fullPath, true) == 0)
					{
						return true;
					}
				}

				return false;
			}
		}

		/// <summary>
		/// Overrides <see cref="System.Configuration.Install.Installer.Install"/> to add the toolbox tab.
		/// </summary>
		/// <param name="stateSaver">An <see cref="System.Collections.IDictionary"/> used to save information needed to perform a commit, rollback, or uninstall operation.</param>
		public override void Install(IDictionary stateSaver)
		{
			base.Install (stateSaver);
			
			MessageFilter.Register();
			AddToolBoxItem(this.VisualStudioDTE);
			MessageFilter.Revoke();
		}

		/// <summary>
		/// Overrides <see cref="System.Configuration.Install.Installer.Uninstall"/> to optionally remove the tab.
		/// </summary>
		/// <param name="savedState">An <see cref="System.Collections.IDictionary"/> used to save information needed to perform a commit, rollback, or uninstall operation.</param>
		public override void Uninstall(IDictionary savedState)
		{
			if (this.UninstallRemovesTab)
			{
				MessageFilter.Register();
				RemoveToolBoxItem(this.VisualStudioDTE);
				MessageFilter.Revoke();
			}

			base.Uninstall (savedState);
		}

		/// <summary>
		/// Gets the latest version VisualStudio Design Time Environment
		/// </summary>
		protected EnvDTE.DTE VisualStudioDTE
		{
			get 
			{
				Type latestDTE = Type.GetTypeFromProgID("VisualStudio.DTE"); // version independent progID
				return Activator.CreateInstance(latestDTE) as EnvDTE.DTE;
			}
		}

		private void AddToolBoxItem(EnvDTE.DTE env)
		{
			EnvDTE.Window toolboxWindow = env.Windows.Item(EnvDTE.Constants.vsWindowKindToolbox);
			EnvDTE.ToolBox toolbox = (EnvDTE.ToolBox)toolboxWindow.Object;
			EnvDTE.ToolBoxTabs toolboxTabs = toolbox.ToolBoxTabs;

			EnvDTE.ToolBoxTab newTab = null;

			bool tabExists = false;
			foreach (EnvDTE.ToolBoxTab tab in toolboxTabs)
			{
				if (tab.Name == this.TabName)
				{
					newTab = tab;
					tabExists = true;
					break;
				}
			}

			if (!tabExists)
			{
				newTab = toolboxTabs.Add(this.TabName);
			}

			newTab.Activate();
			env.ExecuteCommand("View.PropertiesWindow", "");
			newTab.ToolBoxItems.Item(1).Select();

			newTab.ToolBoxItems.Add(this.ComponentName, this.ComponentPath, 
				EnvDTE.vsToolBoxItemFormat.vsToolBoxItemFormatDotNETComponent);

			env.Quit();
		}

		private void RemoveToolBoxItem(EnvDTE.DTE env)
		{
			EnvDTE.Window toolboxWindow = env.Windows.Item(EnvDTE.Constants.vsWindowKindToolbox);
			EnvDTE.ToolBox toolbox = (EnvDTE.ToolBox)toolboxWindow.Object;
			EnvDTE.ToolBoxTabs toolboxTabs = toolbox.ToolBoxTabs;

			foreach (EnvDTE.ToolBoxTab tab in toolboxTabs)
			{
				if (tab.Name == this.TabName)
				{
					foreach (EnvDTE.ToolBoxItem item in tab.ToolBoxItems)
					{
						try
						{
							item.Delete();
						} 
						catch (Exception) {}
					}

					try
					{
						tab.Delete();
					} 
					catch (Exception) {}
				}
			}

			env.Quit();
		}
	}
}
