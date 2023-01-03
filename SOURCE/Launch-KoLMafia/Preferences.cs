using Microsoft.Win32;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System.Security.AccessControl;
using System.Linq;
using System.Security.Principal;

namespace Launch_KoLMafia {
	internal sealed class Preferences {
		public required string PrefPath { get; init; }
		public string InstallLocation { get; private set; } = "";
		public int MaxAttempts { get; private set; } = 3;
		public bool Silent { get; private set; } = false;
		private string? _skippedVersion;
		public string? SkippedVersion {
			get { return _skippedVersion; }
			set {
				if (value is null) throw new ArgumentNullException(nameof(value));
				RegistryKey prefKey = Registry.CurrentUser.OpenSubKey(PrefPath, true)!;
				prefKey.SetValue("SkippedVersion", value, RegistryValueKind.String);
				_skippedVersion = value;
			}
		}
		[SetsRequiredMembers]
		public Preferences(string prefPath) {
			PrefPath = prefPath;
			this.Initialize(false);
		}
		[SetsRequiredMembers]
		public Preferences(string prefPath, bool silent) {
			PrefPath = prefPath;
			this.Initialize(silent);
		}

		public bool ConfirmPermissions() {
			if (!Preferences.PrefsExistAndNotNull(PrefPath)) return false; 
			RegistryKey prefKey = Registry.CurrentUser.OpenSubKey(PrefPath)!;
			RegistrySecurity prefKeySec = prefKey.GetAccessControl();
			AuthorizationRuleCollection rules = prefKeySec.GetAccessRules(true, true, typeof(System.Security.Principal.NTAccount));
			string thisUser = WindowsIdentity.GetCurrent().Name;

			foreach (var rule in rules.Cast<RegistryAccessRule>()) {
				if (rule.IdentityReference.Value == thisUser) {
					bool ruleAllow = rule.AccessControlType == AccessControlType.Allow;
					bool writeAccess = rule.RegistryRights == RegistryRights.FullControl || rule.RegistryRights == RegistryRights.WriteKey;
					if (ruleAllow && writeAccess) {
						return true;
					}
				} else {
					continue;
				}
			}
			return false;
		}
		private void Initialize(bool silent) {
			if (!(Preferences.PrefsExistAndNotNull(PrefPath)) && silent) {
				throw new ArgumentException(Properties.Resources.SilentEmptyRegistryError);
			} else if (!(Preferences.PrefsExistAndNotNull(PrefPath)) && !silent) {
				Preferences.InitPrefs(PrefPath);
			}
			this.Silent = silent;
			this.LoadVals();
		}
		[return: NotNull]
		private static bool PrefsExistAndNotNull([NotNullWhen(true)]string prefPath) {
			if (prefPath == "") { return false; }
			RegistryKey? prefKey = Registry.CurrentUser.OpenSubKey(prefPath);
			if (prefKey is null) { return false; }
			object? pathKey = prefKey.GetValue("PathToKoL", null);
			if (pathKey is null) { return false; }
			if (string.IsNullOrEmpty(pathKey.ToString())) { return false; }
			return true;
		}
		[return: MaybeNull]
		private static string? AskForDir() {
			string msg = Properties.Resources.TargetFolderDialogMsg;
			string title = Properties.Resources.TargetFolderDialogTitle;
			MessageBoxButtons buttons = MessageBoxButtons.YesNo;
			DialogResult nobutton = DialogResult.No;
			MessageBoxIcon question = MessageBoxIcon.Question;
			MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button1;
			DialogResult answer = MessageBox.Show(msg, title, buttons, question, defaultButton);
			if (answer == nobutton) {
				return null;
			} else {
				using FolderBrowserDialog selectFolder = new();
				selectFolder.SelectedPath = Environment.GetEnvironmentVariable("UserProfile") ?? @"C:\";
				selectFolder.Description = Properties.Resources.FolderSelectTitle;
				selectFolder.ShowNewFolderButton = true;
				DialogResult result = selectFolder.ShowDialog();
				if (result == DialogResult.Cancel) {
					return null;
				} else {
					return selectFolder.SelectedPath;
				}
			}
		}
		private static void InitPrefs(string prefPath) {
			if (prefPath is null) {
	throw new UnreachableException();
			}
			RegistryKey? prefKey = Registry.CurrentUser.CreateSubKey(prefPath);
			object jarPath = prefKey.GetValue("PathToKoL", "");
			bool askForDir;
			if (jarPath.ToString() == "") {
				using OpenFileDialog dialog = new();
				dialog.InitialDirectory = Environment.GetEnvironmentVariable("UserProfile");
				dialog.Filter = "JAR files (*.jar)| *.jar";
				dialog.Title = Properties.Resources.JarFileSelectTitle;
				if (dialog.ShowDialog() == DialogResult.OK || dialog.FileName is not null) {
					FileInfo selectedJar = new(dialog.FileName);
					string installPath = selectedJar.Directory!.Name;
					prefKey.SetValue("PathToKoL", installPath, RegistryValueKind.String);
					askForDir = false;
				} else {
					askForDir = true;
				}
			} else {
				askForDir = false;
			}
			if (askForDir) {
				string? installPath = AskForDir();
				if (installPath is null) {
					Environment.Exit(0);
				}
				prefKey.SetValue("PathToKoL", installPath, RegistryValueKind.String);
			}
			if (prefKey.GetValue("MaxDownloadAttempts", null) is null) {
				prefKey.SetValue("MaxDownloadAttempts", 3, RegistryValueKind.DWord);
			}
			if (prefKey.GetValue("SkippedVersion", null) is null) {
				prefKey.SetValue("SkippedVersion", "", RegistryValueKind.String);
			}
		}
		private void LoadVals() {
			if (!(Preferences.PrefsExistAndNotNull(this.PrefPath))) {
				Preferences.InitPrefs(this.PrefPath);
			}
			RegistryKey? prefKey = Registry.CurrentUser.OpenSubKey(PrefPath);
			if (prefKey is null) { throw new UnreachableException(); }
			InstallLocation = prefKey.GetValue("PathToKoL", "").ToString()!;
			MaxAttempts = (int)prefKey.GetValue("MaxDownloadAttempts", 3);
			if (prefKey.GetValue("SkippedVersion", null) is null) {
				prefKey.SetValue("SkippedVersion", "", RegistryValueKind.String);
			}
			_skippedVersion = prefKey.GetValue("SkippedVersion", "").ToString();
		}
	}
}
