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
		public bool Standalone { get; private set; } = false;
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
			Program.LogVerbose("Preference initialization started");
			bool prefsvalid = PrefsExistAndNotNull(PrefPath);
			if (!prefsvalid && silent) {
				throw new ArgumentException(Properties.Resources.SilentEmptyRegistryError);
			} else if (!prefsvalid && !silent) {
				Preferences.InitPrefs(PrefPath);
			}
			this.Silent = silent;
			Program.LogVerbose("Init complete. Loading values into object properties");
			this.LoadVals();
		}
		[return: NotNull]
		private static bool PrefsExistAndNotNull([NotNullWhen(true)]string prefPath) {
			Program.LogVerbose("Pref check started");
			if (prefPath == "") {
				Program.LogVerbose("No pref path provided.");
				return false; 
			}
			RegistryKey? prefKey = Registry.CurrentUser.OpenSubKey(prefPath);
			if (prefKey is null) {
                Program.LogVerbose("Pref key does not exist");
                return false; 
			}
			object? pathKey = prefKey.GetValue("PathToKoL", null);
			if (pathKey is null) {
                Program.LogVerbose("Pref subkey does not exist.");
                return false; 
			}
			if (string.IsNullOrEmpty(pathKey.ToString())) {
                Program.LogVerbose("Path value is null or blank");
                return false; 
			}
			if(pathKey.ToString()!.Substring(1,1) != ":" && pathKey.ToString()!.Substring(1, 1) != @"\") {
				Program.LogVerbose("Invalid path saved previously, re-init will occur.");
				return false;
			}
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
			Program.LogVerbose("Decision to create new prefs.");
			if (prefPath is null) {
                Program.LogVerbose("No pref path provided. Unreachable");
                throw new UnreachableException();
			}
            Program.LogVerbose("Creating pref key");
            RegistryKey? prefKey = Registry.CurrentUser.CreateSubKey(prefPath);
            Program.LogVerbose("Creating jarpath object");
            object jarPath = prefKey.GetValue("PathToKoL", "");
			bool askForDir;
			if (jarPath.ToString() == "" || (jarPath.ToString()!.Substring(1, 1) != ":" && jarPath.ToString()!.Substring(1, 1) != @"\")) {
                Program.LogVerbose("Decision to ask for jarpath from user");
                OpenFileDialog dialog = new() {
                    InitialDirectory = Environment.GetEnvironmentVariable("UserProfile"),
                    Filter = "JAR files (*.jar)| *.jar",
                    Title = Properties.Resources.JarFileSelectTitle,
					ShowHelp = true
                };
                Program.LogVerbose("Creating dialog");
				Program.LogVerbose($"InitialDirectory: {dialog.InitialDirectory}");
				Program.LogVerbose($"Filter: {dialog.Filter}");
				Program.LogVerbose($"Title: {dialog.Title}");
                if (dialog.ShowDialog() == DialogResult.OK && !String.IsNullOrEmpty(dialog.FileName)) {
					Program.LogVerbose($"Selected file: {dialog.FileName};");
					FileInfo selectedJar = new(dialog.FileName);
					Program.LogVerbose(selectedJar.Directory!.FullName);
					string installPath = selectedJar.Directory!.FullName;
                    Program.LogVerbose("JAR identified. Saving path to prefkey");
                    prefKey.SetValue("PathToKoL", installPath, RegistryValueKind.String);
                    Program.LogVerbose($"InstallPath: {installPath};");
                    askForDir = false;
				} else {
					askForDir = true;
				}
				dialog.Dispose();
			} else {
                Program.LogVerbose("Jarpath exits");
                askForDir = false;
			}
			if (askForDir) {
                Program.LogVerbose("Location of JAR was requested, but not provided. Asking for dir to install new jar");
                string? installPath = AskForDir();
				if (installPath is null) {
                    Program.LogVerbose("User cancelled out of request. No further action possible.");
                    Environment.Exit(0);
				}
                Program.LogVerbose("Path identified, saving to prefkey");
                prefKey.SetValue("PathToKoL", installPath, RegistryValueKind.String);
                Program.LogVerbose($"InstallPath: {installPath};");
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
			string AssemblyDir = AppContext.BaseDirectory;
			string[] Runtimes = Directory.GetFiles(AssemblyDir, "vcruntime*.dll");
			if (Runtimes.Length > 0) {
				Standalone = true;
			} else {
				Standalone = false;
			}
		}
	}
}
