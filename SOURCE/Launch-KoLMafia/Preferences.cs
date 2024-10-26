using Microsoft.Win32;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
#if OS_WINDOWS
	using System.Windows.Forms;
#endif


namespace Launch_KoLMafia {
	internal sealed class Preferences {
		private readonly OperatingSystem environment = Environment.OSVersion;
        private string? _skippedVersion;
		private const string DEPRICATED_REG_PATH = @"Software\Sapph Tools\KoLMafia Launcher\";
		private const string DEPRECATED_REG_PARENT = @"Software\Sapph Tools\";
        public required string PrefPath { get; init; }
		public string InstallLocation { get; private set; } = "";
		public int MaxAttempts { get; private set; } = 3;
		public bool Silent { get; private set; } = false;
		public bool Standalone { get; private set; } = false;
		public string? SkippedVersion {
			get { return _skippedVersion; }
			set {
				_skippedVersion = value;
				WritePrefsToFile();
			}
		}
        [SetsRequiredMembers]
        public Preferences(bool silent) {
            if (environment.Platform == PlatformID.Win32NT) {
                PrefPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\KoLMafia Launcher\prefs.json";
            } else {
                PrefPath = "~/.kolmafialauncher/prefs.json";
            }
            this.Initialize(silent);
        }
		private static bool PrefPathIsJson(string prefPath) {
			if (!(new FileInfo(prefPath)).Exists)
				return false;
			string prefData = File.ReadAllText(prefPath);
			try {
				JsonDocument json = JsonDocument.Parse(prefData);
				JsonElement root = json.RootElement;
				if (root.TryGetProperty(nameof(InstallLocation), out JsonElement trash) &&
					root.TryGetProperty(nameof(MaxAttempts), out trash) &&
					root.TryGetProperty(nameof(SkippedVersion), out trash))
						return true;
			} catch (JsonException) {
				return false;
			}
			return false;
		}
		private static bool DeprecatedPrefsAvailable() {
			if (Environment.OSVersion.Platform != PlatformID.Win32NT)
				return false;
            RegistryKey? prefKey = Registry.CurrentUser.OpenSubKey(DEPRICATED_REG_PATH);
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
            if (pathKey.ToString()!.Substring(1, 1) != ":" && pathKey.ToString()!.Substring(1, 1) != @"\") {
                Program.LogVerbose("Invalid path saved previously, re-init will occur.");
                return false;
            }
            return true;
        }
		private void WritePrefsToFile() {
            var preferences = new {
                InstallLocation,
                MaxAttempts,
                SkippedVersion = _skippedVersion,
            };
            string json = JsonSerializer.Serialize(preferences);
            string prefPathParent = Path.GetDirectoryName(PrefPath)!;
            if (!Directory.Exists(prefPathParent)) {
                Directory.CreateDirectory(prefPathParent);
            }
            try {
                File.WriteAllText(PrefPath, json, System.Text.Encoding.UTF8);
            } catch {

            }
        }
		private static bool ConvertRegPrefToFile(string prefPath) {
			try {
				RegistryKey? prefKey = Registry.CurrentUser.OpenSubKey(DEPRICATED_REG_PATH);
				if (prefKey is null) { throw new UnreachableException(); }
				string _installLocation = prefKey.GetValue("PathToKoL", "").ToString()!;
				int _maxAttempts = (int)prefKey.GetValue("MaxDownloadAttempts", 3);
				string _skippedVersion;
				if (prefKey.GetValue("SkippedVersion", null) is null) {
					_skippedVersion = "";
				} else {
					object? value = prefKey.GetValue("SkippedVersion", "");
					if (value is null) {
						_skippedVersion = "";
					} else { 
						_skippedVersion = value.ToString()!;
					}
				}
				var preferences = new {
					InstallLocation = _installLocation,
					MaxAttempts = _maxAttempts,
					SkippedVersion = _skippedVersion,
				};
				string json = JsonSerializer.Serialize(preferences);
				string prefPathParent = Path.GetDirectoryName(prefPath)!;
				if (!Directory.Exists(prefPathParent)) {
					Directory.CreateDirectory(prefPathParent);
				}
				try {
					File.WriteAllText(prefPath, json, System.Text.Encoding.UTF8);
				} catch {
					
				}
				Registry.CurrentUser.DeleteSubKeyTree(DEPRECATED_REG_PARENT);
				return true;
			} catch {
				return false;
			}
        }
        private void Initialize(bool silent) {
			Program.LogVerbose("Preference initialization started");
			if (!PrefsExistAndNotNull(PrefPath)) {
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
			if (PrefPathIsJson(prefPath))
				return true;
			if (DeprecatedPrefsAvailable()) {
				return ConvertRegPrefToFile(prefPath);
			}
			return false;
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
            //Program.LogVerbose("Creating pref file");
			bool askForDir;
            string _installLocation = "";
            int _maxAttempts = 3;
            string _skippedVersion = "";
            Program.LogVerbose("Decision to ask for jarpath from user");
			string userdir;
			if (Environment.OSVersion.Platform == PlatformID.Win32NT) {
				userdir = Environment.GetEnvironmentVariable("UserProfile") ?? "";
			} else {
				userdir = Environment.GetEnvironmentVariable("HOME") ?? "";
			}
			//THIS WILL FAIL ON UNIX ENVIRONMENTS WITHOUT MONO
            OpenFileDialog dialog = new() {
                InitialDirectory = userdir,
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
				_installLocation = selectedJar.Directory!.FullName;
                Program.LogVerbose($"InstallPath: {_installLocation};");
                askForDir = false;
			} else {
				askForDir = true;
			}
			dialog.Dispose();
			if (askForDir) {
                Program.LogVerbose("Location of JAR was requested, but not provided. Asking for dir to install new jar");
                string? installPath = AskForDir();
				if (installPath is null) {
                    Program.LogVerbose("User cancelled out of request. No further action possible.");
                    Environment.Exit(0);
				}
                Program.LogVerbose("Path identified, saving to prefkey");
				_installLocation = installPath;
                Program.LogVerbose($"InstallPath: {installPath};");
            }
            var preferences = new {
                InstallLocation = _installLocation,
                MaxAttempts = _maxAttempts,
                SkippedVersion = _skippedVersion,
            };
            string json = JsonSerializer.Serialize(preferences);
            File.WriteAllText(prefPath, json, System.Text.Encoding.UTF8);
        }
		private void LoadVals() {
			if (!(Preferences.PrefsExistAndNotNull(PrefPath))) {
				Preferences.InitPrefs(PrefPath);
			}
			string jsonString = File.ReadAllText(PrefPath, System.Text.Encoding.UTF8);
			var prefs = JsonSerializerExtensions.DeserializeAnonymousType(jsonString, new { InstallLocation = "", MaxAttempts = 0, SkippedVersion = "" });
			if (prefs is null) {
				throw new UnreachableException("Null prefs JSON after init");
			}
			InstallLocation = prefs.InstallLocation;
			MaxAttempts = prefs.MaxAttempts;
			_skippedVersion = prefs.SkippedVersion;
			if (Environment.OSVersion.Platform == PlatformID.Win32NT) {
                string AssemblyDir = AppContext.BaseDirectory;
                string[] Runtimes = Directory.GetFiles(AssemblyDir, "vcruntime*.dll");
                if (Runtimes.Length > 0) {
                    Standalone = true;
                } else {
                    Standalone = false;
                }
            } else {
				Standalone = true;
			}
		}
	}
}
