using HtmlAgilityPack;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Windows.Bits;

namespace Launch_KoLMafia {
	public static class MyExtensions {
		public static string Hash (this FileInfo file, HashAlgorithm cryptoService) {
			if (!file.Exists) return null;
			StringBuilder builder = new();
			using (cryptoService) {
				using (FileStream fileStream = file.Open(FileMode.Open)) {
					fileStream.Position = 0;
					byte[] bytes = cryptoService.ComputeHash(fileStream);
					foreach (byte b in bytes) {
						builder.Append(b.ToString("x2"));
					}
				}
			}
			return builder.ToString().ToLower();
		}
		public static string GetFirstMatchingDescendent(this HtmlNode node, string ElementType, string Pattern) {
            foreach (HtmlNode dNode in node.Descendants(ElementType)) {
				if (dNode.NodeType == HtmlNodeType.Element && Regex.IsMatch(dNode.InnerHtml, Pattern, RegexOptions.IgnoreCase)) {
                    return dNode.InnerHtml;
                }
            }
			return null;
        }
		public static HtmlNode GetUriBody(this Uri uri) {
            HtmlWeb web = new();
            HtmlAgilityPack.HtmlDocument htmlDoc = web.Load(uri);
            return htmlDoc.DocumentNode.SelectSingleNode("//body");
        }
	}
	internal class Preferences {
		public string PrefPath { get; set; }
		public string InstallLocation { get; set; }
		public int MaxAttempts { get; set; }
		public bool Silent { get; set; }
		public string SkippedVersion { get; set; }
		public Preferences(string prefPath) {
			this.Initialize(prefPath, false);
		}
		public Preferences(string prefPath, bool silent) {
			this.Initialize(prefPath, silent);
		}
		private void Initialize(string prefPath, bool silent) {
			this.PrefPath = prefPath;
			if (!(Preferences.PrefsExistAndNotNull(prefPath)) && silent) {
				throw new ArgumentException("No registry config found, silence precludes requesting install location");
			} else if (!(Preferences.PrefsExistAndNotNull(prefPath)) && !silent) {
				Preferences.InitPrefs(prefPath);
			}
			this.Silent = silent;
			this.LoadVals();
		}
		private static bool PrefsExistAndNotNull(string prefPath) {
			RegistryKey prefKey = Registry.CurrentUser.OpenSubKey(prefPath);
			if (prefKey == null) { return false; }
			object KoLPath = prefKey.GetValue("PathToKoL", null);
			if (KoLPath == null) { return false; }
			if (string.IsNullOrEmpty(KoLPath.ToString())) { return false; }
			return true;
		}
		private static void InitPrefs(string prefPath) {
			RegistryKey prefKey = Registry.CurrentUser.CreateSubKey(prefPath);
			bool askForDir;
			if (string.IsNullOrEmpty(prefKey.GetValue("PathToKoL", null).ToString())) {
				using OpenFileDialog dialog = new();
				dialog.InitialDirectory = Environment.GetEnvironmentVariable("UserProfile");
				dialog.Filter = "JAR files (*.jar)| *.jar";
				dialog.Title = "Select Location of KolMafia JAR";
				if (dialog.ShowDialog() == DialogResult.OK) {
					FileInfo selectedJar = new(dialog.FileName);
					string installPath = selectedJar.Directory.Name;
					prefKey.SetValue("PathToKoL", installPath, RegistryValueKind.String);
					askForDir = false;
				} else {
					askForDir = true;
				}
			} else {
				askForDir = false;
			}
			if (askForDir) {
				string msg = "No file was selected. Should Launcher assume you don't have KoLMafia installed and ask where you want it? (Selecting No will exit immediately)";
				string title = "No file selected";
				MessageBoxButtons buttons = MessageBoxButtons.YesNo;
				DialogResult nobutton = DialogResult.No;
				MessageBoxIcon question = MessageBoxIcon.Question;
				MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button1;
				DialogResult answer = MessageBox.Show(msg, title, buttons, question, defaultButton);
				if (answer == nobutton) {
					Environment.Exit(0);
				} else {
					using FolderBrowserDialog selectFolder = new();
					selectFolder.SelectedPath = Environment.GetEnvironmentVariable("UserProfile");
					selectFolder.Description = "Select where you want to install KoLMafia";
					selectFolder.ShowNewFolderButton = true;
					DialogResult result = selectFolder.ShowDialog();
					if (result == DialogResult.Cancel) {
						Environment.Exit(0);
					} else {
						string installPath = selectFolder.SelectedPath;
						prefKey.SetValue("PathToKoL", installPath, RegistryValueKind.String);
					}
				}
			}
			if (prefKey.GetValue("MaxDownloadAttempts",null) == null) {
				prefKey.SetValue("MaxDownloadAttempts", 3, RegistryValueKind.DWord);
			}
			if (prefKey.GetValue("SkippedVersion",null) == null) {
				prefKey.SetValue("SkippedVersion", "", RegistryValueKind.String);
			}
		}
		private void LoadVals() {
			if (!(Preferences.PrefsExistAndNotNull(this.PrefPath))) {
				Preferences.InitPrefs(this.PrefPath);
			}
			RegistryKey prefKey = Registry.CurrentUser.OpenSubKey(this.PrefPath);
			this.InstallLocation = prefKey.GetValue("PathToKoL").ToString();
			this.MaxAttempts = (int)prefKey.GetValue("MaxDownloadAttempts");
			if (prefKey.GetValue("SkippedVersion",null) == null) {
				prefKey.SetValue("SkippedVersion", "", RegistryValueKind.String);
			}
			this.SkippedVersion = prefKey.GetValue("SkippedVersion").ToString();
		}
	}
    internal static class Program {
		private const string KoLBaseLocation = @"https://builds.kolmafia.us/job/Kolmafia/lastSuccessfulBuild/artifact/dist/";
		static readonly HttpClient client = new();
		static Hashtable GetShellOpenFromExtension(string Extension) {
			string extensionPattern = /* language=regex */ @"\.?(?<Extension>[a-z0-9]{1,}$)";
			string commandPattern = /* language=regex */ @"(?(^"")(?<path>""[^""]*"")|(?<path>[^ ]*)) (?<params>.*)";
			MatchCollection matches = Regex.Matches(Extension, extensionPattern, RegexOptions.IgnoreCase);
			if (matches.Count == 0) {
				return null;
			}
			Extension = "." + matches[0].Groups["Extension"].Value;

			string RegisteredApplication = Registry.ClassesRoot.OpenSubKey(Extension).GetValue("").ToString();
			if (RegisteredApplication == null) { return null; }

			string ShellOpen = Registry.ClassesRoot.OpenSubKey(RegisteredApplication + @"\shell\open\command").GetValue("").ToString();
			if (ShellOpen == null) { return null; }

			matches = Regex.Matches(ShellOpen, commandPattern, RegexOptions.IgnoreCase);
			if (matches.Count > 0) {
				Hashtable command = new();
				foreach (Match match in matches) {
					GroupCollection group = match.Groups;
					command.Add("appPath", group["path"]);
					command.Add("arguments", group["params"]);
				}
				return command;
			} else {
				return null;
			}
		}
		static bool GetWebFile(Uri URI, FileInfo Destination, DownloadPriority Priority, string Fingerprint, HashAlgorithm cryptoService) {
			DownloadManager download = new();
			IDownloadJob job = download.CreateJob("BITS Download", URI.AbsoluteUri, Destination.FullName, Priority);
			job.Resume();
			bool jobIsFinal = false;
			while (!jobIsFinal) {
				DownloadStatus status = job.Status;
				switch (status) {
					case DownloadStatus.Error:
					case DownloadStatus.Transferred:
						job.Complete();
						break;
					case DownloadStatus.Cancelled:
					case DownloadStatus.Acknowledged:
						jobIsFinal = true; 
						break;
					default:
						Task.Delay(500);
						break;
				}
			}
			if (Fingerprint != null) {
				return Destination.Hash(cryptoService) == Fingerprint;
			} else return true;
		}
		static string CheckForUpdate(string SkippedVersion) {
			string installationKey = @"SOFTWARE\WOW6432Node\Sapph Tools\KoLMafia Launcher";
			string currentVersion = Registry.LocalMachine.OpenSubKey(installationKey).GetValue("Version").ToString();
			Uri versionURI = new("https://raw.githubusercontent.com/sapph42/KoLMafia-Launcher/main/version.txt");
			string releaseVersion;
			currentVersion ??= System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
			try {
				var webRequest = new HttpRequestMessage(HttpMethod.Get, versionURI);
				var response = client.Send(webRequest);
				releaseVersion = response.Content.ReadAsStringAsync().Result.Trim();
			} catch (HttpRequestException e) {
				Console.WriteLine("\nException Caught!");
				Console.WriteLine(e.Message);
				return "";
			}
			if (releaseVersion == currentVersion || releaseVersion == SkippedVersion) {
				return "";
			}
			string msg = $"An updated version({releaseVersion}) of KoLMafia Launcher is available! Update now? (Or skip this release ?)";
			string title = "Update available!";
			MessageBoxButtons buttons = MessageBoxButtons.YesNo;
			DialogResult nobutton = DialogResult.No;
			MessageBoxIcon icon = MessageBoxIcon.Question;
			MessageBoxDefaultButton defaultbutton = MessageBoxDefaultButton.Button1;
			DialogResult result = MessageBox.Show(msg, title, buttons, icon, defaultbutton);
			if (result == nobutton) {
				return releaseVersion;
			}
			string targetInstallerName = $"KoLMafia-Launcher_{releaseVersion}.exe";
			Uri sourceURI = new($"https://github.com/sapph42/KoLMafia-Launcher/raw/main/{targetInstallerName}");
			FileInfo destination = new(Environment.GetEnvironmentVariable("Temp") + @"\" + targetInstallerName);
			try {
				if (GetWebFile(sourceURI, destination, DownloadPriority.Foreground, null, null)) {
					ProcessStartInfo processStartInfo = new() {
						FileName = destination.FullName,
						WorkingDirectory = Environment.GetEnvironmentVariable("Temp"),
						ErrorDialog = true,
						UseShellExecute = true,
						Verb = "runas"
					};
					System.Diagnostics.Process.Start(processStartInfo);
					Environment.Exit(0);
					return "";
				}
			} catch {
				return "";
			}
			return "";
		}
		static bool KillAllProcByName(string ProcName) {
			try {
				Process[] processList = Process.GetProcessesByName(ProcName);
				foreach (Process process in processList) {
					process.Kill();
				}
				System.Threading.Thread.Sleep(200);
				return (Process.GetProcessesByName(ProcName).Length == 0);
			} catch {
				return false;
			}
		}
		static void Main(string[] args) {
			bool noLaunch = false;
			bool killOnUpdate = false;
			bool silent = false;
			bool exists = false;
			string javaPath;
			string javaParams;
			string javaName;
			string latestJARName = "";
			string canonicalFingerprint = null;
			Uri KoLBaseLocation = new(Program.KoLBaseLocation);
            Uri jarURI;
            Uri fingerprintURI;
            string[] currentList;
			Hashtable command;
			HashAlgorithm cryptoService = MD5.Create();
			FileInfo currentFile = null;
			FileInfo latestFile = null;

			if (args.Length != 0) {
				noLaunch = args.Contains("--noLaunch", StringComparer.CurrentCultureIgnoreCase);
				killOnUpdate = args.Contains("--killOnUpdate", StringComparer.CurrentCultureIgnoreCase);
				silent = args.Contains("--silent", StringComparer.CurrentCultureIgnoreCase);
			}
			Preferences preferences = new(@"Software\Sapph Tools\KoLMafia Launcher\", silent);
			if (!silent) {
				string releaseVer = CheckForUpdate(preferences.SkippedVersion).ToString();
				if (releaseVer != "") {
					preferences.SkippedVersion = releaseVer;
				}
			}
			command = GetShellOpenFromExtension(".jar");
			if (command["appPath"] == null) {
				javaPath = "javaw.exe";
				javaParams = @"-jar ""%1""";
			} else {
				javaPath = command["appPath"].ToString();
				javaParams = command["arguments"].ToString();
			}
			javaName = Regex.Match(javaPath, @"(?<name>[^\\]*)(?:\.exe)", RegexOptions.IgnoreCase).Groups["name"].Value;
			if (killOnUpdate) {
				if (!KillAllProcByName(javaName)) {
					Environment.Exit(42);
				}
			}
			if (Process.GetProcessesByName("javaw").Length > 0) {
				if (silent) {
					Environment.Exit(43);
				}
				string msg = "A javaw.exe process has been detected. For safety, update cannot continue without killing this process.";
				string title = "Java Interpreter Already Running";
				MessageBoxButtons buttons = MessageBoxButtons.OKCancel;
				DialogResult cancel = DialogResult.Cancel;
				MessageBoxIcon icon = MessageBoxIcon.Warning;
				MessageBoxDefaultButton defaultbutton = MessageBoxDefaultButton.Button2;
				DialogResult answer = MessageBox.Show(msg, title, buttons, icon, defaultbutton);
				if (answer == cancel) {
					Environment.Exit(1);
				} else {
					if (!KillAllProcByName(javaName)) {
						Environment.Exit(42);
					}
				}
			}

			currentList = System.IO.Directory.GetFiles(preferences.InstallLocation, "*.jar");

			if (currentList.Length > 1) {
				Array.Sort(currentList, StringComparer.OrdinalIgnoreCase);
				for (int i = 0; i < currentList.Length - 1; i++) {
					System.IO.File.Delete(currentList[i]);
				}
				currentFile = new(currentList[^1]);
				exists = true;
			} else if (currentList.Length == 0) {
				if (silent) {
					exists = false; 
				} else {
					string msg = $"No jar file found in the provided folder. Download latest mafia to {preferences.InstallLocation}?";
					string title = "Mafia Not Found!";
					MessageBoxButtons buttons = MessageBoxButtons.YesNo;
					DialogResult nobutton = DialogResult.No;
					MessageBoxIcon icon = MessageBoxIcon.Question;
					MessageBoxDefaultButton defaultbutton = MessageBoxDefaultButton.Button1;
					DialogResult answer = MessageBox.Show(msg, title, buttons, icon, defaultbutton);
					if (answer == nobutton) {
						Environment.Exit(0);
					} else {
						exists = false;
					}
				}
			} else {
				exists = true;
				currentFile = new(currentList[0]);
			}
			try {
				latestJARName = KoLBaseLocation.GetUriBody().GetFirstMatchingDescendent("a", @"\.jar$");

				jarURI = new(KoLBaseLocation.AbsoluteUri + latestJARName);
				fingerprintURI = new(jarURI.AbsoluteUri + @"/*fingerprint*/");

				canonicalFingerprint = fingerprintURI.GetUriBody().GetFirstMatchingDescendent("li", @"[a-z0-9]{32}");

				if (!exists || (currentFile.Name != latestJARName) || (currentFile.Hash(cryptoService) != canonicalFingerprint)) {
					int attempts = 1;
					FileInfo destination = new(preferences.InstallLocation + @"\" + latestJARName);
					bool downloadSuccess = GetWebFile(jarURI, destination, DownloadPriority.Foreground, canonicalFingerprint, cryptoService);
					while (!downloadSuccess && attempts < preferences.MaxAttempts) {
						if (destination.Exists) { destination.Delete(); }
						downloadSuccess = GetWebFile(jarURI, destination, DownloadPriority.Foreground, canonicalFingerprint, cryptoService);
						attempts++;
					}
					if (exists && downloadSuccess) {
						currentFile.Delete();
						latestFile = destination;
					}
				} else {
					latestFile = currentFile;
				}
			} catch (HttpRequestException e) {
				Console.WriteLine("\nException Caught!");
				Console.WriteLine(e.Message);
				if (!silent) {
					MessageBox.Show(@"Received a error when attempting to retreive and/or save the latest version. Your previous version has been kept, and will now be run.");
					latestFile = currentFile;
				}
			} finally {
				if (!noLaunch) {
					ProcessStartInfo processStartInfo = new() {
						FileName = javaPath,
						Arguments = javaParams.Replace(@"%1", latestFile.FullName).Replace(@" %*",""),
						WorkingDirectory = latestFile.Directory.FullName
					};
					System.Diagnostics.Process.Start(processStartInfo);
				}
			}
		}
	}
}
