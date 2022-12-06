#nullable enable
using HtmlAgilityPack;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Resources;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Bits;

[assembly: NeutralResourcesLanguageAttribute("en-US")]


namespace Launch_KoLMafia {
	public static class MyExtensions {
		public static string Hash (this FileInfo file, 
									HashAlgorithm cryptoService) {
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
		public static string? GetFirstMatchingDescendent(this HtmlNode node, 
														string ElementType, 
														[StringSyntax(StringSyntaxAttribute.Regex)] string Pattern) {
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
#nullable disable
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
				throw new ArgumentException(Properties.Resources.SilentEmptyRegistryError);
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
				dialog.Title = Properties.Resources.JarFileSelectTitle;
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
				string msg = Properties.Resources.TargetFolderDialogMsg;
				string title = Properties.Resources.TargetFolderDialogTitle;
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
					selectFolder.Description = Properties.Resources.FolderSelectTitle;
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
			RegistryKey prefKey = Registry.CurrentUser.OpenSubKey(PrefPath);
			this.InstallLocation = prefKey.GetValue("PathToKoL").ToString();
			this.MaxAttempts = (int)prefKey.GetValue("MaxDownloadAttempts");
			if (prefKey.GetValue("SkippedVersion",null) == null) {
				prefKey.SetValue("SkippedVersion", "", RegistryValueKind.String);
			}
			this.SkippedVersion = prefKey.GetValue("SkippedVersion").ToString();
		}
#nullable enable
	}
	internal static class Program {
		[DllImport("Shlwapi.dll", CharSet = CharSet.Unicode)]
		public static extern uint AssocQueryString(
			AssocF flags,
			AssocStr str,
			string pszAssoc,
			string? pszExtra,
			[Out] StringBuilder? pszOut,
			ref uint pcchOut
		);
		[Flags]
		public enum AssocF {
			None = 0,
			Init_NoRemapCLSID = 0x1,
			Init_ByExeName = 0x2,
			Open_ByExeName = 0x2,
			Init_DefaultToStar = 0x4,
			Init_DefaultToFolder = 0x8,
			NoUserSettings = 0x10,
			NoTruncate = 0x20,
			Verify = 0x40,
			RemapRunDll = 0x80,
			NoFixUps = 0x100,
			IgnoreBaseClass = 0x200,
			Init_IgnoreUnknown = 0x400,
			Init_Fixed_ProgId = 0x800,
			Is_Protocol = 0x1000,
			Init_For_File = 0x2000
		}
		public enum AssocStr {
			Command = 1,
			Executable,
			FriendlyDocName,
			FriendlyAppName,
			NoOpen,
			ShellNewValue,
			DDECommand,
			DDEIfExec,
			DDEApplication,
			DDETopic,
			InfoTip,
			QuickTip,
			TileInfo,
			ContentType,
			DefaultIcon,
			ShellExtension,
			DropTarget,
			DelegateExecute,
			Supported_Uri_Protocols,
			ProgID,
			AppID,
			AppPublisher,
			AppIconReference,
			Max
		}
		private const string KoLBaseLocation = @"https://builds.kolmafia.us/job/Kolmafia/lastSuccessfulBuild/artifact/dist/";
		static readonly HttpClient client = new();
		static string? AssocQueryString(AssocStr association, string extension) {
			const int S_OK = 0;
			const int S_FALSE = 1;

			uint length = 0;
			uint ret = AssocQueryString(AssocF.None, association, extension, null, null, ref length);
			if (ret != S_FALSE) {
			    return null;
			}

			var sb = new StringBuilder((int)length); // (length-1) will probably work too as the marshaller adds null termination
			ret = AssocQueryString(AssocF.None, association, extension, null, sb, ref length);
			if (ret != S_OK) {
			    return null;
			}

			return sb.ToString();
		}
		static bool GetWebFile(Uri URI, 
								FileInfo Destination, 
								DownloadPriority Priority, 
								string Fingerprint, 
								HashAlgorithm cryptoService) {
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
			string msg = string.Format(Properties.Resources.NewVersionDialogMsg, releaseVersion);
			string title = Properties.Resources.NewVersionDialogTitle;
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
			javaName = AssocQueryString(AssocStr.DDEApplication, ".jar");
			if (killOnUpdate) {
				if (!KillAllProcByName(javaName)) {
					Environment.Exit(42);
				}
			}
			if (Process.GetProcessesByName(javaName).Length > 0) {
				if (silent) {
					Environment.Exit(43);
				}
				string msg = string.Format(Properties.Resources.ProcessConflictDialogMsg, javaName);
				string title = Properties.Resources.ProcessConflictDialogTitle;
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
					string msg = string.Format(Properties.Resources.DownloadToNewConfirmMsg, preferences.InstallLocation);
					string title = Properties.Resources.DownloadToNewConfirmTitle;
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
					MessageBox.Show(Properties.Resources.RetreivalError);
					latestFile = currentFile;
				}
			} finally {
				if (!noLaunch) {
					ProcessStartInfo processStartInfo = new() {
						FileName = latestFile.FullName,
						UseShellExecute= true,
						WorkingDirectory = latestFile.Directory.FullName
					};
					System.Diagnostics.Process.Start(processStartInfo);
				}
			}
		}
	}
}
