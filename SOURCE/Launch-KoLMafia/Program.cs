using HtmlAgilityPack;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Resources;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Bits;

[assembly: NeutralResourcesLanguageAttribute("en-US")]
namespace Launch_KoLMafia {
	internal static class Program {
        private const string KoLBaseLocation = @"https://builds.kolmafia.us/job/Kolmafia/lastSuccessfulBuild/artifact/dist/";
        // private const string KoLAPI = @"https://ci.kolmafia.us/job/Kolmafia/lastSuccessfulBuild/api/json";
        // artifactJSON.RootElement.GetProperty("artifacts")[0].GetProperty("relativePath").ToString();
        static readonly HttpClient client = new();
        
		[return: NotNull]
		static bool GetWebFile(Uri URI, 
			FileInfo Destination, 
			DownloadPriority Priority, 
			string? Fingerprint, 
			HashAlgorithm? cryptoService
		) {
			DownloadManager download = new();
			IDownloadJob job = download.CreateJob("BITS Download", URI.AbsoluteUri, Destination.FullName, Priority);
			job.NoProgressTimeout = 20;
			job.Resume();
			bool jobIsFinal = false;
			while (!jobIsFinal) {
				DownloadStatus status = job.Status;
				switch (status) {
					case DownloadStatus.Error:
					case DownloadStatus.Suspended:
						job.Cancel();
						return false;
					case DownloadStatus.Transferred:
						job.Complete();
						jobIsFinal = true;
						break;
					case DownloadStatus.Cancelled:
						return false;
					case DownloadStatus.Acknowledged:
						jobIsFinal = true; 
						break;
					default:
						Task.Delay(500);
						break;
				}
			}
			if (Fingerprint is not null && cryptoService is not null) {
				return Destination.Hash(cryptoService) == Fingerprint;
			} else return true;
		}
		[return: NotNull]
		static string CheckForUpdate(string? SkippedVersion) {
			SkippedVersion ??= "";
			RegistryKey? installationKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Sapph Tools\KoLMafia Launcher");
			string currentVersion = null!;
			Uri versionURI = new("https://raw.githubusercontent.com/sapph42/KoLMafia-Launcher/main/version.txt");
			string releaseVersion;
			Version? version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
			if (installationKey is not null) {
			    object versionKey = installationKey.GetValue("Version", "");
			    currentVersion = ("" ?? versionKey.ToString());
			}
			if (version is not null && string.IsNullOrEmpty(currentVersion)) {
				currentVersion = version.ToString();
			}
			try {
				HttpRequestMessage webRequest = new(HttpMethod.Get, versionURI);
				HttpResponseMessage response = client.Send(webRequest);
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
				}
			} catch {} 
			return "";
		}
		[return: NotNull]
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

			Application.ThreadException += new ThreadExceptionEventHandler(Application_ThreadException);
			AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

			bool noLaunch = false;
			bool killOnUpdate = false;
			bool silent = false;
			bool exists = false;
			string javaName;
			string latestJARName = "";
			string canonicalFingerprint = null!;
			Uri KoLBaseLocation = new(Program.KoLBaseLocation);
			Uri jarURI;
			Uri fingerprintURI;
			string[] currentList;
			HashAlgorithm cryptoService = MD5.Create();
			FileInfo currentFile = null!;
			FileInfo latestFile = null!;

			if (args.Length != 0) {
				noLaunch = args.Contains("--noLaunch", StringComparer.CurrentCultureIgnoreCase);
				killOnUpdate = args.Contains("--killOnUpdate", StringComparer.CurrentCultureIgnoreCase);
				silent = args.Contains("--silent", StringComparer.CurrentCultureIgnoreCase);
			}
			Preferences preferences = new("""Software\Sapph Tools\KoLMafia Launcher\""", silent);
			if (!silent) {
				string releaseVer = CheckForUpdate(preferences.SkippedVersion).ToString();
				if (releaseVer != "") {
					preferences.SkippedVersion = releaseVer;
				}
			}
			javaName = WinAPI.AssocQueryString(WinAPI.AssocStr.DDEApplication, ".jar") ?? "";
			if (javaName == "") { killOnUpdate = false; }
			if (killOnUpdate) {
				if (!KillAllProcByName(javaName)) {
					Environment.Exit(42);
				}
			}
			if (javaName != "" && Process.GetProcessesByName(javaName).Length > 0) {
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
				HtmlNode? body = KoLBaseLocation.GetUriBody();
				latestJARName = body.GetFirstMatchingDescendent("a", @"\.jar$")! ?? "";
				if (latestJARName == "") {
					throw new ExternalException("New jar name not found at KoLMafia site.  Parse error.");
				}
				jarURI = new(KoLBaseLocation.AbsoluteUri + latestJARName);
				fingerprintURI = new(jarURI.AbsoluteUri + @"/*fingerprint*/");

				canonicalFingerprint = fingerprintURI.GetUriBody().GetFirstMatchingDescendent("li", @"[a-z0-9]{32}") ?? "";

				if (!exists || 
					(currentFile.Name != latestJARName) || 
					(canonicalFingerprint != "" && 
						(currentFile.Hash(cryptoService) != canonicalFingerprint))) {
					int attempts = 1;
					FileInfo destination = new(preferences.InstallLocation + @"\" + latestJARName);
					bool downloadSuccess = GetWebFile(jarURI, destination, DownloadPriority.Foreground, canonicalFingerprint, cryptoService);
					while (!downloadSuccess && attempts < preferences.MaxAttempts) {
						if (destination.Exists) { destination.Delete(); }
						downloadSuccess = GetWebFile(jarURI, destination, DownloadPriority.Foreground, canonicalFingerprint, cryptoService);
						downloadSuccess = downloadSuccess && 
											destination.Exists && 
											currentFile.Hash(cryptoService) == destination.Hash(cryptoService);
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
				if (!silent && currentFile.Exists) {
					MessageBox.Show(Properties.Resources.RetreivalError);
					latestFile = currentFile;
				}
			} catch (ExternalException e) {
			    Console.WriteLine("\nException Caught!");
			    Console.WriteLine(e.Message);
			    if (!silent && currentFile.Exists) {
			        MessageBox.Show(Properties.Resources.RetreivalError);
			        latestFile = currentFile;
			    }
			} finally {
				if (!noLaunch) {
					ProcessStartInfo processStartInfo = new() {
						FileName = latestFile.FullName,
						UseShellExecute= true,
						WorkingDirectory = latestFile.Directory!.FullName
					};
					System.Diagnostics.Process.Start(processStartInfo);
				}
			}
		}

		static void Application_ThreadException(object sender, ThreadExceptionEventArgs e) {
			// Log the exception, display it, etc
			Debug.WriteLine(e.Exception.Message);
			Console.WriteLine(e.Exception.Message+ Environment.NewLine);
			Console.WriteLine(e.Exception.InnerException);
			Console.WriteLine(e.Exception.StackTrace);
		}

		static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e) {
			Exception ex;
			if (e is not null && e.ExceptionObject is not null) {
				ex = (Exception)e.ExceptionObject;
			} else {
				return;
			}
			Debug.WriteLine(ex.Message);
			Console.WriteLine(ex.Message + Environment.NewLine);
			Console.WriteLine(ex.InnerException);
			Console.WriteLine(ex.StackTrace);
		}


	}
}
