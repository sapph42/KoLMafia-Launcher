using HtmlAgilityPack;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Resources;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
#if OS_WINDOWS
	using System.Windows.Forms;
#endif

[assembly: NeutralResourcesLanguageAttribute("en-US")]
namespace Launch_KoLMafia {
	internal static class Program {
		private const string KoLBaseLocation = @"https://api.github.com/repos/kolmafia/kolmafia/releases";
		// private const string KoLAPI = @"https://ci.kolmafia.us/job/Kolmafia/lastSuccessfulBuild/api/json";
		// artifactJSON.RootElement.GetProperty("artifacts")[0].GetProperty("relativePath").ToString();
		static readonly HttpClient client = new();
		static readonly HashAlgorithm cryptoService = MD5.Create();
		private static Preferences preferences = null!;

		public static bool verbose = false;

		[return: NotNull]
		private static bool IsInternetAvailable() {
			const string NCSI_TEST_URL = @"http://www.msftncsi.com/ncsi.txt";
			const string NCSI_TEST_RESULT = @"Microsoft NCSI";
			const string NCSI_DNS = @"dns.msftncsi.com";
			const string NCSI_DNS_IP_ADDRESS = @"131.107.255.255";

			try {
				HttpRequestMessage webRequest = new(HttpMethod.Get, NCSI_TEST_URL);
				HttpResponseMessage response = client.Send(webRequest);
				string result = response.Content.ReadAsStringAsync().Result;
				if (result != NCSI_TEST_RESULT) return false;
				IPHostEntry dnsHost = Dns.GetHostEntry(NCSI_DNS);
				if (dnsHost.AddressList.Length < 0 || dnsHost.AddressList[0].ToString() != NCSI_DNS_IP_ADDRESS) return false;
			} catch (Exception ex) {
				Console.WriteLine(ex); 
				return false;
			}
			return true;
		}

		[return: NotNull]
		static void DownloadFileAsync(Uri uri, FileInfo Destination) {
			HttpRequestMessage webRequest = new(HttpMethod.Get, uri);
			HttpResponseMessage response = client.Send(webRequest);
			byte[] fileBytes = response.Content.ReadAsByteArrayAsync().Result;
			File.WriteAllBytes(Destination.FullName, fileBytes);
		}
		static int GetWebFile(Uri URI, 
			FileInfo Destination, 
			string? Fingerprint
		) {
			DownloadFileAsync(URI, Destination);
			if (!Destination.Exists) { return 500; }
			if (Fingerprint is not null && cryptoService is not null) {
				string destinationHash = Destination.Hash(cryptoService)!;
				if (destinationHash == Fingerprint) {
					return 200;
				} else {
					return 404;
				}
			} else return 200;
		}
		[return: NotNull]
		static string CheckForUpdate(string? RegSkippedVersion) {
			Version? skippedVersion;
			if (RegSkippedVersion is null) {
				skippedVersion = new Version("0.0.0.0");
			} else if (!Version.TryParse(RegSkippedVersion, out skippedVersion)) {
				skippedVersion = new Version("0.0.0.0");
			} 
			RegistryKey? installationKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Sapph Tools\KoLMafia Launcher");
			Uri versionURI = new("https://raw.githubusercontent.com/sapph42/KoLMafia-Launcher/main/version.txt");
			Version? releaseVersion;
			Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version!;
			try {
				HttpRequestMessage webRequest = new(HttpMethod.Get, versionURI);
				HttpResponseMessage response = client.Send(webRequest);
				if (!Version.TryParse(response.Content.ReadAsStringAsync().Result.Trim(), out releaseVersion)) {
					Console.WriteLine("\nException Caught!");
					return "";
				}
			} catch (HttpRequestException e) {
				if (e.InnerException is not null && e.InnerException is System.Net.Sockets.SocketException) {
					if (!preferences.Silent) {
						MessageBox.Show("It appears there is no network connection. Launcher will now terminate.");
					}
					Environment.Exit(e.InnerException.HResult);
				}
				Console.WriteLine("\nException Caught!");
				Console.WriteLine(e.Message);
				return "";
			}
			if (releaseVersion.Equals(version) || releaseVersion.Equals(skippedVersion) || version.IsGreater(releaseVersion)) {
                ConsoleOut("Launch-KoLMafia is up-to-date!", ConsoleColor.Black, ConsoleColor.Green); 
				return "";
			}
			string msg = string.Format(Properties.Resources.NewVersionDialogMsg, releaseVersion.ToString());
			string title = Properties.Resources.NewVersionDialogTitle;
			MessageBoxButtons buttons = MessageBoxButtons.YesNo;
			DialogResult nobutton = DialogResult.No;
			MessageBoxIcon icon = MessageBoxIcon.Question;
			MessageBoxDefaultButton defaultbutton = MessageBoxDefaultButton.Button1;
			DialogResult result = MessageBox.Show(msg, title, buttons, icon, defaultbutton);
			if (result == nobutton) {
				return releaseVersion.ToString();
			}
			string targetInstallerName = $"KoLMafia-Launcher_{releaseVersion}.exe";
			if (preferences.Standalone) {
				targetInstallerName = $"KoLMafia-Launcher_standalone_{releaseVersion}.exe";
			}
			
			Uri sourceURI = new($"https://github.com/sapph42/KoLMafia-Launcher/raw/main/{targetInstallerName}");
			FileInfo destination = new(Environment.GetEnvironmentVariable("Temp") + @"\" + targetInstallerName);
			try {
				if (GetWebFile(sourceURI, destination, null)==200) {
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
		public static void LogVerbose(string log) {
			if (verbose)
				Console.WriteLine(log);
		}
		public static void LogVerbose(string[] logs) {
			foreach(string log in logs) {
				LogVerbose(log);
			}
		}
		public static void ConsoleOut(string log, ConsoleColor back = ConsoleColor.Black, ConsoleColor fore = ConsoleColor.White, bool newline = true) {
			Console.BackgroundColor = back;
			Console.ForegroundColor = fore;
			if (newline) {
				Console.WriteLine(log);
			} else {
				Console.Write(log);
			}
			Console.BackgroundColor = ConsoleColor.Black;
			Console.ForegroundColor = ConsoleColor.White;
		}
		[STAThread]
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
			FileInfo currentFile = null!;
			FileInfo latestFile = null!;

			if (args.Length != 0) {
				noLaunch = args.Contains("--noLaunch", StringComparer.CurrentCultureIgnoreCase);
				killOnUpdate = args.Contains("--killOnUpdate", StringComparer.CurrentCultureIgnoreCase);
				silent = args.Contains("--silent", StringComparer.CurrentCultureIgnoreCase);
				verbose = args.Contains("--verbose", StringComparer.CurrentCultureIgnoreCase) ;
			}
			LogVerbose("Arguments parsed.");
			ConsoleOut($"INFO: noLaunch:{noLaunch}; killOnUpdate:{killOnUpdate}; silent:{silent}; verbose:{verbose}",ConsoleColor.Black, ConsoleColor.DarkYellow);
			if (!IsInternetAvailable()) {
				if (!silent) {
					MessageBox.Show("It appears there is no network connection. Launcher will now terminate.");
				}
				Environment.Exit(1);
			}
			LogVerbose("Internet connection confirmed.");

			preferences = new(silent);
			LogVerbose("Preference object initialized");
			if (!silent) {
				string releaseVer = CheckForUpdate(preferences.SkippedVersion).ToString();
				if (releaseVer != "") {
					preferences.SkippedVersion = releaseVer;
				}
			}
			if (Environment.OSVersion.Platform == PlatformID.Win32NT) {
				javaName = WinAPI.AssocQueryString(WinAPI.AssocStr.DDEApplication, ".jar") ?? "";
			} else {
                string mimeType = "application/java-archive";
                Process process = new();
                process.StartInfo.FileName = "xdg-mime";
                process.StartInfo.Arguments = $"query default {mimeType}";
                process.StartInfo.RedirectStandardOutput = true;
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                javaName = output.Trim();
            }
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
			LogVerbose($"Retreived Install Location: {preferences.InstallLocation}");
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
                //HtmlNode? body = KoLBaseLocation.GetUriBody();
                //latestJARName = body.GetFirstMatchingDescendent("a", @"\.jar$")! ?? "";
                string assetID = "";
                var client = new HttpClient();
                var request = new HttpRequestMessage(HttpMethod.Get, "https://api.github.com/repos/kolmafia/kolmafia/releases/latest");
                request.Headers.Add("Accept", @"application/json");
				request.Headers.Add("User-Agent", @"Launch-KolMafia/5.1.4.0");
				request.Headers.Add("Host", @"api.github.com");
                var response = Task.Run<HttpResponseMessage>(async () => await client.SendAsync(request)).Result ;
                response.EnsureSuccessStatusCode();
				var content = Task.Run<string>(async () => await response.Content.ReadAsStringAsync()).Result ;
                var apiResponse =
                    System.Text.Json.JsonDocument.Parse(content);
                for (int i = 0; i < apiResponse.RootElement.GetProperty("assets").GetArrayLength(); i++) {
                    var asset = apiResponse.RootElement.GetProperty("assets")[i];
                    if (Regex.IsMatch(asset.GetProperty("name").ToString(), @"\.jar$")) {
                        assetID = asset.GetProperty("id").ToString();
                        latestJARName = asset.GetProperty("name").ToString();
                        break;
                    }
                }


				if (!exists || 
					currentFile.Name != latestJARName) {
					int attempts = 1;
					FileInfo destination = new(preferences.InstallLocation + @"\" + latestJARName);
					ConsoleOut("New version of Mafia found! Download in progress!", ConsoleColor.Black, ConsoleColor.DarkYellow);

                    request = new HttpRequestMessage(HttpMethod.Get, $"https://api.github.com/repos/kolmafia/kolmafia/releases/assets/{assetID}");
                    request.Headers.Add("Accept", @"application/json");
                    request.Headers.Add("User-Agent", @"Launch-KolMafia/5.1.4.0");
                    request.Headers.Add("Host", @"api.github.com");
                    response = Task.Run<HttpResponseMessage>(async () => await client.SendAsync(request)).Result;
                    response.EnsureSuccessStatusCode();
                    content = Task.Run<string>(async () => await response.Content.ReadAsStringAsync()).Result;
                    apiResponse = System.Text.Json.JsonDocument.Parse(content);
                    jarURI = new Uri(apiResponse.RootElement.GetProperty("browser_download_url").ToString());

                    int downloadStatus = GetWebFile(jarURI, destination, canonicalFingerprint);
					bool downloadSuccess = downloadStatus == 200;
					while (downloadStatus != 200 && attempts < preferences.MaxAttempts) {
						if (destination.Exists) { destination.Delete(); }
						downloadStatus = GetWebFile(jarURI, destination, canonicalFingerprint);
						downloadSuccess = downloadStatus == 200 && 
											destination.Exists && 
											currentFile.Hash(cryptoService) == destination.Hash(cryptoService);
						attempts++;
					}
					if (downloadSuccess) {
						if (exists) currentFile.Delete();
						latestFile = destination;
						ConsoleOut("Download complete!", ConsoleColor.Black, ConsoleColor.Green);
					} else {
						if (!silent) {
							string title = Properties.Resources.RetreivalErrorTitle;
							string msg = Properties.Resources.RetreivalError;
							MessageBoxButtons buttons = MessageBoxButtons.OKCancel;
							MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button1;
							MessageBoxIcon messageBoxIcon = MessageBoxIcon.Error;
							DialogResult result = MessageBox.Show(msg, title, buttons, messageBoxIcon, defaultButton);
							if (result == DialogResult.Cancel) {
								MessageBox.Show($"Attempts: {attempts}\r\nSuccess: {downloadSuccess}\r\nExits: {exists}\r\nCF: {canonicalFingerprint}\r\nDownload Exists: {destination.Exists}\r\nDF: {destination.Hash(cryptoService) ?? "N/A"}", "Troubleshooting Info");
								if (destination.Exists) destination.Delete();
								Environment.Exit(0);
							} else {
								latestFile = currentFile;
							}
						} else {
							Environment.Exit(269);
						}
					}
				} else {
					ConsoleOut("KoLMafia already up-to-date.", ConsoleColor.Black, ConsoleColor.Green);
					latestFile = currentFile;
				}
				if (currentFile is null) {
					throw new UnreachableException("Heisenbug found determining current Mafia JAR and/or replacing it.");
				}
			} catch (HttpRequestException e) {
				Console.WriteLine("\nException Caught!");
				Console.WriteLine(e.Message);
				if (!silent && currentFile.Exists) {
					MessageBox.Show(Properties.Resources.RetreivalError);
					latestFile = currentFile;
				}
				if (currentFile is null) {
					throw new UnreachableException("Heisenbug found determining current Mafia JAR and/or replacing it.");
				}
			} catch (ExternalException e) {
				Console.WriteLine("\nException Caught!");
				Console.WriteLine(e.Message);
				if (!silent && currentFile.Exists) {
					MessageBox.Show(Properties.Resources.RetreivalError);
					latestFile = currentFile;
				}
				if (currentFile is null) {
					throw new UnreachableException("Heisenbug found determining current Mafia JAR and/or replacing it.");
				}
			} finally {
				if (!noLaunch) {
					ProcessStartInfo processStartInfo = new() {
						FileName = latestFile.FullName,
						UseShellExecute= true,
						WorkingDirectory = latestFile.Directory!.FullName
					};
                    ConsoleOut("Launching Mafia ", ConsoleColor.Black, ConsoleColor.Green, false);
                    Process? jar = Process.Start(processStartInfo);
					Thread.Sleep(500);
					if (Environment.OSVersion.Platform != PlatformID.Win32NT)
						Application.Exit();
					if (jar is null) {
                        ConsoleOut("JRE process failed to instantiate after 0.5 seconds", ConsoleColor.Black, ConsoleColor.Red);
						Thread.Sleep(2000);
						Application.Exit();
                    }
					if (jar!.WaitForInputIdle(10000)) {
						IntPtr handle = jar.MainWindowHandle;
						int maxloop = 20;
						int loopcount = 0;
						while (handle == (IntPtr)0 && loopcount < maxloop) {
                            ConsoleOut(".", ConsoleColor.Black, ConsoleColor.Green, false);
                            Thread.Sleep(500);
							handle = jar.MainWindowHandle;
							loopcount++;
						}
						if (loopcount == maxloop) {
							ConsoleOut("");
                            ConsoleOut("Mafia window not loaded after 10 seconds.", ConsoleColor.Black, ConsoleColor.Red);
                            Thread.Sleep(2000);
							Application.Exit();
                        }
						loopcount = 0;
						while (!WinAPI.IsWindow(handle) && loopcount < maxloop) {
                            ConsoleOut(".", ConsoleColor.Black, ConsoleColor.Green, false);
                            Thread.Sleep(100);
							loopcount++;
						}
                        if (loopcount == maxloop) {
                            ConsoleOut("");
                            ConsoleOut("Mafia window not visible after 2 seconds.", ConsoleColor.Black, ConsoleColor.Red);
                            Thread.Sleep(2000);
                            Application.Exit();
                        }
                        ConsoleOut("");
                        ConsoleOut("Launch complete!", ConsoleColor.Black, ConsoleColor.Green);
                        Thread.Sleep(1000);
                    } else {
                        ConsoleOut("");
                        ConsoleOut("Mafia JAR not loaded after 10 seconds.", ConsoleColor.Black, ConsoleColor.Red);
                        Thread.Sleep(2000);
                    }
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
