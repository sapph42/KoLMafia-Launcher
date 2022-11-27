using HtmlAgilityPack;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Bits;

namespace Launch_KoLMafia {
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
        static readonly HttpClient client = new();
        static Hashtable GetShellOpenFromExtension(string Extension) {
            if (!Regex.IsMatch(Extension, @"\.?(?<Extension>[a-z0-9]{1,3}$)", RegexOptions.IgnoreCase)) {
                return null;
            }
            Extension = "." + Regex.Matches(Extension, @"\.?(?<Extension>[a-z0-9]{1,3}$)", RegexOptions.IgnoreCase)[0].Groups["Extension"].Value;
            string RegisteredApplication = Registry.ClassesRoot.OpenSubKey(Extension).GetValue("").ToString();
            if (RegisteredApplication == null) { return null; }
            string ShellOpen = Registry.ClassesRoot.OpenSubKey(RegisteredApplication + @"\shell\open\command").GetValue("").ToString();
            if (ShellOpen == null) { return null; }
            string commandPattern = @"(?(^"")(?<path>""[^""]*"")|(?<path>[^ ]*)) (?<params>.*)";
            if (Regex.IsMatch(ShellOpen, commandPattern, RegexOptions.IgnoreCase)) {
                Hashtable command = new();
                MatchCollection matches = Regex.Matches(ShellOpen, commandPattern, RegexOptions.IgnoreCase);
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
        static bool GetWebFile(string URI, string Destination, DownloadPriority Priority, string Fingerprint, HashAlgorithm cryptoService) {
            DownloadManager download = new();
            IDownloadJob job = download.CreateJob("BITS Download", URI, Destination, Priority);
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
                string targetHash = GetLocalHash(Destination, cryptoService).ToLower();
                return targetHash == Fingerprint;
            } else return true;
        }
        static string GetLocalHash(string file, HashAlgorithm cryptoService) {
            StringBuilder builder = new();
            using (cryptoService) {
                var fileStream = new FileStream(file,
                                                FileMode.Open,
                                                FileAccess.Read,
                                                FileShare.ReadWrite);
                byte[] bytes = cryptoService.ComputeHash(fileStream);
                fileStream.Dispose();
                foreach (byte b in bytes) {
                    builder.Append(b.ToString("x2"));
                }
            }
            return builder.ToString();
        }
        static string CheckForUpdate(string SkippedVersion) {
            string installationKey = @"SOFTWARE\WOW6432Node\Sapph Tools\KoLMafia Launcher";
            string currentVersion = Registry.LocalMachine.OpenSubKey(installationKey).GetValue("Version").ToString();
            const string versionURI = @"https://raw.githubusercontent.com/sapph42/KoLMafia-Launcher/main/version.txt";
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
            string sourceURI = $"https://github.com/sapph42/KoLMafia-Launcher/raw/main/{targetInstallerName}";
            string destination = Environment.GetEnvironmentVariable("Temp") + @"\" + targetInstallerName;
            try {
                if (GetWebFile(sourceURI, destination, DownloadPriority.Foreground, null, null)) {
                    ProcessStartInfo processStartInfo = new() {
                        FileName = destination,
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
            string current = "";
            string latest = "";
            string jarURI = "";
            string fingerprintURI;
            string localFingerprint = "";
            string canonicalFingerprint = null;
            string KoLBaseLocation = @"https://builds.kolmafia.us/job/Kolmafia/lastSuccessfulBuild/artifact/dist/";
            string[] currentList;
            Hashtable command;
            HashAlgorithm cryptoService = MD5.Create();
            FileInfo currentFile;

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
                current = currentList[^1];
                localFingerprint = GetLocalHash(current, cryptoService);
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
                current = currentList[0];
                localFingerprint = GetLocalHash(current, cryptoService);
            }
            currentFile = new(current);
            try {
                HtmlWeb web = new();
                HtmlAgilityPack.HtmlDocument htmlDoc = web.Load(KoLBaseLocation);
                HtmlNode body = htmlDoc.DocumentNode.SelectSingleNode("//body");
                foreach (HtmlNode nNode in body.Descendants("a")) {
                    if (nNode.NodeType == HtmlNodeType.Element && nNode.InnerHtml.EndsWith(".jar")) {
                        latest = nNode.InnerHtml;
                    }
                }

                jarURI = KoLBaseLocation + latest;
                fingerprintURI = jarURI + @"/*fingerprint*/";

                web = new();
                htmlDoc = web.Load(fingerprintURI);
                body = htmlDoc.DocumentNode.SelectSingleNode("//body");
                foreach (HtmlNode nNode in body.Descendants("li")) {
                    if (nNode.NodeType == HtmlNodeType.Element && Regex.IsMatch(nNode.InnerHtml, @"[a-z0-9]{32}", RegexOptions.IgnoreCase)) {
                        canonicalFingerprint = nNode.InnerHtml;
                    }
                }

                if (!exists || (currentFile.Name != latest) || (localFingerprint != canonicalFingerprint)) {
                    int attempts = 1;
                    FileInfo destination = new(preferences.InstallLocation + @"\" + latest);
                    bool downloadSuccess = GetWebFile(jarURI, destination.FullName, DownloadPriority.Foreground, canonicalFingerprint, cryptoService);
                    while (!downloadSuccess && attempts < preferences.MaxAttempts) {
                        if (destination.Exists) { destination.Delete(); }
                        downloadSuccess = GetWebFile(jarURI, destination.FullName, DownloadPriority.Foreground, canonicalFingerprint, cryptoService);
                        attempts++;
                    }
                    if (exists && downloadSuccess) {
                        currentFile.Delete();
                        latest = destination.FullName;
                    }
                } else {
                    latest = currentFile.FullName;
                }
            } catch (HttpRequestException e) {
                Console.WriteLine("\nException Caught!");
                Console.WriteLine(e.Message);
                if (!silent) {
                    MessageBox.Show(@"Received a error when attempting to retreive and/or save the latest version. Your previous version has been kept, and will now be run.");
                    latest = currentFile.FullName;
                }
            } finally {
                if (!noLaunch) {
                    FileInfo thisFile = new(latest);
                    ProcessStartInfo processStartInfo = new() {
                        FileName = javaPath,
                        Arguments = javaParams.Replace(@"%1", thisFile.FullName).Replace(@" %*",""),
                        WorkingDirectory = thisFile.Directory.FullName
                    };
                    System.Diagnostics.Process.Start(processStartInfo);
                }
            }
        }
    }
}
