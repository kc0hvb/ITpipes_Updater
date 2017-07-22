using ITpipes_Updater.Util;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Windows;

namespace ITpipes_Updater
{
    public class InstallationLogic : IDisposable
    {
        public event EventHandler<ProgressReportEventArgs> ProgressUpdateAvailable;

        private readonly bool
            OVERWRITE_FILES = true,
            DO_NOT_OVERWRITE_FILES = false;

        public const string DEFAULT_INSTALL_DIRECTORY = @"C:\Program Files\InspectIT";

        private const string msiExecFile = @"C:\Windows\System32\msiexec.exe";
        private bool disposed = false;
        private string
            pathToUninstaller,
            updaterLaunchDirectory,
            systemFolder = Environment.GetFolderPath(Environment.SpecialFolder.SystemX86);

        private DateTime _startTime = DateTime.Now;
        private StreamWriter _logWriter = null;

        private DirectoryInfo mainProgramDirectoryInfo;

        private OleDbConnection liveItPipesSetupConn,
                                oldDbSetupConn;
        
        private RegistrationHelper _regHelper = new RegistrationHelper();
        private InstallationRequest _config;
        public bool bErrorExists = false;

        public static string GetPathToUninstaller()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), @"ITpipes\uninstall.exe");
        }

        public InstallationLogic(InstallationRequest config)
        {
            try {
                this.pathToUninstaller = GetPathToUninstaller();
                updaterLaunchDirectory = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
                mainProgramDirectoryInfo = new DirectoryInfo(config.InstallationPath);
                _config = config;
                string logFilePath = Path.Combine(updaterLaunchDirectory, $@"Logs\{_startTime.Ticks}.log");
                if (!Directory.Exists(Path.GetDirectoryName(logFilePath))) Directory.CreateDirectory(Path.GetDirectoryName(logFilePath));
                File.Create(logFilePath).Close();
                _logWriter = new StreamWriter(logFilePath);
                _logWriter.AutoFlush = true;
            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
                bErrorExists = true;
            }
            }

        private void updateStatus(string message)
        {
            try {
                log(message);
                ProgressUpdateAvailable?.Invoke(this, new ProgressReportEventArgs(message));
            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
                bErrorExists = true;
            }
            }

        private void log(string message)
        {
            _logWriter.WriteLine(message);
        }

        private void createNewSettingsRow(string oleConnString)
        {
            try
            {
                using (OleDbConnection curConn = new OleDbConnection(oleConnString))
                using (OleDbCommand curCommand = curConn.CreateCommand())
                {
                    curConn.Open();

                    curCommand.CommandText = "SELECT TOP 1 S_ID FROM S";

                    int? settingsRowPK = (int?)curCommand.ExecuteScalar();

                    if (settingsRowPK != null)
                    {
                        return;
                    }

                    curCommand.CommandText =
                        "INSERT INTO `S` (" +
                        "S_ID, Video_Quality, Counter_Calibration, Speech_Speed, P_ID, Files_DefaultProjectPath, Misc_SplashScreen, Video_WriteData, Video_Format, Counter_Type, Misc_Backup_Files, Misc_DBUseADODC) " +
                        "VALUES (1, 'CD', 1, 'Normal', 0, 'C:\\IT Projects', 0, 0, 'WMV', 'USDigital', 1, 1)";

                    try
                    {
                        curCommand.ExecuteNonQuery();
                    }
                    catch
                    {
                        curCommand.CommandText =
                        "INSERT INTO `S` (" +
                        "Video_Quality, Counter_Calibration, Speech_Speed, P_ID, Files_DefaultProjectPath, Misc_SplashScreen, Video_WriteData, Video_Format, Counter_Type, Misc_Backup_Files, Misc_DBUseADODC) " +
                        "VALUES ('CD', 1, 'Normal', 0, 'C:\\IT Projects', 0, 0, 'WMV', 'USDigital', 1, 1)";
                        try
                        {
                            curCommand.ExecuteNonQuery();
                        }
                        catch
                        {
                            throw new InstallationFailedException("Could not create new settings row");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
                bErrorExists = true;
            }
            }

        private void installRedists()
        {
            string redistDirectoryPath = updaterLaunchDirectory + @"\Redist";

            if (!Directory.Exists(redistDirectoryPath))
            {
                throw new InstallationFailedException("Cannot install ITpipes: Directory \"Redist\" does not exist in updater folder.");
            }

            foreach (string curRedistFile in Directory.GetFiles(redistDirectoryPath, "*", SearchOption.TopDirectoryOnly))
            {
                installMsiSilently(curRedistFile);
            }
        }

        private void installMsiSilently(string redistFile)
        {
            if (string.Equals(Path.GetExtension(redistFile), ".msi", StringComparison.InvariantCultureIgnoreCase))
            {
                log($"Installing Redist MSI '{redistFile}'");
                Process msiInstallProcess = new Process();
                msiInstallProcess.StartInfo.FileName = msiExecFile;
                msiInstallProcess.StartInfo.Arguments = string.Format("/i \"{0}\" /quiet /qn", redistFile);
                msiInstallProcess.StartInfo.RedirectStandardOutput = true;
                msiInstallProcess.StartInfo.CreateNoWindow = true;
                msiInstallProcess.StartInfo.UseShellExecute = false;
                msiInstallProcess.Start();
                msiInstallProcess.WaitForExit();
            }
        }

        private void createProgramShortcuts()
        {
            string itpipesExecutablePath = Path.Combine(_config.InstallationPath, "InspectIT.exe"); //setting this to our "dummy" executable so I can do things like checking if ITpipes is running.
            string templateEditorExecutablePath = Path.Combine(_config.InstallationPath, "TemplateEditor.exe");
            string advancedMergeExecutablePath = Path.Combine(_config.InstallationPath, "Merge.exe");
            string manageItExecutablePath = Path.Combine(_config.InstallationPath, "ManageIT.exe");
            string itpipesUninstallerExecutablePath = pathToUninstaller;
            string pathToInstallationManager = Assembly.GetCallingAssembly().Location;

            string startMenuITpipesDirectory = Path.Combine(UtilFunctions.GetAllUsersStartMenuFolder(), "ITpipes");

            if (Directory.Exists(startMenuITpipesDirectory) == false)
            {
                Directory.CreateDirectory(startMenuITpipesDirectory);
            }

            log("Creating start menu and desktop shortcuts");

            try
            {
                UtilFunctions.CreateShortcut(itpipesExecutablePath, UtilFunctions.GetAllUsersDesktopFolder(), "ITpipes");
                UtilFunctions.CreateShortcut(itpipesExecutablePath, startMenuITpipesDirectory, "ITpipes");
                UtilFunctions.CreateShortcut(templateEditorExecutablePath, startMenuITpipesDirectory, "Template Editor");
                UtilFunctions.CreateShortcut(manageItExecutablePath, startMenuITpipesDirectory, "ManageIT (SQL)");
                UtilFunctions.CreateShortcut(pathToInstallationManager, startMenuITpipesDirectory, "Installation Manager");
                //UtilFunctions.CreateShortcut(itpipesUninstallerExecutablePath, startMenuITpipesDirectory, "Uninstall ITpipes");
            }
            catch (Exception ex)
            {
                log($"Exception while creating shortcuts: {ex.Message}");
            }
        }

        public void RunAsInstaller()
        {
            updateStatus("Installing program dependencies...");
            installRedists();

            updateStatus("Installing default user profiles...");
            copyDefaultUsers();

            updateStatus("Adding information to system registry...");
            addRegKeyForProgramInstallDirectory();

            if (_config.BackupToRestore != null)
            {
                updateStatus("Restoring configuration and user backup...");
                ITP_Backup.RestoreBackup(_config.InstallationPath, _config.BackupToRestore);
            }

            RunAsUpdater();

            updateStatus("Setting up default ITpipes settings...");
            createNewSettingsRow(string.Format("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = {0}", (_config.InstallationPath + @"\setup.mdb")));
            addDisableVideoFeedCloningMarker();

            updateStatus("Creating program shortcuts...");
            createProgramShortcuts();
        }

        private void copyDefaultUsers()
        {
            _updateDirectory("Users", "Users", (fileToCopy) => !File.Exists(fileToCopy), (fileToRegister) => false);
        }

        private void addRegKeyForProgramInstallDirectory()
        {
            //while we do handle both 32 and 64 bit registry values for checking if ITpipes is installed and for uninstalling ITpipes, when writing the new
            //Installer registry info, we should NOT have special handling for 32/64 bit systems. Windows will handle any 32/64 bit shenanigans for us.
            //The checks and uninstaller removal of both 32/64 bit is necessary because of the previous installer.

            using (RegistryKey installLocationRegKey = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\ITpipes", RegistryKeyPermissionCheck.ReadWriteSubTree))
            using (RegistryKey uninstallerRegKey = Registry.LocalMachine.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\ITpipes"))
            {
                installLocationRegKey.SetValue("Install_Dir", @"C:\Program Files\InspectIT", RegistryValueKind.String);

                uninstallerRegKey.SetValue("DisplayName", "ITpipes");
                uninstallerRegKey.SetValue("DisplayIcon", pathToUninstaller + ",0");
                uninstallerRegKey.SetValue("DisplayVersion", _config.VersionNumber.Replace("v", "")); //don't want the 'v' in front of version number in uninstall info
                uninstallerRegKey.SetValue("EstimatedSize", UtilFunctions.GetDirectorySizeInBytes(updaterLaunchDirectory) / 1024, RegistryValueKind.DWord);
                uninstallerRegKey.SetValue("Publisher", "Infrastructure Technologies");
                uninstallerRegKey.SetValue("UninstallString", pathToUninstaller);
                uninstallerRegKey.SetValue("NoModify", 1);
                uninstallerRegKey.SetValue("NoRepair", 1);
            }
        }

        public void RunAsUpdater()
        {
            //Only create a backup if backup wasn't restored during installation
            if (_config.BackupToRestore == null)
            {
                updateStatus("Backing up sensitive files...");
                backupSettingsAndConfig();
            }
            else
            {
                //still need to copy the setup.mdb to setup.bak to ensure settings are properly transferred to new setup.mdb file:
                string setupMdb = Path.Combine(_config.InstallationPath, "setup.mdb");

                if (File.Exists(setupMdb))
                {
                    UtilFunctions.copyFile(setupMdb, setupMdb + ".bak", true);
                }
            }

            updateStatus("Compositing file write virtualizations into ITpipes data...");
            try
            {
                VirtualStoreHelper.IntegrateDirectory(_config.InstallationPath);
            }
            catch (Exception ex)
            {
                log($"Exception in VirtualStore compositing logic: {ex.Message}");
            }

            updateStatus("Updating libraries...");
            copyAndRegSystemFiles();
            registerMontivisionITPipesLogo();
            reRegisterProblematicActiveXControls();

            updateStatus("Updating overlay control drivers...");
            copyAndRegDriversFolder();

            updateStatus("Updating Registry Keys...");
            registerRegKeys();

            updateStatus("Updating default databases...");
            replaceBlankDBFiles();

            updateStatus("Updating ITpipes executables...");
            copyFILEDirectory();

            updateStatus("Updating Template files...");
            copyTemplates();

            updateStatus("Transferring ITpipes settings to updated settings format...");
            updateSetupMDB();

            updateStatus("Installing latest version of ViewIT...");
            replaceViewIT();

            updateStatus("Adding Report temporary directory");
            addReportTempDirectory();

            updateStatus("Placing FFmpeg in default location.");
            processFfmpegInstall();

            updateStatus("Installing Mpeg4 if enabled.");
            processInstallationMpeg4();



            updateStatus("Scanning and fixing corrupt user files...");
            try
            {
                UtilFunctions.scanAndFixCorruptUserFiles(_config.InstallationPath);
            }
            catch (Exception ex)
            {
                //Not a failure state
                log($"Exception scanning and fixing corrupt user files: {ex.Message}");
            }

            updateStatus("Updating ITpipes directory permissions...");
            setAccessPermissionsForProgramFolders();
        }

        private void addDisableVideoFeedCloningMarker()
        {
            string pathToCommandFile = Path.Combine(_config.InstallationPath, "turnoverlay.off");

            File.Create(pathToCommandFile).Close();
        }

        private void replaceViewIT()
        {
            try
            {
                _updateDirectory("ViewIT", "ViewIT", null, (x) => false);
            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
                bErrorExists = true;
            }
            }

        private void copyTemplates()
        {
            try
            {
                _updateDirectory("Templates", @"Templates\Stock Templates");
                if (!_config.OverwriteTemplates && !_config.InstallerMode)
                {
                    return;
                }

                Func<string, bool> shouldCopyFile = (targetPath) =>
                {
                    if (_config.OverwriteTemplates || !File.Exists(targetPath))
                    {
                        return true;
                    }
                    return false;
                };

                _updateDirectory("Templates", "Templates", shouldCopyFile);
            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
                bErrorExists = true;
            }
            }

        private void backupSettingsAndConfig()
        {
            //using the same backup code as the Uninstaller/Config:
            try
            {
                UtilFunctions.CreateBackup();
            }
            catch (Exception ex)
            {
                log("Failed to create backup for ITpipes data. Reason: " + ex.Message);
            }
        }

        private void registerMontivisionITPipesLogo()
        {
            Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Montivision\DemoSource\DefaultLogo");
            RegistryKey demoLogoKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Montivision\DemoSource\DefaultLogo", true);
            demoLogoKey.SetValue("", @"C:\Program Files\InspectIT\Libraries\itpipes_logo.bmp");
        }

        private void registerRegKeys()
        {
            try
            {
                string regKeysFolder = updaterLaunchDirectory + @"\RegKeys";

                if (Directory.Exists(regKeysFolder) == false)
                {
                    log("Could not find the updater RegKeys folder--no registry keys could be added.");
                    return;
                }

                string[] availableRegKeys = Directory.GetFiles(regKeysFolder, "*.reg", SearchOption.AllDirectories);

                foreach (string curRegKey in availableRegKeys.AsEnumerable())
                {
                    Process regeditProcess = Process.Start("regedit.exe", "/s \"" + curRegKey + "\"");
                    regeditProcess.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
                bErrorExists = true;
            }
            }

        private void _updateDirectory(string folderName, string targetSubfolder, Func<string, bool> shouldCopyFileCallback = null, Func<string, bool> shouldRegisterFileCallback = null)
        {
            if (shouldCopyFileCallback == null)
            {
                shouldCopyFileCallback = (x) => true;
            }
            if (shouldRegisterFileCallback == null)
            {
                shouldRegisterFileCallback = (x) => true;
            }

            string sourceDir = UtilFunctions.FixSlashes($@"{updaterLaunchDirectory}\{folderName}");
            string targetDir = UtilFunctions.FixSlashes($@"{_config.InstallationPath}\{targetSubfolder}");
            if (!Directory.Exists(sourceDir))
            {
                throw new InstallationFailedException($"Could not locate updater directory subfolder '{folderName}'");
            }

            foreach (string curFile in Directory.GetFiles(sourceDir, "*", SearchOption.AllDirectories))
            {
                string targetPath = curFile.Replace(sourceDir, targetDir);

                if (shouldCopyFileCallback(targetPath))
                {
                    if (!Directory.Exists(Path.GetDirectoryName(targetPath)))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(targetPath));
                    }
                    UtilFunctions.copyFile(curFile, targetPath, true);
                }
            }

            //Register after all files are copied--if there are dependent libraries/assemblies registration could fail if file copying is incomplete.
            foreach (string curFile in Directory.GetFiles(targetDir, "*", SearchOption.AllDirectories))
            {
                if (shouldRegisterFileCallback(curFile) && _regHelper.IsFileRegisterableType(curFile))
                {
                    _regHelper.RegisterFile(curFile);
                }
            }
        }

        private void copyAndRegSystemFiles()
        {
            try {
                string source = "SYSTEM32";
                string targetSubDir = "Libraries";
                Func<string, bool> shouldCopyCallback = (filePath) =>
                {
                    string fileName = Path.GetFileNameWithoutExtension(filePath);

                //Since demosource.ax is a virtual video input, Chrome frequently locks the file to access the video input for Hangouts--only copy and register if the file doesn't already exist or the updater will fail if Chrome is open.
                if (fileName.Equals("Demosource", StringComparison.InvariantCultureIgnoreCase) && File.Exists(filePath))
                    {
                        _regHelper.RegisterAsCom(filePath);
                        return false;
                    }

                    return true;
                };

                _updateDirectory(source, targetSubDir, shouldCopyCallback);
            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
                bErrorExists = true;
            }
            }

        #region Copying the File Directory from UpdateIT
        private void copyFILEDirectory()
        {
            try
            {
                string source = "FILE";
                string targetSubDir = @"";
                Func<string, bool> shouldCopyCallback = (targetFilePath) =>
                {
                    string fileNameOnly = Path.GetFileNameWithoutExtension(targetFilePath);
                    if (Path.GetExtension(targetFilePath).ToLower() == ".ini" || fileNameOnly == "address" || fileNameOnly == "setup")
                    {
                        if (File.Exists(targetFilePath))
                        {
                            return false;
                        }
                    }
                    else if ((fileNameOnly == "qsb_counter" || fileNameOnly == "usdqsb" || fileNameOnly == "usdigital") && !_config.InstallQSB)
                    {
                        return false;
                    }
                    else if (fileNameOnly == "uninstall")
                    {
                        UtilFunctions.copyFile(Path.Combine(updaterLaunchDirectory, $@"FILE\uninstall.exe"), pathToUninstaller, true);
                        return false;
                    }
                    return true;
                };

                _updateDirectory(source, targetSubDir, shouldCopyCallback);
            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
                bErrorExists = true;
            }
        }
        #endregion

        #region Copying and Registering Driver Files
        private void copyAndRegDriversFolder()
        {
            try
            {
                string source = "Drivers";
                string targetSubdir = "Drivers";
                _updateDirectory(source, targetSubdir);
            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
                bErrorExists = true;
            }
        }
        #endregion

        #region Replacing the BlankDB Folder with UpdateIT's folder
        private void replaceBlankDBFiles()
        {
            try {
                string source = "BlankDB";
                string targetSubdir = "BlankDB";
                _updateDirectory(source, targetSubdir);
            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
                bErrorExists = true;
            }
        }
        #endregion

        #region Getting the connection String for the Setup.mdb
        private string getConnString(string pathToFile)
        {
            return "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + pathToFile;
        }
        #endregion


        #region Updating the Setup.MDB
        private bool updateSetupMDB()
        {
            //Implementation note: must process blankDB update first, since the template setup.mdb file is contained within the blankDB folder

            bool succeeded = true;
            bool oldSetupMDBPresent = true;
            string locationOfSetupFileBackup = @"C:\Program Files\InspectIT\setup.mdb.bak";
            string itPipesSetupPath = _config.InstallationPath + @"\setup.mdb";
            liveItPipesSetupConn = new OleDbConnection(getConnString(itPipesSetupPath));


            DataSet oldSetupDS = new DataSet(),
                    liveSetupDS = new DataSet();

            if (File.Exists(locationOfSetupFileBackup) == false)
            {
                oldSetupMDBPresent = false;
            }

            try
            {
                UtilFunctions.copyFile(_config.InstallationPath + @"\BlankDB\setup.mdb", itPipesSetupPath);
            }
            catch (Exception ex)
            {
                string error = $"Could not replace old setup.mdb file with newest setup database. Verify that setup.mdb is not open currently: {ex.Message}";
                log(error);
                throw new InstallationFailedException(error);
            }

            if (oldSetupMDBPresent)
            {
                oldDbSetupConn = new OleDbConnection(getConnString(locationOfSetupFileBackup));

                OleDbDataAdapter liveSetupDA = new OleDbDataAdapter("", liveItPipesSetupConn),
                                 oldDbSetupDA = new OleDbDataAdapter("", oldDbSetupConn);

                OleDbCommand liveSetupCommand = new OleDbCommand("", liveItPipesSetupConn);

                try
                {
                    fillSetupDataSet(liveSetupDA, liveItPipesSetupConn, liveSetupDS);
                    fillSetupDataSet(oldDbSetupDA, oldDbSetupConn, oldSetupDS);
                }
                catch (Exception ex)
                {
                    string error = $"Failed to fill DataSets from backed-up setup.mdb and live setup.mdb: {ex.Message}";
                    log(error);
                    throw new InstallationFailedException(error);
                }

                processSTableUpdate(oldSetupDS, liveSetupDS, liveSetupCommand);
                persistMpegRecordingSettings(oldSetupDS, liveSetupDS, liveSetupCommand);
                addSVideoProfileIfAvailable(oldSetupDS, liveSetupDS, liveSetupCommand);
                processRVideoProfilesUpdate(oldSetupDS, liveSetupDS, liveSetupCommand);
                processProjectListMigration(oldSetupDS, liveSetupDS, getConnString(itPipesSetupPath));
            }

            populateIbakPathsInSetupDb(liveItPipesSetupConn);

            return succeeded;
        }
        #endregion
        private void fillSetupDataSet(OleDbDataAdapter curAdapter, OleDbConnection curConnection, DataSet curDS)
        {
            if (curConnection.State != ConnectionState.Open) { curConnection.Open(); }

            curAdapter.SelectCommand.CommandText = "SELECT TOP 1 * FROM S";
            curAdapter.Fill(curDS, "S");

            curAdapter.SelectCommand.CommandText = "SELECT * FROM P";
            curAdapter.Fill(curDS, "P");

            curAdapter.SelectCommand.CommandText = "SELECT * FROM R_Video_Profiles";
            curAdapter.Fill(curDS, "R_Video_Profiles");

            curAdapter.SelectCommand.CommandText = "SELECT * FROM S_Category";
            curAdapter.Fill(curDS, "S_Category");

            curAdapter.SelectCommand.CommandText = "SELECT * FROM S_CategoryGroup";
            curAdapter.Fill(curDS, "S_CategoryGroup");

            curAdapter.SelectCommand.CommandText = "SELECT * FROM S_CategoryOption";
            curAdapter.Fill(curDS, "S_CategoryOption");

            curAdapter.SelectCommand.CommandText = "SELECT * FROM S_Video";
            curAdapter.Fill(curDS, "S_Video");

            curAdapter.SelectCommand.CommandText = "SELECT [Version_ID], [Version], [Date_Modified], [Notes] FROM Version";
            curAdapter.Fill(curDS, "Version");

            curConnection.Close();
        }

        #region Placing ffmpeg.exe in the C:\FFmpeg\

        private void processFfmpegInstall()
        {
            try
            {
                System.IO.Directory.CreateDirectory("C:\\FFmpeg\\");
                bool os_platform = System.Environment.Is64BitOperatingSystem;
                if (os_platform == false)
                {
                    string source = updaterLaunchDirectory + "\\FFmpeg\\ffmpegx86.exe";
                    string targetDir = "C:\\FFMpeg\\ffmpeg.exe";
                    System.IO.File.Copy(source, targetDir, true);
                }
                else
                {
                    string source = updaterLaunchDirectory + "\\FFmpeg\\ffmpegx64.exe";
                    string targetDir = "C:\\FFMpeg\\ffmpeg.exe";
                    System.IO.File.Copy(source, targetDir, true);
                }
            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
                bErrorExists = true;
            }
        }

        #endregion

        #region Installation of MP4Box
        private void processInstallationMpeg4()
        {
            try
            {
                bool isInstallMp4Enabled = _config.InstallMP4;
                if (isInstallMp4Enabled == true)
                {
                    bool os_platform64 = false;
                    os_platform64 = System.Environment.Is64BitOperatingSystem;
                    string sStartPath = updaterLaunchDirectory + "\\MP4Box\\";
                    Process vLaunchMp4BoxInstaller = new Process();
                    bool mp4BoxIsInstalled = IsSoftwareInstalled("GPAC");
                    if (mp4BoxIsInstalled == false)
                    {
                        if (os_platform64 == true)
                        {
                            string[] sFile64bit = Directory.GetFiles(sStartPath, "*x64.exe");
                            vLaunchMp4BoxInstaller = Process.Start(sFile64bit[0]);
                            vLaunchMp4BoxInstaller.WaitForExit();
                        }
                        else
                        {
                            string[] sFile32bit = Directory.GetFiles(sStartPath, "*x32.exe");
                            vLaunchMp4BoxInstaller = Process.Start((updaterLaunchDirectory) + "\\MP4Box\\" + sFile32bit);
                            vLaunchMp4BoxInstaller.WaitForExit();
                        }
                    }
                    string sMp4boxLocation = Mp4BoxLocation();
                    string sMp4BoxSource = sMp4boxLocation + "mp4box.exe";
                    string sMp4BoxDest = _config.InstallationPath + "\\mp4box.exe";
                    System.IO.File.Copy(sMp4BoxSource, sMp4BoxDest, true);
                    string sCodecLocation = _config.InstallationPath + "\\LeadToolsUnlock.exe";
                    var vCodecInstall = Process.Start(sCodecLocation);
                    vCodecInstall.WaitForExit();
                }

            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
                bErrorExists = true;
            }
        }

        #endregion
        #region Finding if MP4Box is installed.
        public bool IsSoftwareInstalled(string sSoftwareName)
        {
            try
            {
                string displayName;
                string registryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
                RegistryKey key = Registry.LocalMachine.OpenSubKey(registryKey);
                if (key != null)
                {
                    foreach (RegistryKey subkey in key.GetSubKeyNames().Select(keyName => key.OpenSubKey(keyName)))
                    {
                        displayName = subkey.GetValue("DisplayName") as string;
                        if (displayName != null && displayName.Contains(sSoftwareName))
                        {
                            return true;
                        }
                    }
                    key.Close();
                }

                registryKey = @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall";
                key = Registry.LocalMachine.OpenSubKey(registryKey);
                if (key != null)
                {
                    foreach (RegistryKey subkey in key.GetSubKeyNames().Select(keyName => key.OpenSubKey(keyName)))
                    {
                        displayName = subkey.GetValue("DisplayName") as string;
                        if (displayName != null && displayName.Contains(sSoftwareName))
                        {
                            return true;
                        }
                    }
                    key.Close();
                }
                return false;
            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
                bErrorExists = true;
                return false;
            }
        }

        #endregion

        #region Getting the MP4BoxLocation

        public string Mp4BoxLocation()
        {
            try
            {
                string displayName;
                string registryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
                RegistryKey key = Registry.LocalMachine.OpenSubKey(registryKey);
                if (key != null)
                {
                    foreach (RegistryKey subkey in key.GetSubKeyNames().Select(keyName => key.OpenSubKey(keyName)))
                    {
                        displayName = subkey.GetValue("UninstallString") as string;
                        if (displayName != null && displayName.Contains("GPAC"))
                        {
                            string sLocation = displayName.Replace("uninstall.exe", "");
                            return sLocation;
                        }
                    }
                    key.Close();
                }

                registryKey = @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall";
                key = Registry.LocalMachine.OpenSubKey(registryKey);
                if (key != null)
                {
                    foreach (RegistryKey subkey in key.GetSubKeyNames().Select(keyName => key.OpenSubKey(keyName)))
                    {
                        displayName = subkey.GetValue("UninstallString") as string;
                        if (displayName != null && displayName.Contains("GPAC"))
                        {
                            string sLocation = displayName.Replace ("uninstall.exe", "");
                            return sLocation;
                        }
                    }
                    key.Close();
                }
                return null;
            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
                bErrorExists = true;
                return null;
            }
}
        #endregion
        private void processSTableUpdate(DataSet oldDS, DataSet liveDS, OleDbCommand liveSetupCommand)
        {
            try
            {
                if (liveSetupCommand.Connection.State != ConnectionState.Open) { liveSetupCommand.Connection.Open(); }


                if (oldDS.Tables["S"].Rows.Count > 0)
                {
                    liveSetupCommand.Parameters.Clear();

                    List<string> columnNames = new List<string>();
                    List<char> parameters = new List<char>();

                    foreach (DataColumn curColumn in oldDS.Tables["S"].Columns)
                    {
                        if (oldDS.Tables["S"].Rows[0][curColumn.ColumnName].Equals(DBNull.Value) == false && curColumn.ColumnName != "S_ID")
                        {
                            if (!liveDS.Tables["S"].Columns.Contains(curColumn.ColumnName))
                            {
                                continue;
                            }

                            parameters.Add('?');
                            columnNames.Add(curColumn.ColumnName);
                            liveSetupCommand.Parameters.AddWithValue(curColumn.ColumnName, oldDS.Tables["S"].Rows[0][curColumn.ColumnName]);
                        }
                    }
                    string query = $"INSERT INTO S ({string.Join(", ", columnNames)}) VALUES ({string.Join(", ", parameters)});";
                    liveSetupCommand.CommandText = query;
                    liveSetupCommand.ExecuteNonQuery();

                    liveSetupCommand.Parameters.Clear();
                }

                liveSetupCommand.Connection.Close();
            }
            catch (Exception ex)
            {
                string error = $"Failed to move settings (S table) from old setup.mdb into new setup.mdb: {ex.Message}";
                log(error);
                throw new InstallationFailedException(error);
            }
        }

        private void writeAllRowsToTargetDatabase(DataTable sourceTable, DataTable targetTable, string tableName, string targetDbConnString)
        {
            using (OleDbConnection curConn = new OleDbConnection(targetDbConnString))
            using (OleDbCommand curCommand = curConn.CreateCommand())
            {
                curConn.Open();

                using (OleDbTransaction curTransaction = curConn.BeginTransaction())
                {
                    curCommand.Transaction = curTransaction;

                    foreach (DataRow curRow in sourceTable.Rows)
                    {
                        List<string> fieldNames = new List<string>(), 
                            fieldValues = new List<string>();

                        foreach (DataColumn curColumn in sourceTable.Columns)
                        {
                            if (!targetTable.Columns.Contains(curColumn.ColumnName))
                            {
                                continue;
                            }

                            object sourceRowValue = curRow[curColumn.ColumnName];

                            if (sourceRowValue == null || sourceRowValue == DBNull.Value)
                            {
                                continue;
                            }

                            fieldNames.Add(curColumn.ColumnName);
                            fieldValues.Add(getQuerySafe(curRow[curColumn.ColumnName]));
                        }

                        curCommand.CommandText = $"INSERT INTO {tableName} ({string.Join(", ", fieldNames)}) VALUES ({string.Join(", ", fieldValues)});";
                        curCommand.ExecuteNonQuery();
                    }

                    curTransaction.Commit();
                }
            }
        }

        private string getQuerySafe(object inValue)
        {
            if (inValue == null || inValue == DBNull.Value)
            {
                return "NULL";
            }

            Type objType = inValue.GetType();
            if (objType == typeof(string) || objType == typeof(DateTime))
            {
                return $"'{inValue.ToString().Replace("'", "''")}'";
            }

            if (objType == typeof(bool))
            {
                return (bool)inValue ? "1" : "0";
            }

            return inValue.ToString();
        }

        private void processProjectListMigration(DataSet oldDS, DataSet liveDS, string liveSetupConnString)
        {
            string tblName = "P";
            DataTable sourceTbl = oldDS.Tables[tblName];
            DataTable targetTbl = liveDS.Tables[tblName];
            writeAllRowsToTargetDatabase(sourceTbl, targetTbl, tblName, liveSetupConnString);
        }

        private void processProjectListMigration(DataSet oldDS, DataSet liveDS, OleDbCommand liveSetupCommand)
        {
            try
            {
                if (liveSetupCommand.Connection.State != ConnectionState.Open)
                {
                    liveSetupCommand.Connection.Open();
                }

                foreach (DataRow oldProfileRow in oldDS.Tables["P"].Rows)
                {

                    liveSetupCommand.CommandText = "INSERT INTO `P` ([ProjectName], " +
                                                                  "[ProjectDataPath], " +
                                                                  "[ProjectDriveSerial], " +
                                                                  "[ProjectDBName], " +
                                                                  "[ProjectStreet], " +
                                                                  "[ProjectState], " +
                                                                  "[ProjectZip], " +
                                                                  "[ProjectPhone], " +
                                                                  "[ProjectCity], " +
                                                                  "[CreateProjectFolder], " +
                                                                  "[ProjectHidden], " +
                                                                  "[Identity], " +
                                                                  "[ProjectType], " +
                                                                  "[SQLUser], " +
                                                                  "[SQLPassword], " +
                                                                  "[SQLIP], " +
                                                                  "[Template_ID], " +
                                                                  "[Template_Modified_Date]) " +
                                                                  "Values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);";

                    liveSetupCommand.Parameters.Clear();

                    liveSetupCommand.Parameters.AddWithValue("ProjectName", oldProfileRow["ProjectName"]);
                    liveSetupCommand.Parameters.AddWithValue("ProjectDataPath", oldProfileRow["ProjectDataPath"]);
                    liveSetupCommand.Parameters.AddWithValue("ProjectDriveSerial", oldProfileRow["ProjectDriveSerial"]);
                    liveSetupCommand.Parameters.AddWithValue("ProjectDBName", oldProfileRow["ProjectDBName"]);
                    liveSetupCommand.Parameters.AddWithValue("ProjectStreet", oldProfileRow["ProjectStreet"]);
                    liveSetupCommand.Parameters.AddWithValue("ProjectState", oldProfileRow["ProjectState"]);
                    liveSetupCommand.Parameters.AddWithValue("ProjectZip", oldProfileRow["ProjectZip"]);
                    liveSetupCommand.Parameters.AddWithValue("ProjectPhone", oldProfileRow["ProjectPhone"]);
                    liveSetupCommand.Parameters.AddWithValue("ProjectCity", oldProfileRow["ProjectCity"]);
                    liveSetupCommand.Parameters.AddWithValue("CreateProjectFolder", oldProfileRow["CreateProjectFolder"]);
                    liveSetupCommand.Parameters.AddWithValue("ProjectHidden", oldProfileRow["ProjectHidden"]);
                    liveSetupCommand.Parameters.AddWithValue("Identity", oldProfileRow["Identity"]);
                    liveSetupCommand.Parameters.AddWithValue("ProjectType", oldProfileRow["ProjectType"]);
                    liveSetupCommand.Parameters.AddWithValue("SQLUser", oldProfileRow["SQLUser"]);
                    liveSetupCommand.Parameters.AddWithValue("SQLPassword", oldProfileRow["SQLPassword"]);
                    liveSetupCommand.Parameters.AddWithValue("SQLIP", oldProfileRow["SQLIP"]);
                    liveSetupCommand.Parameters.AddWithValue("Template_ID", oldProfileRow["Template_ID"]);
                    liveSetupCommand.Parameters.AddWithValue("Template_Modified_Date", oldProfileRow["Template_Modified_Date"]);

                    liveSetupCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                string error = $"Failed to merge project list from old setup.mdb into new setup.mdb: {ex.Message}";
                log(error);
                throw new InstallationFailedException(error);
            }
            finally
            {
                if (liveSetupCommand != null && liveSetupCommand.Connection != null)
                {
                    liveSetupCommand.Connection.Close();
                }
            }
        }

        private void persistMpegRecordingSettings(DataSet oldDS, DataSet liveDS, OleDbCommand liveSetupCommand)
        {
            //IMPORTANT! This is bad, and needs to be replaced. This is a workaround

            try
            {
                bool mpeg4Found = false,
                     mpeg2Found = false;
                foreach (DataRow curRow in oldDS.Tables["S_CategoryOption"].Rows)
                {
                    if (curRow["CategoryOptionName"].ToString() == "MPEG4")
                    {
                        mpeg4Found = true;
                    }
                    else if (curRow["CategoryOptionName"].ToString() == "MPEG2")
                    {
                        mpeg2Found = true;
                    }
                }

                if (mpeg2Found == true)
                {
                    if (liveSetupCommand.Connection.State != ConnectionState.Open)
                    {
                        liveSetupCommand.Connection.Open();
                    }

                    //with the current setup.mdb, the categorygroupid field for video formats is hard-bound to 94.
                    liveSetupCommand.CommandText = "INSERT INTO S_CategoryOption (CategoryOptionName, CategoryGroupID) VALUES (\"MPEG2\", 94)";
                    liveSetupCommand.ExecuteNonQuery();
                }

                if (mpeg4Found == true)
                {
                    if (liveSetupCommand.Connection.State != ConnectionState.Open)
                    {
                        liveSetupCommand.Connection.Open();
                    }

                    //with the current setup.mdb, the categorygroupid field for video formats is hard-bound to 94.
                    liveSetupCommand.CommandText = "INSERT INTO S_CategoryOption (CategoryOptionName, CategoryGroupID) VALUES (\"MPEG4\", 94)";
                    liveSetupCommand.ExecuteNonQuery();
                }

                if (liveSetupCommand.Connection.State == ConnectionState.Open)
                {
                    liveSetupCommand.Connection.Close();
                }
            }
            catch (Exception ex)
            {
                string error = $"Failed to process MPEG 2/4 dropdown option settings: {ex.Message}";
                log(error);
                throw new InstallationFailedException(error);
            }
        }

        private void addSVideoProfileIfAvailable(DataSet oldDS, DataSet liveDS, OleDbCommand liveSetupCommand)
        {
            try
            {
                if (oldDS.Tables["S_Video"].Rows.Count == 0)
                {
                    return;
                }

                string tblName = "S_Video";
                DataTable sourceTbl = oldDS.Tables[tblName],
                    targetTbl = liveDS.Tables[tblName];
                string connString = liveItPipesSetupConn.ConnectionString;

                writeAllRowsToTargetDatabase(sourceTbl, targetTbl, tblName, connString);
                return;
            }
            catch (Exception ex)
            {
                string error = $"Failed to move video capture profile (S_Video Table) to new setup.mdb: {ex.Message}";
                log(error);
                throw new InstallationFailedException(error);
            }
        }

        private string replaceEmptyWithNull(string input)
        {
            if (input.Trim() == string.Empty)
            {
                return "NULL";
            }

            return "\'" + input + "\'";
        }

        private void processRVideoProfilesUpdate(DataSet oldDS, DataSet liveDS, OleDbCommand liveSetupCommand)
        {
            try
            {
                if (liveSetupCommand.Connection.State != ConnectionState.Open)
                {
                    liveSetupCommand.Connection.Open();
                }

                foreach (DataRow oldProfileRow in oldDS.Tables["R_Video_Profiles"].Rows)
                {
                    DataRow[] matchingRowsInLiveSetupDb = liveDS.Tables["R_Video_Profiles"].Select("ProfileName = \'" + oldProfileRow["ProfileName"].ToString() + "\'");

                    if (matchingRowsInLiveSetupDb.Length == 0)
                    {
                        liveSetupCommand.CommandText = "INSERT INTO R_Video_Profiles (ProfileName, WMVVersion, Bitrate, VideoType, MaxKeyFrameSpacing, Quality, VBR) Values (\"" +
                                                       oldProfileRow["ProfileName"].ToString() + "\", \"" +
                                                       oldProfileRow["WMVVersion"].ToString() + "\", " +
                                                       oldProfileRow["Bitrate"].ToString() + ", \"" +
                                                       oldProfileRow["VideoType"].ToString() + "\", " +
                                                       oldProfileRow["MaxKeyFrameSpacing"].ToString() + ", " +
                                                       oldProfileRow["Quality"].ToString() + ", " +
                                                       oldProfileRow["VBR"].ToString() + ")";

                        liveSetupCommand.ExecuteNonQuery();
                    }

                    else
                    {
                        liveSetupCommand.CommandText = "UPDATE R_Video_Profiles SET " +
                                                       "WMVVersion = \"" + oldProfileRow["WMVVersion"].ToString() + "\", " +
                                                       "Bitrate = " + oldProfileRow["Bitrate"].ToString() + ", " +
                                                       "VideoType = \"" + oldProfileRow["VideoType"].ToString() + "\", " +
                                                       "MaxKeyFrameSpacing = " + oldProfileRow["MaxKeyFrameSpacing"].ToString() + ", " +
                                                       "Quality = " + oldProfileRow["Quality"].ToString() + ", " +
                                                       "VBR = " + oldProfileRow["VBR"].ToString() + " " +
                                                       "WHERE ProfileName = \'" + oldProfileRow["ProfileName"].ToString() + "\'";

                        liveSetupCommand.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                string error = $"Failed to persist video recording profiles (R_Video_Profiles Table) to new setup.mdb: {ex.Message}";
                log(error);
                throw new InstallationFailedException(error);
            }
            finally
            {
                if (liveSetupCommand.Connection != null && liveSetupCommand.Connection.State != ConnectionState.Closed)
                {
                    liveSetupCommand.Connection.Close();
                }
            }
        }

        private void setAccessPermissionsForProgramFolders()
        {
            try
            {
                var defaultDirSecurity = mainProgramDirectoryInfo.GetAccessControl();
                defaultDirSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, null),
                                                 FileSystemRights.Write | FileSystemRights.ReadAndExecute | FileSystemRights.Modify | FileSystemRights.CreateFiles |
                                                 FileSystemRights.CreateDirectories | FileSystemRights.DeleteSubdirectoriesAndFiles | FileSystemRights.ListDirectory |
                                                 FileSystemRights.AppendData | FileSystemRights.ExecuteFile | FileSystemRights.ReadAndExecute | FileSystemRights.Read |
                                                 FileSystemRights.ReadPermissions | FileSystemRights.Traverse,
                                                 InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                                 0,
                                                 AccessControlType.Allow));

                defaultDirSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.AuthenticatedUserSid, null),
                                 FileSystemRights.Write | FileSystemRights.ReadAndExecute | FileSystemRights.Modify | FileSystemRights.CreateFiles |
                                 FileSystemRights.CreateDirectories | FileSystemRights.DeleteSubdirectoriesAndFiles | FileSystemRights.ListDirectory |
                                 FileSystemRights.AppendData | FileSystemRights.ExecuteFile | FileSystemRights.ReadAndExecute | FileSystemRights.Read |
                                 FileSystemRights.ReadPermissions | FileSystemRights.Traverse,
                                 InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                 0,
                                 AccessControlType.Allow));

                mainProgramDirectoryInfo.SetAccessControl(defaultDirSecurity);

                if (Directory.Exists(@"C:\IT Projects") == false)
                {
                    Directory.CreateDirectory(@"C:\IT Projects");
                }

                DirectoryInfo itProjectInfo = new DirectoryInfo(@"C:\IT Projects");

                defaultDirSecurity = itProjectInfo.GetAccessControl();

                defaultDirSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, null),
                                             FileSystemRights.Write | FileSystemRights.ReadAndExecute | FileSystemRights.Modify | FileSystemRights.CreateFiles |
                                             FileSystemRights.CreateDirectories | FileSystemRights.DeleteSubdirectoriesAndFiles | FileSystemRights.ListDirectory |
                                             FileSystemRights.AppendData | FileSystemRights.ExecuteFile | FileSystemRights.ReadAndExecute | FileSystemRights.Read |
                                             FileSystemRights.ReadPermissions | FileSystemRights.Traverse,
                                             InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                                             0,
                                             AccessControlType.Allow));

                defaultDirSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.AuthenticatedUserSid, null),
                             FileSystemRights.Write | FileSystemRights.ReadAndExecute | FileSystemRights.Modify | FileSystemRights.CreateFiles |
                             FileSystemRights.CreateDirectories | FileSystemRights.DeleteSubdirectoriesAndFiles | FileSystemRights.ListDirectory |
                             FileSystemRights.AppendData | FileSystemRights.ExecuteFile | FileSystemRights.ReadAndExecute | FileSystemRights.Read |
                             FileSystemRights.ReadPermissions | FileSystemRights.Traverse,
                             InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                             0,
                             AccessControlType.Allow));

                itProjectInfo.SetAccessControl(defaultDirSecurity);

            }
            catch (Exception ex)
            {
                string error = $@"Failed to set folder permissions for the ITpipes program directory and C:\IT Projects: {ex.Message}";
                log(error);
                throw new InstallationFailedException(error);
            }
        }

        private void populateIbakPathsInSetupDb(OleDbConnection curConnection)
        {
            using (OleDbCommand ibakPathCommand = new OleDbCommand(@"UPDATE S SET IBAK_Folder = 'C:\Program Files\InspectIT\ViewIT\'", curConnection))
            {
                try
                {
                    object ibakRegistryKeyNotFound = "NotFound";
                    object getKeyReturnValue = Registry.GetValue(@"HKEY_CURRENT_USER\Software\IBAK\PANORAMO\Scanner", "PathCtrlFile", ibakRegistryKeyNotFound);

                    if (getKeyReturnValue == null)
                    {
                        //If the key is not present, the IBAK_Folder field should be set to the location of ViewIT so PWVermessung.exe is available.
                        ibakPathCommand.CommandText = @"UPDATE S SET IBAK_Folder = 'C:\Program Files\InspectIT\ViewIT\'";
                        if (curConnection.State != ConnectionState.Open)
                        {
                            curConnection.Open();
                        }
                        ibakPathCommand.ExecuteNonQuery();
                        curConnection.Close();
                        ibakPathCommand.Dispose();
                        return;
                    }

                    string ibakIKAS32Location = (string)getKeyReturnValue;

                    if (ibakIKAS32Location != "NotFound")
                    {
                        //the Pano Scanner app location is located within the same parent directory as IKAS32, so if the key is found need to replace IKAS32\ with PanoScan\Panoramo_Scanner.exe
                        ibakPathCommand.CommandText = @"UPDATE S SET IBAK_CaptureEXEPath = '" + ibakIKAS32Location.Replace(@"IKAS32\", @"PanoScan\Panoramo_Scanner.exe") +
                                                        "',  IBAK_Folder = '" + ibakIKAS32Location + "\'";

                        if (curConnection.State != ConnectionState.Open)
                        {
                            curConnection.Open();
                        }

                        ibakPathCommand.ExecuteNonQuery();
                    }

                    curConnection.Close();
                }
                catch (Exception ex)
                {
                    string error = $"Failed to populate setup.mdb with Ibak reference paths: {ex.Message}";
                    log(error);
                    throw new InstallationFailedException(error);
                }
            }
        }


        private void addReportTempDirectory()
        {
            //Without ensuring that the _Temp directory exists in the ITpipes program folder, some setups will encounter a bizarre out-of-memory error when running reports in batch mode.

            try
            {
                if (Directory.Exists(_config.InstallationPath + @"\_Temp") == false)
                {
                    Directory.CreateDirectory(_config.InstallationPath + @"\_Temp");
                }
            }
            catch (Exception ex)
            {
                _logWriter.WriteLine(ex);
            }
        }

        /// <summary>
        /// Collection of files which must be re-registered to correct an issue which occurs on some computers where ITpipes can launch, but complains of active x controls it can't find/load and crashes.
        /// </summary>
        private static readonly string[] systemForceReRegFiles = { "mscomct2.ocx", "comdlg32.ocx", "comctl32.ocx", "mscomctl.ocx" };
        private void reRegisterProblematicActiveXControls()
        {
            foreach (string curFile in systemForceReRegFiles)
            {
                _regHelper.RegisterAsCom(Path.Combine(systemFolder, curFile), true);
            }
        }

        private void restoreMostRecentSettingsBackup(string[] backupFiles)
        {
            List<string> FileExtractErrors = new List<string>();

            if (backupFiles == null || backupFiles.Length == 0)
            {
                return;
            }

            try
            {
                FileInfo backupToRestore = new FileInfo(backupFiles[0]);

                foreach (string curBackup in backupFiles)
                {
                    FileInfo curBackupInfo = new FileInfo(curBackup);
                    if (curBackupInfo.CreationTimeUtc > backupToRestore.CreationTimeUtc)
                    {

                        backupToRestore = curBackupInfo;
                    }
                }

                using (Ionic.Zip.ZipFile curBackupZip = new Ionic.Zip.ZipFile(backupToRestore.FullName))
                {

                    if (Directory.Exists(_config.InstallationPath) == false)
                    {
                        Directory.CreateDirectory(_config.InstallationPath);
                    }

                    curBackupZip.ZipErrorAction = Ionic.Zip.ZipErrorAction.Skip;

                    foreach (Ionic.Zip.ZipEntry curEntry in curBackupZip.Entries)
                    {

                        if (curEntry.IsDirectory)
                        { //directories are handled per-file to prevent deleting other files in the ITpipes directory.
                            continue;
                        }

                        string newPath = Path.Combine(_config.InstallationPath, curEntry.FileName);

                        if (File.Exists(newPath))
                        {
                            File.Delete(newPath);
                        }

                        if (Directory.Exists(Path.GetDirectoryName(newPath)) == false)
                        {
                            Directory.CreateDirectory(Path.GetDirectoryName(newPath));
                        }

                        try
                        {
                            curEntry.Extract(_config.InstallationPath, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently);
                        }
                        catch
                        {
                            FileExtractErrors.Add(string.Format("Failed to extract file: {0}", curEntry.FileName));
                        }
                    }
                }
            }

            catch (Exception ex)
            {
            }

            if (FileExtractErrors.Count > 0)
            {
                if (Debugger.IsAttached)
                {
                    Debugger.Break();
                }
            }
        }

        public void Dispose()
        {
            if (_logWriter != null)
            {
                try
                {
                    _logWriter.Flush();
                    _logWriter.Dispose();
                }
                catch (ObjectDisposedException)
                {
                    //Nothing to worry about here. The deed is already done.
                }
            }
            disposed = true;
        }

        ~InstallationLogic()
        {
            if (!this.disposed)
            {
                Dispose();
            }
        }

        private class FolderRegistrationRules
        {
            bool RegisterFiles { get; set; }
            List<string> SkipRegisteringFilesNamed { get; set; } = new List<string>();
        }
    }

    public class InstallationRequest
    {
        public string InstallationPath { get; set; } = @"C:\Program Files\InspectIT";
        public bool OverwriteTemplates { get; set; } = false;
        public bool InstallerMode { get; set; } = false;
        public bool InstallQSB { get; set; } = false;
        public string BackupToRestore { get; set; } = null;
        public string VersionNumber { get; set; }
        public bool InstallMP4 { get; set; } = false;

        public InstallationRequest(string installPath, bool overwriteTemplates, bool forceReinstallation, bool installQSB, string versionNumber, bool installMP4)
        {
            InstallationPath = installPath;
            OverwriteTemplates = overwriteTemplates;
            InstallerMode = forceReinstallation;
            InstallQSB = installQSB;
            VersionNumber = versionNumber;
            InstallMP4 = installMP4;
        }

    }
}
