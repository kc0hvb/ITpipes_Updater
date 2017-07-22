//#define DEBUG_MESSAGEBOXES_ENABLED
//#define DISABLE_LOGO_BACKUPS
//#define ignoreFileWriteFailues

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Windows;
using Ionic.Zip;
using System.Security.AccessControl;
using System.Security.Principal;

namespace ITpipes_Updater.Util {
    public static class UtilFunctions {

        public static string FixSlashes(string input)
        {
            while (input.Contains(@"\\"))
            {
                input = input.Replace(@"\\", @"\");
            }

            return input;
        }

        public static bool CreateShortcut(string pathToTargetFile, string directoryReceivingShortcut, string shortcutFileName) {

            string errorMessage = null;

            if (Directory.Exists(directoryReceivingShortcut) == false) {

                errorMessage = string.Format("UtilFunctions.CreateShortcut was passed an invalid directory path: {0}\nError: Directory does not exist.", directoryReceivingShortcut);
                System.Diagnostics.Debug.WriteLine(errorMessage);

#if DEBUG_MESSAGEBOXES_ENABLED
                MessageBox.Show(errorMessage, "Cannot create shortcut");
#endif
                return false;
            }

            if (File.Exists(pathToTargetFile) == false) {

                errorMessage = string.Format("UtilFunctions.CreateShortcut was passed an invalid target file path: {0}\nError: The file does not exist.", pathToTargetFile);
                System.Diagnostics.Debug.WriteLine(errorMessage);

#if DEBUG_MESSAGEBOXES_ENABLED
                MessageBox.Show(errorMessage, "Cannot create shortcut");
#endif
                return false;
            }

            if (shortcutFileName.Contains('.')) {

                System.Diagnostics.Debug.WriteLine(string.Format("shortcutFileName passed to UtilFunctions.CreateShortcut (\"{0}\") had a file extension included. This file extension has been truncated."), shortcutFileName);
                shortcutFileName = Path.GetFileNameWithoutExtension(shortcutFileName);
            }

            try {

                IShellLink link = (IShellLink)new ShellLink();

                link.SetDescription("ITpipes Asset Inspection and Management");
                link.SetPath(pathToTargetFile);
                link.SetWorkingDirectory("");

                IPersistFile shortcutFile = (IPersistFile)link;
                shortcutFile.Save(Path.Combine(directoryReceivingShortcut, shortcutFileName + ".lnk"), false);
                return true;
            }
            catch (Exception ex) {

                errorMessage = string.Format("Failed to create short to\nFile: \"{0}\"\nIn directory: \"{1}\"\nDue to exception: {2}", pathToTargetFile, directoryReceivingShortcut, ex);
                System.Diagnostics.Debug.WriteLine(errorMessage);

#if DEBUG_MESSAGEBOXES_ENABLED
                MessageBox.Show(errorMessage, "Cannot create shortcut");
#endif
                return false;
            }
        }

        public static string GetAllUsersStartMenuFolder() {

            StringBuilder returnPathSB = new StringBuilder(260);
            SHGetSpecialFolderPath(IntPtr.Zero, returnPathSB, CSIDL_COMMON_STARTMENU, false);
            return returnPathSB.ToString();
        }

        public static string GetAllUsersDesktopFolder() {
            StringBuilder returnPathSB = new StringBuilder(260);
            SHGetSpecialFolderPath(IntPtr.Zero, returnPathSB, CSIDL_COMMON_DESKTOPDIRECTORY, false);
            return returnPathSB.ToString();
        }

        public static void copyFile(string sourceFile, string destinationFile, bool overwrite = true) {

            try {

                if (Directory.Exists(Path.GetDirectoryName(destinationFile)) == false) {
                    Directory.CreateDirectory(Path.GetDirectoryName(destinationFile));
                }

                if (File.Exists(destinationFile)) {
                    if (overwrite == false) {
                        return;
                    }

                    FileInfo destinationFileInfo = new FileInfo(destinationFile);
                    if ((destinationFileInfo.Attributes & FileAttributes.Hidden) == FileAttributes.Hidden) {
                        destinationFileInfo.Attributes = (destinationFileInfo.Attributes & ~FileAttributes.Hidden);
                    }

                    destinationFileInfo.Delete();
                }
                //Creates a backup of the setup.mdb.
                using (FileStream sourceStream = new FileStream(sourceFile, FileMode.Open))
                using (FileStream targetStream = new FileStream(destinationFile, FileMode.CreateNew)) {
                    sourceStream.CopyTo(targetStream);
                }
                

                //Have to set the creation time and last modified date manually--filestreams and fileinfo.copyto set these values to current time.
                File.SetCreationTime(destinationFile, File.GetCreationTime(sourceFile));
                File.SetLastWriteTime(destinationFile, File.GetLastWriteTime(sourceFile));
            }
            catch {

#if ignoreFileWriteFailues
#else
                throw;
#endif
            }
        }

        public static string CreateBackup() {

            //return value is the path to the backup file.

            string sBackupPath = ITP_Backup.CreateBackupOfAllConfigurations(getBackupDirectory());
            return sBackupPath;

        }


        public static string getBackupDirectory() {
            string sBackupDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), @"ITpipes\ITpipesConfigBackup");
            return sBackupDirectory;
        }

        public static string getBackupFileExtension() {
            return ITP_Backup.BACKUP_FILE_EXTENSION;
        }

        private enum UserFileRecoveryState {
            NoIssue,
            Recovered,
            Unrecoverable
        }


        public static void scanAndFixCorruptUserFiles(string pathToITpipesDirectory) {

            string pathToUsersDirectory = Path.Combine(pathToITpipesDirectory, "Users");
            
            if (Directory.Exists(pathToUsersDirectory) == false) {
                return;
            }

            
            string[] userFiles = Directory.GetFiles(pathToUsersDirectory, "*.usr");
            
            foreach (string curUser in userFiles) {

                var recoveryState = getUserFileRecovery(curUser);

                if (recoveryState == UserFileRecoveryState.Recovered) {

                    
                }
                else if (recoveryState == UserFileRecoveryState.Unrecoverable) {

                    //MessageBox.Show("Debug Message - Unrecoverable user: " + curUser);
                }
            }
        }

        private static UserFileRecoveryState getUserFileRecovery(string pathToUser) {

            string pathToTempDir = pathToUser + "_deleteme";
            
            try {

                using (ZipFile usrZip = new ZipFile(pathToUser)) {

                    usrZip.Password = "ITWCA321";
                    
                    foreach (ZipEntry curEntry in usrZip.Entries) {
                        
                        curEntry.Extract(pathToTempDir, ExtractExistingFileAction.OverwriteSilently);
                    }
                    return UserFileRecoveryState.NoIssue;
                }
            }
            catch (ZipException) {

                string curBackupFile = pathToUser.Replace(".usr", ".bak");

                if (!File.Exists(curBackupFile)) {
                    return UserFileRecoveryState.Unrecoverable;
                }

                try {
                    copyFile(curBackupFile, pathToUser, true);
                    return UserFileRecoveryState.Recovered;
                }
                catch (Exception) {
                    return UserFileRecoveryState.Unrecoverable;
                }
            }
            finally {
                if (Directory.Exists(pathToTempDir)) {
                    Directory.Delete(pathToTempDir, true);
                }
            }
        }



        public static long GetDirectorySizeInBytes(string pathToDirectory)
        {
            string[] allFilesInDirectory = Directory.GetFiles(pathToDirectory, "*.*", SearchOption.AllDirectories);

            long totalSize = 0;

            foreach (string curFile in allFilesInDirectory)
            {
                FileInfo curFileInfo = new FileInfo(curFile);

                totalSize += curFileInfo.Length;
            }

            return totalSize;
        }




        #region COM Imports

        //Using this to ensure backwards compatibility with XP+. We don't technically support XP, but people are going to run the updater on XP systems anyway--it's already happening.
        [DllImport("shell32.dll")]
        static extern bool SHGetSpecialFolderPath(IntPtr hwndOwner,
        [Out] StringBuilder lpszPath, int nFolder, bool fCreate);
        const int CSIDL_COMMON_STARTMENU = 0x16;  // All Users\Start Menu
        const int CSIDL_COMMON_DESKTOPDIRECTORY = 0x19; //All Users\Desktop

#endregion COM Imports

    }

    class ITP_Backup {


        private static string _itpipesInstallDir = null;

        public static readonly string BACKUP_FILE_EXTENSION = "itbak";

        //these are stored as paths relative to the ITpipes installation directory.
        private static readonly string[] FILES_TO_BACKUP = {
            @"blankdb\project.mdb",
            @"blankdb\wmv.prx",
            @"blankdb\wmv_dvd.prx",
            @"blankdb\wmv9.prx",
            @"blankdb\wmv9_dvd.prx",
            @"ViewIT\ViewIT.exe.config",
            @"address.mdb",
            @"setup.mdb",
            @"merge.ini"
        };

        private static readonly string[] DIRECTORIES_TO_BACKUP = {
            "Logos",
            "Data Validation Rulesets",
            "Users",
            "Templates"
        };


        public static string CreateBackupOfAllConfigurations(string directoryToSaveBackupTo) {

            string _itpipesDirectory = getITpipesInstallationDirectory();

            if (_itpipesDirectory == null) {
                _itpipesDirectory = @"C:\Program Files\InspectIT";
            }

            if (Directory.Exists(_itpipesDirectory) == false) {
                throw new FileNotFoundException(string.Format("Could not locate ITpipes installation directory: \"{0}\"", _itpipesDirectory));
            }

            if (Directory.Exists(directoryToSaveBackupTo) == false) {
                Directory.CreateDirectory(directoryToSaveBackupTo);
            }

            //pre-backup processing:
            string setupMdb = Path.Combine(_itpipesDirectory, "setup.mdb");

            if (File.Exists(setupMdb)) {

                UtilFunctions.copyFile(setupMdb, setupMdb + ".bak", true);
            }
#if DISABLE_LOGO_BACKUPS
#else
            transferAddressBookLogosToLogosDirectory(_itpipesDirectory);
            transferTemplateLogosToLogosDirectory(Path.Combine(_itpipesDirectory, @"Templates"));
#endif

            string newBackupZipLocation = Path.Combine(directoryToSaveBackupTo, string.Format("ITpipes Config Backup - {0}.{1}", DateTime.Now.ToLongDateString() + "_" + DateTime.Now.Ticks.ToString(), BACKUP_FILE_EXTENSION));

            int backupNum = 0;

            if (File.Exists(newBackupZipLocation)) {

                string tempFileName = new string(newBackupZipLocation.ToCharArray());

                while (true) {
                    backupNum++;
                    tempFileName = appendNumberToExistingFileRecord(newBackupZipLocation, backupNum);
                    if (File.Exists(tempFileName) == false) {
                        newBackupZipLocation = tempFileName;
                        break;
                    }
                }
            }
            
            using (ZipFile backupZip = new ZipFile(newBackupZipLocation)) {

                foreach (string curDirectory in DIRECTORIES_TO_BACKUP) {
                    addDirectoryToZipObject(backupZip, _itpipesDirectory, _itpipesDirectory + @"\" + curDirectory);
                }

                foreach (string curFile in FILES_TO_BACKUP) {
                    addFileToZipObjectIfPathExists(backupZip, _itpipesDirectory, _itpipesDirectory + @"\" + curFile);
                }

                //dotnetzip doesn't handle its temp file's name already existing:
                if (File.Exists(newBackupZipLocation + ".tmp")) {
                    File.Delete(newBackupZipLocation + ".tmp");
                }

                if (Directory.Exists(Path.GetDirectoryName(newBackupZipLocation)) == false) {
                    Directory.CreateDirectory(Path.GetDirectoryName(newBackupZipLocation));
                }



                setBackupDirectoryAccessPermissions(directoryToSaveBackupTo);

                try {
                    backupZip.Save();
                }
                catch {
                    //TODO: log this somehow. It's entirely possible that the backup failed because there were no extant files to add to the backup zip.
                }
            }


            return newBackupZipLocation;
        }

        private static string copyTemplateLogoFromExistingLocation(string relativePathToExistingLogo, string newLogoDirectory, string templateName, string contactName) {

            System.Text.RegularExpressions.Regex filenameRegex = new System.Text.RegularExpressions.Regex("[^a-z^A-Z^0-9]");

            if (_itpipesInstallDir == null) {
                _itpipesInstallDir = getITpipesInstallationDirectory();
            }

            string curLogoFullPath = Path.Combine(_itpipesInstallDir, relativePathToExistingLogo);
            string newFileName = string.Format(@"{0}\Logos\{1}.{2}.bmp", _itpipesInstallDir, templateName, filenameRegex.Replace(contactName, string.Empty));
            string newFileRelativePath = newFileName.Replace(_itpipesInstallDir, "");

            if (File.Exists(newFileName)) {
                return newFileName;
            }

            if (Path.GetExtension(relativePathToExistingLogo).ToUpper().Contains("BMP")) {

                UtilFunctions.copyFile(curLogoFullPath, newFileName);
            }

            else {

                using (Image convertImage = Image.FromFile(curLogoFullPath)) {

                    convertImage.Save(newFileName, System.Drawing.Imaging.ImageFormat.Bmp);
                }
            }


            return newFileRelativePath;
        }

        private static void transferAddressBookLogosToLogosDirectory(string pathToItInstallDir) {

            string addressMdbPath = Path.Combine(pathToItInstallDir, "address.mdb");
            _itpipesInstallDir = getITpipesInstallationDirectory();

            if (File.Exists(addressMdbPath)) {

                using (OleDbConnection curOleConn = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + addressMdbPath))
                using (OleDbCommand curOleCommand = curOleConn.CreateCommand())
                using (DataTable curAddressBook = new DataTable()) {

                    curOleConn.Open();

                    curOleCommand.CommandText = "SELECT * FROM [Address]";

                    curAddressBook.Load(curOleCommand.ExecuteReader());
                    if (curAddressBook.Rows.Count == 0) {
                        return;
                    }

                    foreach (DataRow curRow in curAddressBook.Rows) {

                        if (curRow["Logo"] != DBNull.Value && File.Exists(Path.Combine(_itpipesInstallDir, (string)curRow["Logo"])) && curRow["Contact_Name"] != DBNull.Value) {

                            string relativePathToLogoFile = (string)curRow["Logo"];

                            string relativePathToNewLogoFile = copyTemplateLogoFromExistingLocation(relativePathToLogoFile, Path.Combine(pathToItInstallDir, "Logos"), "AddressBook", (string)curRow["Contact_Name"]);
                            //UtilFunctions.copyFile(pathToLogoFile, pathToNewLogoFile, true);

                            if (relativePathToNewLogoFile == relativePathToLogoFile) {
                                continue;
                            }


                            curOleCommand.CommandText = "UPDATE [Address] SET [Logo] = ? WHERE [Address_ID] = ?";

                            curOleCommand.Parameters.Clear();
                            curOleCommand.Parameters.AddWithValue("Logo", relativePathToNewLogoFile);
                            curOleCommand.Parameters.AddWithValue("Info_ID", (int)curRow["Address_ID"]);

                            curOleCommand.ExecuteNonQuery();
                        }
                    }
                }
            }
        }

        private static void transferTemplateLogosToLogosDirectory(string pathToItInstallDir) {

            string pathToTemplatesDir = Path.Combine(pathToItInstallDir, "Templates");

            if (Directory.Exists(pathToTemplatesDir) == false) {
                return;
            }

            string tempTemplates = Path.Combine(Path.GetTempPath(), @"ITP_Temp_Templates");

            if (Directory.Exists(tempTemplates)) {
                Directory.Delete(tempTemplates);
            }

            Directory.CreateDirectory(tempTemplates);

            foreach (string curTemplate in Directory.GetFiles(pathToTemplatesDir, "*.tpl", SearchOption.TopDirectoryOnly)) {

                string tempTplFilePath = Path.Combine(tempTemplates, Path.GetFileName(curTemplate));

                UtilFunctions.copyFile(curTemplate, tempTplFilePath, true);

                using (ZipFile tplZip = new ZipFile(tempTplFilePath)) {

                    tplZip.Encryption = EncryptionAlgorithm.PkzipWeak;
                    tplZip.Password = "ITWCA321";
                    tplZip.ExtractAll(tempTemplates, ExtractExistingFileAction.OverwriteSilently);

                    string tempProjectDbPath = Path.Combine(tempTemplates, "project.mdb");

                    if (File.Exists(tempProjectDbPath) == false) {
                        continue;
                    }

                    using (OleDbConnection curOleConn = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + tempProjectDbPath))
                    using (OleDbCommand curOleCommand = curOleConn.CreateCommand())
                    using (DataTable curAddressBook = new DataTable()) {

                        curOleConn.Open();

                        curOleCommand.CommandText = "SELECT * FROM [Info]";

                        curAddressBook.Load(curOleCommand.ExecuteReader());

                        if (curAddressBook.Rows.Count == 0) {
                            continue;
                        }

                        foreach (DataRow curRow in curAddressBook.Rows) {

                            string relativePathToLogoFile = curRow["Logo"] == DBNull.Value ? null : (string)curRow["Logo"];

                            if (relativePathToLogoFile != null && File.Exists(Path.Combine(pathToItInstallDir, relativePathToLogoFile)) && curRow["Contact_Name"] != DBNull.Value) {

                                string relativePathToNewLogoFile =
                                    copyTemplateLogoFromExistingLocation(
                                        relativePathToLogoFile,
                                        Path.Combine(pathToItInstallDir, "Logos"),
                                        Path.GetFileNameWithoutExtension(curTemplate),
                                        (string)curRow["Contact_Name"]);




                                //Path.Combine(_itpipesInstallDir,
                                //    string.Format(@"Logos\{0}.{1}.bmp",
                                //    Path.GetFileNameWithoutExtension(curTemplate),
                                //    (string)curRow["Contact_Name"]));

                                if (relativePathToNewLogoFile == relativePathToLogoFile) {

                                    continue;
                                }

                                curOleCommand.CommandText = "UPDATE [Info] SET [Logo] = ? WHERE [Info_ID] = ?";

                                curOleCommand.Parameters.Clear();
                                curOleCommand.Parameters.AddWithValue("Logo", relativePathToNewLogoFile);
                                curOleCommand.Parameters.AddWithValue("Info_ID", (int)curRow["Info_ID"]);

                                curOleCommand.ExecuteNonQuery();
                            }
                        }

                        ZipEntry projectZipEntry = null;

                        foreach (ZipEntry curEntry in tplZip.Entries) {

                            if (curEntry.FileName == "project.mdb") {

                                projectZipEntry = curEntry;
                                break;
                            }
                        }

                        if (projectZipEntry != null) {
                            tplZip.RemoveEntry(projectZipEntry);
                        }
                        else {
                            throw new Exception(string.Format("Could not locate project.mdb in template: {0}", curTemplate));
                        }

                        tplZip.AddFile(tempProjectDbPath, @"\");

                        foreach (ZipEntry curEntry in tplZip.Entries) {
                            curEntry.Password = "ITWCA321";
                        }

                        try {
                            if (File.Exists(curTemplate + ".bak")) {
                                File.Delete(curTemplate + ".bak");
                            }

                            File.Move(curTemplate, curTemplate + ".bak");

                            tplZip.Save(curTemplate);

                            File.Delete(curTemplate + ".bak");
                        }
                        catch {
                            throw;
                        }
                    }
                }
            }

            try {

                Directory.Delete(tempTemplates, true);
            }
            catch {
                //Not gonna sweat it.
            }
        }


        private static void setBackupDirectoryAccessPermissions(string directoryToSaveBackupTo) {


            DirectoryInfo itBackupInfo = new DirectoryInfo(directoryToSaveBackupTo);

            var defaultDirSecurity = itBackupInfo.GetAccessControl();

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

            itBackupInfo.SetAccessControl(defaultDirSecurity);
        }


        private static string appendNumberToExistingFileRecord(string fileName, int num) {
            int extensionIndex = fileName.LastIndexOf('.');
            if (extensionIndex == -1) { extensionIndex = fileName.Length; }
            string fileNameOnly = fileName.Substring(0, extensionIndex);
            string extensionOnly = string.Empty;
            if (extensionIndex != fileName.Length) { extensionOnly = fileName.Substring(extensionIndex, fileName.Length - extensionIndex); }
            fileNameOnly += ("_" + num.ToString());
            return fileNameOnly + extensionOnly;
        }

        private static void addFileToZipObjectIfPathExists(ZipFile zipfile, string itpipesDirectory, string pathToFile) {

            if (File.Exists(pathToFile)) {

                string relativePath = Path.GetDirectoryName(pathToFile.Replace(itpipesDirectory, ""));

                zipfile.AddFile(pathToFile, relativePath);
            }
        }

        private static void addDirectoryToZipObject(ZipFile curZip, string itpipesDirectory, string pathToDirectory) {

            if (Directory.Exists(pathToDirectory)) {

                string relativePath = pathToDirectory.Replace(itpipesDirectory, "");

                if (relativePath == pathToDirectory) {

                    //Can't add this to the directory because it does not exist within the ITpipes directory.
                    return;
                }

                curZip.AddDirectory(pathToDirectory, relativePath);
                return;
            }
        }

        public static string getITpipesInstallationDirectory(string installPathOverride = null) {
            if (installPathOverride != null && Directory.Exists(installPathOverride)) {

                return installPathOverride;
            }

            RegistryKey installLocationRegKey_x64 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node\ITpipes", false);
            RegistryKey installLocationRegKey_x32 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\ITpipes", false);


            if (installLocationRegKey_x64 != null) {

                object installDirObj = installLocationRegKey_x64.GetValue("Install_Dir");

                if (installDirObj == null) {
                    return null;
                }

                return (string)installDirObj;
            }
            else if (installLocationRegKey_x32 != null) {
                object installDirObj = installLocationRegKey_x32.GetValue("Install_Dir");

                if (installDirObj == null) {
                    return null;
                }

                return (string)installDirObj;
            }

            return null;
        }

        public static void RestoreBackup(string installPath, string pathToBackupFile)
        {
            if (!Directory.Exists(installPath))
            {
                Directory.CreateDirectory(installPath);
            }

            using (ZipFile backupZip = ZipFile.Read(pathToBackupFile))
            {
                backupZip.ExtractAll(installPath, ExtractExistingFileAction.OverwriteSilently);


                foreach (var curItem in backupZip.Entries)
                {
                    if (curItem.IsDirectory)
                    {
                        continue;
                    }

                    string fullPath = Path.Combine(installPath, curItem.FileName);
                    string dirPath = Path.GetDirectoryName(fullPath);
                    if (!Directory.Exists(dirPath))
                    {
                        Directory.CreateDirectory(dirPath);
                    }
                    curItem.Extract(installPath, ExtractExistingFileAction.OverwriteSilently);
                }
            }
        }

    }

#region COM Imports

    [ComImport]
    [Guid("00021401-0000-0000-C000-000000000046")]
    internal class ShellLink {
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("000214F9-0000-0000-C000-000000000046")]
    internal interface IShellLink {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, out IntPtr pfd, int fFlags);
        void GetIDList(out IntPtr ppidl);
        void SetIDList(IntPtr pidl);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
        void GetHotkey(out short pwHotkey);
        void SetHotkey(short wHotkey);
        void GetShowCmd(out int piShowCmd);
        void SetShowCmd(int iShowCmd);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, int dwReserved);
        void Resolve(IntPtr hwnd, int fFlags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
    }

#endregion
}
