#define DEBUG


using System;
using System.IO;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Resources;
using System.Windows.Data;
using Microsoft.Win32;
using Ionic.Zip;
using ITpipes_Uninstaller.Util;
using System.Diagnostics;

namespace ITpipes_Uninstaller.Uninstaller {
    class Uninstaller {

        //These are files which are absolutely known to be created for, or only used by, ITPipes.
        private static string[] ITPIPES_ONLY_LIBRARIES = {
            "BrowseFor.dll", "DemoSource.ax", "Deinterlace.ax", "H264Encoder.ax", "Imapi2.Interop.dll", "Interop.IBAKBiPanoView.dll",
            "IT_Aries3000.ocx", "ITBurner2.tlb", "ITDB.dll", "ITFootageCounter.dll", "ITFunctions.dll", "ITImportExport.dll", "ITLinkIT.dll",
            "ITLinkITServer.dll", "ITMerge.dll", "ITPACP.dll", "IT_InternetControls.dll", "IT_IpekCounter.dll", "IT_JPEGResizer.ocx",
            "IT_MeasureIT.ocx", "IT_PANO3D.ocx", "IT_SubCam_Inf.dll", "ITBurner2.dll", "ITVideoConverter.ocx", "ITVideoEdit.dll",
            "VCRControl.ocx", "ITpipesSidescanProcessingControl.dll", "ITTemplate.dll", "ITUpdate.dll",
            "BOB4Driver.dll", "EDE7Driver.dll", "IBAKControls.dll", "IPEKDriver.dll"
        };

        private static string[] REGISTERABLE_FILE_TYPES = {
            ".dll", ".ax", ".ocx"
        };

        private static string _itpipesInstallDir = null;

        public static bool UninstallItpipes(string programInstallDirectoryOverride = null, bool forceUninstall = true) {
            
            if (doesProcessHaveAdminAccess() == false) {
                throw new PrivilegeNotHeldException("Uninstallation cannot proceed without administrator permissions. Please uninstall ITpipes through Control Panel.");
            }

            try {
                CloseITpipesIfRunning();
            }
            catch (Exception ex) {
                if (forceUninstall == false) {
                    throw;
                }
#if DEBUG
                Debugger.Break();
#endif
            }

            _itpipesInstallDir = getITpipesInstallationDirectory(programInstallDirectoryOverride);

            if (_itpipesInstallDir == null && forceUninstall == false) {
                
                throw new ProgramNotInstalledException();
            }

            if (Directory.Exists(_itpipesInstallDir) == false && forceUninstall == false) {
                throw new DirectoryNotFoundException("ITpipes installation directory does not exist.");
            }

            try {
                //Backup all ITpipes configurations - The path is taken from the default path for the config's backup model:
                UtilFunctions.CreateBackup();
                //CreateBackupOfAllConfigurations(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), @"ITpipes\ITpipesConfigBackup"));
            }
            catch (Exception) {
                if (forceUninstall == false) {
                    throw;
                }
            }

            try {
                SafelyDeleteLibraries();
            }
            catch (Exception) {
                if (forceUninstall == false) {
                    throw;
                }
            }
            try {
                SafelyDeleteDirectory(_itpipesInstallDir);
            }
            catch (Exception) {
                if (forceUninstall == false) {
                    throw;
                }
            }
            try {
                DeleteITpipesRegistryKeys();
            }
            catch (Exception) {
                if (forceUninstall == false) {
                    throw;
                }
#if DEBUG
                System.Diagnostics.Debugger.Break();
#endif
            }

            try {
                DeleteProgramShortcuts();
            }
            catch (Exception) {
                if (forceUninstall == false) {
                    throw;
                }
#if DEBUG
                Debugger.Break();
#endif
            }
            

            return true;
        }

        private static void CloseITpipesIfRunning() {

            Process[] liveProcesses = Process.GetProcessesByName("ITpipes");

                foreach (Process curProc in liveProcesses) {
                    curProc.Kill();
                while (curProc.HasExited == false) {
                    System.Threading.Thread.Sleep(20);
                }
                }
        }

        private static void DeleteProgramShortcuts() {
            
            if (File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "ITpipes.lnk"))) {

                File.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ITpipes.lnk"));
            }

            if (File.Exists(Path.Combine(UtilFunctions.GetAllUsersDesktopFolder(), "ITpipes.lnk"))) {

                File.Delete(Path.Combine(UtilFunctions.GetAllUsersDesktopFolder(), "ITpipes.lnk"));
            }

            if (File.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "InspectIT.lnk"))) {

                File.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "InspectIT.lnk"));
            }

            if (File.Exists(Path.Combine(UtilFunctions.GetAllUsersDesktopFolder(), "InspectIT.lnk"))) {

                File.Delete(Path.Combine(UtilFunctions.GetAllUsersDesktopFolder(), "InspectIT.lnk"));
            }

            if (Directory.Exists(Path.Combine(UtilFunctions.GetAllUsersStartMenuFolder(), "ITpipes"))) {

                Directory.Delete(Path.Combine(UtilFunctions.GetAllUsersStartMenuFolder(), "ITpipes"), true);
            }
        }

        private static void SafelyDeleteLibraries() {

            if (_itpipesInstallDir == null) {
                return;
            }

            string libDir = Path.Combine(_itpipesInstallDir, "Libraries");

            if (Directory.Exists(libDir)) {


                string[] libFiles = Directory.GetFiles(libDir);

                foreach (string curLibFile in libFiles) {

                    if (REGISTERABLE_FILE_TYPES.Contains(Path.GetExtension(curLibFile))) {

                        if (ITPIPES_ONLY_LIBRARIES.Contains(Path.GetFileNameWithoutExtension(curLibFile))) {

                            unregisterFile(curLibFile);
                        }
                        else {

                            //need to save the file--don't want to break other programs that rely on common libraries...

                            moveAndRegisterFile(curLibFile);
                        }
                    }
                }
            }

            string driversDir = Path.Combine(_itpipesInstallDir, "Drivers");

            if (Directory.Exists(driversDir)) {

                string[] driversFiles = Directory.GetFiles(driversDir);

                foreach (string curDrvFile in driversFiles) {

                    if (REGISTERABLE_FILE_TYPES.Contains(Path.GetExtension(curDrvFile))) {

                        if (ITPIPES_ONLY_LIBRARIES.Contains(Path.GetFileNameWithoutExtension(curDrvFile))) {

                            unregisterFile(curDrvFile);
                        }
                        else {

                            //need to save the file--don't want to break other programs that rely on common libraries...

                            moveAndRegisterFile(curDrvFile);
                        }
                    }
                }
            }
        }

        private static void SafelyDeleteDirectory(string dirToDel) {

            if (dirToDel == null) {
                return;
            }

            try {
                Directory.Delete(dirToDel, true);
            }
            catch {
                string[] allFilesInDirStructure = Directory.GetFiles(dirToDel, "*.*", SearchOption.AllDirectories);

                foreach (string curFile in allFilesInDirStructure) {
                    try {
                        File.Delete(curFile);
                    }
                    catch {
                        //don't handle individual failures at the moment--doesn't really matter to me
                        //Any files remaining after this process will be the files I'm unable to delete.
                    }
                }
            }
        }

        private static void DeleteITpipesRegistryKeys() {

            using (RegistryKey softwareLocationRegKey32 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE", true)) {

                string[] softwareSubkeys = softwareLocationRegKey32.GetSubKeyNames();

                try {
                    softwareLocationRegKey32.DeleteSubKeyTree("ITpipes");
                }
                catch {
                    //not handling this here
                }

                softwareLocationRegKey32.Dispose();
            }

            using (RegistryKey softwareLocationRegKey64 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node", true)) {

                string[] softwareSubkeys = softwareLocationRegKey64.GetSubKeyNames();

                try {
                    softwareLocationRegKey64.DeleteSubKeyTree("ITpipes");
                }
                catch {
                    //not handling this here
                }

                softwareLocationRegKey64.Dispose();
            }

            //Also need to remove the uninstaller information, or ITpipes will show up in the installed programs list.
            //removing this for the Nullsoft installer's info as well--just to be sure.
            using (RegistryKey uninstallerRegKey32 = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall", true)) {

                string[] installedPrograms = uninstallerRegKey32.GetSubKeyNames();

                if (installedPrograms.Contains("ITpipes")) {

                    try {
                        uninstallerRegKey32.DeleteSubKeyTree("ITpipes");
                    }
                    catch {

                    }
                }
            }

            using (RegistryKey uninstallerRegKey64 = Registry.LocalMachine.OpenSubKey(@"Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall", true)) {

                string[] installedPrograms = uninstallerRegKey64.GetSubKeyNames();

                if (installedPrograms.Contains("ITpipes")) {

                    try {
                        uninstallerRegKey64.DeleteSubKeyTree("ITpipes");
                    }
                    catch {

                    }
                }
            }
        }


        private static int unregisterFile(string fileToUnreg) {

            int returnInt = -1;

            using (System.Diagnostics.Process regsvr32Process = new System.Diagnostics.Process()) {

                regsvr32Process.StartInfo.FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.SystemX86), "regsvr32.exe");
                regsvr32Process.StartInfo.CreateNoWindow = true;
                regsvr32Process.StartInfo.RedirectStandardOutput = true;
                regsvr32Process.StartInfo.UseShellExecute = false;

                regsvr32Process.StartInfo.Arguments = "/s /u \"" + fileToUnreg + "\"";
                regsvr32Process.Start();
                regsvr32Process.WaitForExit();
                returnInt = regsvr32Process.ExitCode;
            }

            return returnInt;
        }

        private static int moveAndRegisterFile(string fileToSaveAndReg) {

            int returnInt = -1; //file registration return code

            string curFileDirectory = Path.GetDirectoryName(fileToSaveAndReg);
            string systemDirectory = Environment.GetFolderPath(Environment.SpecialFolder.SystemX86);

            string newFileLocation = fileToSaveAndReg.Replace(curFileDirectory, systemDirectory);

            if (File.Exists(newFileLocation) == false) {

                File.Copy(fileToSaveAndReg, newFileLocation, false);
            }

            try {
                File.Delete(fileToSaveAndReg);
            }
            catch {
                //don't really care right now
            }

            using (System.Diagnostics.Process regsvr32Process = new System.Diagnostics.Process()) {

                regsvr32Process.StartInfo.FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.SystemX86), "regsvr32.exe");
                regsvr32Process.StartInfo.CreateNoWindow = true;
                regsvr32Process.StartInfo.RedirectStandardOutput = true;
                regsvr32Process.StartInfo.UseShellExecute = false;

                regsvr32Process.StartInfo.Arguments = string.Format("/s \"{0}\"", newFileLocation);
                regsvr32Process.Start();
                regsvr32Process.WaitForExit();
                returnInt = regsvr32Process.ExitCode;
            }

            //need to handle the possibility that the file is a .Net assembly, and not a COM library. If COM registration was successful, the return code will have been 0:
            if (returnInt != 0) {

                System.Reflection.Assembly curAssembly = null;

                try {

                    curAssembly = System.Reflection.Assembly.LoadFrom(newFileLocation);
                }
                catch {
                    return returnInt;
                }
                
                System.Runtime.InteropServices.RegistrationServices regAsm = new System.Runtime.InteropServices.RegistrationServices();
                
                try { regAsm.RegisterAssembly(curAssembly, System.Runtime.InteropServices.AssemblyRegistrationFlags.SetCodeBase); }
                catch {
                    try {
                        regAsm.RegisterAssembly(curAssembly, System.Runtime.InteropServices.AssemblyRegistrationFlags.None);

                    }
                    catch {
                        return returnInt;
                    }
                }
            }


            return returnInt;
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
            else if ( installLocationRegKey_x32 != null) {
                object installDirObj = installLocationRegKey_x32.GetValue("Install_Dir");

                if (installDirObj == null) {
                    return null;
                }

                return (string)installDirObj;
            }

            return null;
        }


        private static bool doesProcessHaveAdminAccess() {
            //Ran into issues testing permissions on systems with UAC disabled. The only consistent test I could find was just to verify that I have write access to
            //protected folders.

            //If non-administrator users have write permission to Program Files and Windows\System32 directories, then these people are beyond our help.

            //TODO: Add a regsvr32 attempt to verify that files are registerable.

            bool returnValue = false;

            string testFilePath = @"C:\Program Files\test.txt";
            try {
                File.Create(testFilePath).Close();
                File.Delete(testFilePath);

                testFilePath = @"C:\Windows\System32\test.txt";
                File.Create(testFilePath).Close();
                File.Delete(testFilePath);

                returnValue = true;
            }
            catch {
                returnValue = false;
            }

            return returnValue;
        }
    }

    public class ProgramNotInstalledException : Exception { }

}
