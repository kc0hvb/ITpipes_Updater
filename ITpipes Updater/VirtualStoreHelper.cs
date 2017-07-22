using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using ITpipes_Updater.Util;
using static ITpipes_Updater.Util.UtilFunctions;

namespace ITpipes_Updater
{
    public static class VirtualStoreHelper
    {
        //Virtual Store is a location used to store changes to protected directories (like Program Files, where ITpipes is forced to be installed).
        //Basically, when UAC was added (Vista) many applications used protected system directories to store configuration and user data (like ITpipes does)
        //Program Files is now only writable for admin uses, so to prevent older applications from crashing, any time a user account is used to run an application
        //that attempts to write to a protected folder, the write is performed to a "virtualstore" folder in the user's appdata->local directory
        //Because the update is run as an administrator, it's running in a different user context than ITpipes will be run, meaning that Windows will not silently call the
        //VirtualStore files--it will directly check for file existence and file contents from the Program Files directory
        //This causes havoc after the update is complete because ITpipes will see files (like templates) that don't actually exist in C:\Program Files\InspectIT
        //The best way of handling this is to attempt to integrate the VirtualStore files, then move them so Windows doesn't use them anymore.
        //After this updater is used, the ITpipes application directory will allow user-level write access, so virtualStore won't be forced to be used for new/modified files anymore.

        //Note: this is only a concern with updating old installations--all new installations will never use VirtualStore because write permissions are granted to the application directory at install-time

        public static void IntegrateDirectory(string directoryPath)
        {
            //ntfs only allows : chars in drive specifiers--since VS only applies to dirs on the system volume we can just strip the drive letter
            string relativePath = directoryPath.Contains(":") ? directoryPath.Substring(directoryPath.IndexOf(':') + 1) : directoryPath;

            List<string> allVsFiles = _getVsFiles(relativePath);

            foreach (string curFile in allVsFiles)
            {
                FileInfo VSFileInfo = new FileInfo(curFile);

                if (VSFileInfo.Attributes == FileAttributes.Hidden)
                {
                    VSFileInfo.Attributes = FileAttributes.Normal;
                }
                string ITpipesFilePath = _replaceVsPath(curFile, relativePath);

                if (!File.Exists(ITpipesFilePath))
                {
                    UtilFunctions.copyFile(curFile, ITpipesFilePath, true);
                    continue;
                }

                FileInfo ITFileInfo = new FileInfo(ITpipesFilePath);

                if (VSFileInfo.LastAccessTime != null && ITFileInfo.LastAccessTime != null && VSFileInfo.LastAccessTime > ITFileInfo.LastAccessTime)
                {
                    copyFile(curFile, ITFileInfo.FullName, true);
                }
            }

            _disarmVirtualStore(relativePath);
        }

        private static string _replaceVsPath(string input, string directoryToRebaseTo)
        {
            string relativePath = input.Substring(input.IndexOf(@"AppData\Local\VirtualStore") + @"AppData\Local\VirtualStore".Length);
            return FixSlashes(relativePath);
        }

        private static List<string> _getVsFiles(string directoryPathRelativeToVolumeRoot)
        {
            List<string> returnList = new List<string>();

            string[] allUserFolders = Directory.GetDirectories(@"\Users\");

            foreach (string currentDirectory in allUserFolders)
            {
                string virtualStorePath = FixSlashes(Path.Combine(currentDirectory, $@"AppData\Local\VirtualStore\{directoryPathRelativeToVolumeRoot}"));
                try
                {
                    if (Directory.Exists(virtualStorePath))
                    {
                        returnList.AddRange(Directory.GetFiles(virtualStorePath, "*", SearchOption.AllDirectories));
                    }
                }
                catch (Exception ex)
                {
                }
            }
            return returnList;
        }

        private static void _disarmVirtualStore(string directoryPathRelativeToVolumeRoot)
        {
            string[] allUserFolders = Directory.GetDirectories(@"\Users\");

            foreach (string currentDirectory in allUserFolders)
            {
                string virtualStorePath = FixSlashes(Path.Combine(currentDirectory, $@"AppData\Local\VirtualStore\{directoryPathRelativeToVolumeRoot}"));
                try
                {
                    if (Directory.Exists(virtualStorePath))
                    {
                        Directory.Move(virtualStorePath, virtualStorePath + "_Obsoleted");
                    }
                }
                catch (Exception ex)
                {
                }
            }
        }
    }
}
