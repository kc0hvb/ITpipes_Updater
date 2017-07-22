using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;

namespace InspectIT
{
    public class DummyExe
    {
        public const string UpdateItDir = @"C:\UpdateIT";
        public const string ItpipesDir = @"C:\Program Files\InspectIT";

        static void Main(string[] args)
        {
            if (isItPipesRunning())
            {
                MessageBox.Show("ITpipes is already running");
                return;
            }

            if (!isItPipesLicensed())
            {
                string pathToCheckIt = getPathToCheckIT();
                if (File.Exists(pathToCheckIt))
                {
                    Process checkIt = new Process();
                    checkIt.StartInfo.FileName = pathToCheckIt;
                    MessageBox.Show("This installation of ITpipes is not licensed. Please create a license request file using the CheckIT application and contact ITpipes support to complete the licensing process.", "No License Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    checkIt.Start();

                    if (Path.GetFileName(pathToCheckIt).Equals("checkit_simple.exe", StringComparison.CurrentCultureIgnoreCase))
                    {
                        //CheckIT_Simple creates the ckf file in its application directory--need to show users where it is
                        checkIt.WaitForExit();
                        string pathToDir = Path.GetDirectoryName(pathToCheckIt);
                        Process openDirProc = new Process();
                        openDirProc.StartInfo.FileName = pathToDir;
                        openDirProc.Start();
                    }

                    return;
                }
            }
                Process itPipesProcess = new Process();
                itPipesProcess.StartInfo.FileName = Path.Combine(ItpipesDir, "ITpipes.exe");
                itPipesProcess.Start();
        }


        private static bool isItPipesRunning()
        {
            var myProcessId = Process.GetCurrentProcess().Id;
            Process[] itPipesProcesses = Process.GetProcessesByName("ITpipes");
            foreach (var proc in itPipesProcesses)
            {
                if (proc.Id != myProcessId)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool isItPipesLicensed()
        {
            string[] licenses = Directory.GetFiles(ItpipesDir, "*.lic", SearchOption.TopDirectoryOnly);
            return licenses.Length != 0;
        }

        private static string getPathToCheckIT()
        {
            //prefer the newest version:
            if (!Directory.Exists(UpdateItDir))
            {
                return null;
            }

            DirectoryInfo updateIt = new DirectoryInfo(UpdateItDir);
            foreach (var curDir in updateIt.GetDirectories().OrderByDescending(x => x.CreationTime))
            {
                foreach (var subDir in curDir.EnumerateDirectories())
                {
                    if (subDir.Name.Equals("checkit", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if(subDir.GetFiles().Any(x => x.Name.Equals("checkit_simple.exe", StringComparison.CurrentCultureIgnoreCase)))
                        {
                            //new simple checkit
                            return Path.Combine(subDir.FullName, "checkit_simple.exe");
                        }

                        return Path.Combine(subDir.FullName, "checkit_a.exe");
                    }
                }
            }

            return null;
        }
    }
}
