using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Reflection;

namespace ITpipes_Updater
{
    public class RegistrationHelper
    {
        public static readonly string REGASM = @"C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"; //this never changes location
        private string baseForceReRegDirectory = Directory.Exists(@"C:\Windows\syswow64") ? @"C:\Windows\syswow64\" : @"C:\Windows\System32\",
                        systemFolder = Environment.GetFolderPath(Environment.SpecialFolder.SystemX86);

        //The doNotRegManifest is created by attempting to register every file as COM, moving all failures to RegAsm, then moving all of the failures for that to the do not reg list.
        private string[] comRegManifest,
                         asmRegManifest,
                         doNotRegManifest;

        public static readonly string[] REGISTERABLE_FILE_EXTENSIONS = { ".dll", ".ocx", ".ax" };

        Process 
            regsvr32Process, 
            Regasm;

        public RegistrationHelper()
        {
            populateRegManifests();
            Regasm = new Process();
            Regasm.StartInfo.FileName = REGASM;
            Regasm.StartInfo.CreateNoWindow = true;
            Regasm.StartInfo.RedirectStandardOutput = true;
            Regasm.StartInfo.UseShellExecute = false;
            
            regsvr32Process = new Process();
            regsvr32Process.StartInfo.FileName = systemFolder + @"\regsvr32.exe";
            regsvr32Process.StartInfo.CreateNoWindow = true;
            regsvr32Process.StartInfo.RedirectStandardOutput = true;
            regsvr32Process.StartInfo.UseShellExecute = false;
        }

        public bool IsFileRegisterableType(string filename)
        {
            return REGISTERABLE_FILE_EXTENSIONS.Contains(Path.GetExtension(filename), StringComparer.CurrentCultureIgnoreCase);
        }

        public void RegisterFile(string pathToFile, bool unregisterFirst = false)
        {
            if (doNotRegManifest.Contains(Path.GetFileName(pathToFile), StringComparer.InvariantCultureIgnoreCase))
            {
                return;
            }

            else if (comRegManifest.Contains(Path.GetFileName(pathToFile), StringComparer.InvariantCultureIgnoreCase))
            {
                RegisterAsCom(pathToFile, unregisterFirst);
            }

            else if (asmRegManifest.Contains(Path.GetFileName(pathToFile), StringComparer.InvariantCultureIgnoreCase))
            {
                registerAsAssembly(pathToFile, unregisterFirst);
            }
            else
            {
                RegisterAsCom(pathToFile, unregisterFirst);
                registerAsAssembly(pathToFile, unregisterFirst);
            }
        }

        public void RegisterAsCom(string filePathToReg, bool unregisterFirst = false)
        {
            if (unregisterFirst)
            {
                regsvr32Process.StartInfo.Arguments = $"/s /u \"{filePathToReg}\"";
                regsvr32Process.Start();
                regsvr32Process.WaitForExit();
            }

            regsvr32Process.StartInfo.Arguments = "/s \"" + filePathToReg + "\"";
            regsvr32Process.Start();
            regsvr32Process.WaitForExit();
            int regSvr32ExitCode = regsvr32Process.ExitCode;
            if (regSvr32ExitCode != 0)
            {
                Console.WriteLine(string.Format("Registration failed for file: {0}\nResult: {1}", filePathToReg, regSvr32ExitCode.ToString()));
            }

            regsvr32Process.Close();
        }

        private void registerAsAssembly(string filePathToReg, bool unregisterFirst = false)
        {
            if (unregisterFirst)
            {
                Regasm.StartInfo.Arguments = $"/silent /unregister \"{filePathToReg}\"";
                Regasm.Start();
                Regasm.WaitForExit();
            }

            string tlbName = Path.GetFileNameWithoutExtension(filePathToReg) + ".tlb";

            try
            {
                Regasm.StartInfo.Arguments = $"/silent /codebase \"{filePathToReg}\" /tlb:{tlbName}";
                Regasm.Start();
                Regasm.WaitForExit();

                if (Regasm.ExitCode != 0)
                {
                    Regasm.StartInfo.Arguments = $"/silent \"{filePathToReg}\"";
                    Regasm.Start();
                    Regasm.WaitForExit();
                    if (Regasm.ExitCode != 0)
                    {
                        throw new Exception(Regasm.StandardError?.ReadToEnd());
                    }
                }
            }
            catch
            {
                //not worried--if it didn't register it's because it's not a registerable assembly
            }
        }

        private void populateRegManifests()
        {
            string regManifestString = "";
            using (Stream regManifestStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ITpipes_Updater.Resources.RegManifest"))
            {
                using (StreamReader regManifestReader = new StreamReader(regManifestStream))
                {
                    regManifestString = regManifestReader.ReadToEnd();
                }
            }

            //remove the carriage returns--only interested in the newlines:
            regManifestString = regManifestString.Replace("\r", "");

            int beginComRegParseIndex = regManifestString.IndexOf("COM Registration:"),
                endComRegParseIndex = regManifestString.IndexOf("Assembly Registration:");

            //Set start index to the first character of the new line following the text "COM Registration:"
            beginComRegParseIndex = regManifestString.IndexOf('\n', beginComRegParseIndex) + 1;

            comRegManifest = regManifestString.Substring(beginComRegParseIndex, endComRegParseIndex - beginComRegParseIndex)
                             .Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

            int beginAsmParseIndex = regManifestString.IndexOf('\n', endComRegParseIndex) + 1,
                endAsmParseIndex = regManifestString.IndexOf("Do Not Register:");

            asmRegManifest = regManifestString.Substring(beginAsmParseIndex, endAsmParseIndex - beginAsmParseIndex)
                             .Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

            int beginDoNotRegIndex = regManifestString.IndexOf('\n', endAsmParseIndex) + 1,
                endDoNotRegIndex = regManifestString.Length;

            doNotRegManifest = regManifestString.Substring(beginDoNotRegIndex, endDoNotRegIndex - beginDoNotRegIndex)
                             .Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
        }
    }
}
