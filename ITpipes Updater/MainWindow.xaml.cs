//#define DEBUG_MESSAGEBOXES_ENABLED
//#define DEBUG_DISABLE_INTERNAL_RESOURCE_LOADING
//#define DEBUG_DISABLE_ERROR_STOPS

using ITpipes_Updater.Util;
using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Navigation;

namespace ITpipes_Updater
{
    public partial class MainWindow : Window
    {
        public string buttonToEngageUpdateOrInstallDefaultText
        {
            get { return (string)GetValue(buttonToEngageUpdateOrInstallDefaultTextProperty); }
            set { SetValue(buttonToEngageUpdateOrInstallDefaultTextProperty, value); }
        }
        public static readonly DependencyProperty buttonToEngageUpdateOrInstallDefaultTextProperty =
            DependencyProperty.Register("buttonToEngageUpdateOrInstallDefaultText", typeof(string), typeof(MainWindow), new PropertyMetadata("Update ITpipes"));

        private static bool runConfigAfterUpdate = false, updaterWasRun = false;

        public static MainWindow curProgramWindow;
        public static string
            updaterLaunchDirectory,
            versionNumString;

        BackgroundWorker bwUpdaterWorker;

        //made static so the background worker can view these settings lazily
        public static bool installQSBCounter = false,
                           overwriteTemplatesWithNewest = false,
                           forceInstallerMode = false,
                           installMpeg4 = false;


        public Visibility displayForceInstallerModeCheckbox { get; set; } = Visibility.Visible;

        public MainWindow()
        {
            //todo: replace all of the everything with an MVVM-based model.

            if (isDotNet35Installed() == false)
            {
                MessageBox.Show("Microsoft .Net 3.5 is required by ITpipes. You will be directed to the Microsoft download page for .Net 3.5", ".Net 3.5 Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                launchDotNet35DownloadPage();
                this.Close();
                return;
            }

            DataContext = this;

            if (doesProcessHaveAdminAccess() == false)
            {
                MessageBox.Show("This program can only be run as a local, domain, or UAC elevated administration" + newLine() +
                                "UAC is disabled: You must log in to Windows as an Administrator to run the updater.", "Insufficient Permissions",
                                MessageBoxButton.OK,
                                MessageBoxImage.Exclamation);
                this.Close();
            }

            ChangeLog curChangeLog = new ChangeLog();

            if (isITpipesAlreadyInstalled() == false)
            {
                forceInstallerMode = true;
                buttonToEngageUpdateOrInstallDefaultText = "Install ITpipes";
                displayForceInstallerModeCheckbox = Visibility.Collapsed;
            }

            InitializeComponent();
            curProgramWindow = this;
            updaterLaunchDirectory = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);

            versionNumString = getItpipesVersionNumber();

            lblVersionNumber.Content = versionNumString;

            rtfboxChangeLog.AddHandler(Hyperlink.RequestNavigateEvent, new RequestNavigateEventHandler(HandleRequestNavigate));

            setChangeLogText();
        }

        private bool isITpipesAlreadyInstalled()
        {

            using (RegistryKey installLocationRegKey32 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\ITpipes", false))
            using (RegistryKey installLocationRegKey64 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node\ITpipes", false))
            {

                if (installLocationRegKey32 == null && installLocationRegKey64 == null)
                {
                    return false;
                }
            }

            return true;
        }

        private void HandleRequestNavigate(object sender, RequestNavigateEventArgs args)
        {
            //The requestnavigateevent is being handled in a new thread because it was taking a long time for the link to be launched.
            //Now that that happens in a separate thread, it goes very quickly! Woohoo!
            System.Threading.Thread t = new System.Threading.Thread(startHyperlink);
            t.Start(args.Uri.ToString());
        }

        private void startHyperlink(object linkText)
        {
            Process.Start((string)linkText);
        }

        private void setChangeLogText()
        {
            Stream logFileStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ITpipes_Updater.Resources.changes.rtf");
            rtfboxChangeLog.Selection.Load(logFileStream, DataFormats.Rtf);
            rtfboxChangeLog.CaretPosition = rtfboxChangeLog.Selection.Start; //to prevent the entire paragraph from flashing selected when the user clicks in the richtextbox
        }
        
        private void disableEngageButton()
        {
            this.butEngage.IsEnabled = false;
        }

        private void enableEngageButton()
        {
            this.butEngage.IsEnabled = true;
        }

        private bool isITpipesCurrentlyRunning()
        {
            bool itPipesIsRunning = true,
                 itPipesIsNotRunning = false;

            Process[] oldExecutablesRunning = Process.GetProcessesByName("inspectIT");

            if (oldExecutablesRunning.Length != 0)
            {
                return itPipesIsRunning;
            }

            Process[] itPipesExecutablesRunning = Process.GetProcessesByName("ITpipes");

            if (itPipesExecutablesRunning.Length != 0)
            {
                return itPipesIsRunning;
            }

            return itPipesIsNotRunning;
        }

        private bool isTemplateEditorRunning()
        {
            Process[] teProc = Process.GetProcessesByName("TemplateEditor");
            return teProc.Length != 0;
        }

        private bool isManageItRunning()
        {
            Process[] miProc = Process.GetProcessesByName("ManageIT");
            return miProc.Length != 0;
        }

        private void writeEulaAcceptanceRegKeys(string nameOfUserAcceptingEula, string FullEulaText)
        {
            string newKeyName = DateTime.UtcNow.ToFileTimeUtc().ToString();
            string EulaAcceptanceDate = DateTime.UtcNow.ToShortDateString() + " " + DateTime.UtcNow.ToShortTimeString();

            using (RegistryKey EulaRegKey = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\ITP_Lic", RegistryKeyPermissionCheck.ReadWriteSubTree))
            using (RegistryKey newEulaKey = EulaRegKey.CreateSubKey(newKeyName))
            {
                newEulaKey.SetValue("EULA_Accepted_By", nameOfUserAcceptingEula, RegistryValueKind.String);
                newEulaKey.SetValue("Full_EULA_Text", FullEulaText, RegistryValueKind.String);
                newEulaKey.SetValue("DateTime_EULA_Accepted_UTC", EulaAcceptanceDate, RegistryValueKind.String);
            }
        }
        
        private void closeUpdater(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void BwUpdaterWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var installConfig = new InstallationRequest(InstallationLogic.DEFAULT_INSTALL_DIRECTORY, overwriteTemplatesWithNewest, forceInstallerMode, installQSBCounter, versionNumString, installMpeg4);
            var installationLogic = new InstallationLogic(installConfig);
            installationLogic.ProgressUpdateAvailable += (x, y) =>
            {
                setUpdateStatusUsingUIThread(y.StatusMessage);
            };

            try
            {
                if (!isITpipesAlreadyInstalled())
                {
                    EULA.EULA_Window curEulaWindow = null;

                    curProgramWindow.Dispatcher.Invoke(new Action(() =>
                    {
                        curEulaWindow = new EULA.EULA_Window();

                        curEulaWindow.ShowDialog();
                    }));


                    if (curEulaWindow.EulaWasAccepted)
                    {
                        writeEulaAcceptanceRegKeys(curEulaWindow.NameOfUserAcceptingEULA, EULA.EULA_Window.EULA_PLAINTEXT);
                    }
                    else
                    {
                        setUpdateStatusUsingUIThread("End User License Agreement Not Accepted -- Click Here to Close Installer");
                        curProgramWindow.Dispatcher.Invoke(new Action(() =>
                        {
                            butEngage.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(0xff, 0xC9, 0x11, 0x11));
                            butEngage.IsEnabled = true;
                            butEngage.Click -= butEngage_Click;
                            butEngage.Click += closeUpdater;
                        }));

                        return;
                    }
                }

                if (forceInstallerMode)
                {
                    installConfig.BackupToRestore = getBackupFileToUse();
                    installationLogic.RunAsInstaller();

                    if (installationLogic.bErrorExists == false)
                    {
                        setUpdateStatusUsingUIThread("ITpipes Installation Completed!");
                    }
                    else
                    {
                        setUpdateStatusUsingUIThread("ITPipes Installation Completed. Errors did occur. Please see logs to errors.");
                    }
                }
                else
                {
                    installationLogic.RunAsUpdater();

                    if (installationLogic.bErrorExists == false)
                    {
                        setUpdateStatusUsingUIThread("ITpipes Update Completed!");
                    }
                    else
                    {
                        setUpdateStatusUsingUIThread("ITPipes Update Complete. Errors did occur. Please see log for errors.");
                    }
                }

                curProgramWindow.Dispatcher.Invoke(new Action(() =>
                {
                    butEngage.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(0xFF, 0x36, 0xE0, 0x3A));
                    butEngage.IsEnabled = true;
                    butEngage.Click -= butEngage_Click;
                    butEngage.Click += closeUpdater;
                }));

                if (MainWindow.runConfigAfterUpdate)
                {
                    string configPath = Path.Combine(updaterLaunchDirectory, @"Config\ITpipes Config.exe");

                    if (File.Exists(configPath))
                    {
                        Process configProcess = new Process();
                        configProcess.StartInfo.FileName = configPath;

                        configProcess.Start();
                    }
                }
            }
            catch (Exception ex)
            {
                curProgramWindow.Dispatcher.Invoke(new Action(() =>
                {

                    butEngage.Content += "Error: Installation cancelled";
                    butEngage.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(0xFF, 0xF3, 0x08, 0x08)); //#FFF30808
                    butEngage.IsEnabled = true;
                    butEngage.Click -= butEngage_Click;
                    butEngage.Click += closeUpdater;
                    MessageBox.Show($"Error installing ITpipes: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
                }
                ));
            }
        }

        private string getBackupFileToUse()
        {
            string itpBackupsDir = UtilFunctions.getBackupDirectory();
            string[] availableBackupFiles = Directory.GetFiles(updaterLaunchDirectory, "*." + UtilFunctions.getBackupFileExtension(), SearchOption.TopDirectoryOnly);

            if (availableBackupFiles.Length > 0)
            {
                MessageBoxResult resRestoreLatestBackup = MessageBox.Show(
                    string.Format("An ITpipes configuration backup file has been included with this installer.{0}{0}Would you like to restore this backup during installation?", Environment.NewLine),
                    "Backup included with installer",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (resRestoreLatestBackup == MessageBoxResult.Yes)
                {
                    return availableBackupFiles.OrderByDescending(x => new FileInfo(x).CreationTime).First();
                }
            }

            if (Directory.Exists(itpBackupsDir))
            {
                availableBackupFiles = Directory.GetFiles(itpBackupsDir, string.Format("*.{0}", UtilFunctions.getBackupFileExtension()));

                if (availableBackupFiles.Length > 0)
                {
                    FileInfo latestBackupFI = (from filePath in availableBackupFiles select new FileInfo(filePath)).OrderByDescending(x => x.CreationTime).First();

                    MessageBoxResult resRestoreLatestBackup =
                    MessageBox.Show(
                        string.Format("One or more configuration backups exist from a previous installation of ITpipes.{0}{0}" +
                            "Would you like to restore the most recent configuration backup?{0}(Backup Date: {1} - {2})",
                            Environment.NewLine,
                            latestBackupFI.CreationTime.ToShortDateString(),
                            latestBackupFI.CreationTime.ToShortTimeString()),
                        "ITpipes Backups Found",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (resRestoreLatestBackup == MessageBoxResult.Yes)
                    {
                        return latestBackupFI.FullName;
                    }
                }
            }

            return null;
        }
        
        private void setUpdateStatusUsingUIThread(string statusToSet)
        {
            curProgramWindow.Dispatcher.Invoke(new Action(() =>
            {
                butEngage.Content = statusToSet;
            }
            ));
        }

        private void butEngage_Click(object sender, RoutedEventArgs e)
        {
            if (isITpipesCurrentlyRunning())
            {
                MessageBox.Show("ITpipes is currently running. Please close ITpipes before beginning the update.", "Update Cannot Start", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (isTemplateEditorRunning())
            {
                MessageBox.Show("Template Editor is currently running. Please close Template Editor before beginning the update.", "Updated Cannot Start", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (isManageItRunning())
            {
                MessageBox.Show("ManageIT is currently running. Please close ManageIT before beginning the update.", "Updated Cannot Start", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (overwriteTemplatesWithNewest == false || (overwriteTemplatesWithNewest == true && verifyUserWantsToOverwriteTemplates()))
            {
                if (bwUpdaterWorker != null && bwUpdaterWorker.IsBusy == true)
                {
                    //Should not be possible, since the engage button should be disabled, but just to be safe:
                    return;
                }

#if DEBUG_DISABLE_INTERNAL_RESOURCE_LOADING
#else
                disableEngageButton();
#endif

                if (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift))
                {
                    runConfigAfterUpdate = true;
                }

                bwUpdaterWorker = new BackgroundWorker();
                bwUpdaterWorker.DoWork += BwUpdaterWorker_DoWork;
                bwUpdaterWorker.RunWorkerAsync();
                bwUpdaterWorker.RunWorkerCompleted += (x, y) => { updaterWasRun = true; };
            }
        }

        private string newLine(int count = 1)
        {
            string returnString = "";
            for (int i = 0; i < count; i++)
            {
                returnString += Environment.NewLine;
            }
            return returnString;
        }

        private bool verifyUserWantsToOverwriteTemplates()
        {
            var result = MessageBox.Show("Are you sure that you wish to overwrite your templates with the latest default templates?" + newLine(2) +
                                         "All of your templates will be backed up, but if you have a custom template for inspections this may overwrite the template you use." + newLine(2),
                                         "Replace Templates", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                return true;
            }

            return false;
        }

        private void cboxReplaceTemplates_Click(object sender, RoutedEventArgs e)
        {
            overwriteTemplatesWithNewest = (bool)cboxReplaceTemplates.IsChecked;
        }

        private void cboxInstallQSB_Click(object sender, RoutedEventArgs e)
        {
            installQSBCounter = (bool)cboxInstallQSB.IsChecked;
        }

        private void cboxMpeg4_Click(object sender, RoutedEventArgs e)
        {
            installMpeg4 = (bool)cboxMpeg4.IsChecked;
        }

        private void chkBxForceReinstallation_Click(object sender, RoutedEventArgs e)
        {
            forceInstallerMode = (bool)chkBxForceReinstallation.IsChecked;

            if (forceInstallerMode == true)
            {
                buttonToEngageUpdateOrInstallDefaultText = "Re-Install ITpipes";
            }
            else
            {
                buttonToEngageUpdateOrInstallDefaultText = "Update ITpipes";
            }
        }

        private void HandleRequestNavigate(object sender, MouseButtonEventArgs e)
        {
            System.Threading.Thread t = new System.Threading.Thread(startHyperlink);
            t.Start(@"http://www.itpipes.com");
        }

        private string getItpipesVersionNumber()
        {
            try
            {
                return "v" + FileVersionInfo.GetVersionInfo(updaterLaunchDirectory + @"\FILE\ITpipes.exe")
                             .FileVersion
                             .Replace(',', '.')
                             .Replace(".00", ".0").Replace(".0.0", ".0.").Replace("..", ".");
            }
            catch
            {
                return "UNKNOWN";
            }
        }


        private bool doesProcessHaveAdminAccess()
        {
            //Ran into issues testing permissions on systems with UAC disabled. The only consistent test I could find was just to verify that I have write access to
            //protected folders.

            //If non-administrator users have write permission to Program Files and Windows\System32 directories, then these people are beyond our help.

            //TODO: Add a regsvr32 attempt to verify that files are registerable.

            bool returnValue = false;

            string testFilePath = @"C:\Program Files\test.txt";
            try
            {
                File.Create(testFilePath).Close();
                File.Delete(testFilePath);

                testFilePath = @"C:\Windows\System32\test.txt";
                File.Create(testFilePath).Close();
                File.Delete(testFilePath);

                returnValue = true;
            }
            catch
            {
                returnValue = false;
            }

            return returnValue;
        }
        
        private bool isDotNet35Installed()
        {
            using (RegistryKey DotNet35RegKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\.NETFramework\AssemblyFolders\v3.5"))
            {
                if (DotNet35RegKey == null)
                {
                    return false;
                }
            }

            return true;
        }

        private void launchDotNet35DownloadPage()
        {

            System.Threading.Thread t = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(
                x =>
                {
                    Process.Start(@"http://www.microsoft.com/en-us/download/details.aspx?id=25150");
                }));

            t.Start();
        }
    }

    public class UpdaterEventArgs : EventArgs
    {
        string ProgressText;

        public UpdaterEventArgs(string newProgressText)
        {
            ProgressText = newProgressText;
        }
    }

    public class ChangeLog : INotifyPropertyChanged
    {
        string internalChangeLog = "";
        public string changeLogText
        {
            get { return internalChangeLog; }
            set
            {
                this.internalChangeLog = value;
                OnPropertyChanged("changeLogText");
            }
        }

        public ChangeLog() { }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string info)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(info));
        }
    }
}
