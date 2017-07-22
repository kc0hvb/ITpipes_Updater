//#define DEBUG_UNINSTALLER_TESTING

using System.Windows;
using System.IO;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ITpipes_Uninstaller.Uninstaller;

namespace ITpipes_Uninstaller
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
#if DEBUG_UNINSTALLER_TESTING

            InitializeComponent();

#else
            if (doesProcessHaveAdminAccess()) {

                MessageBoxResult userInput = MessageBox.Show("Uninstall ITpipes and all of its components?", "Uninstall ITpipes", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (userInput == MessageBoxResult.Yes) {

                    bool uninstallResult = false;

                    try {
                        uninstallResult = Uninstaller.Uninstaller.UninstallItpipes();
                    }
                    catch (ProgramNotInstalledException ex) {
                        MessageBox.Show("Uninstallation canceled: ITpipes is not installed.");
                        this.Close();
                        return;
                    }
                    catch (System.Security.AccessControl.PrivilegeNotHeldException ex) { //this is only thrown if forceUninstall flag in the UninstallITpipes static function is false.
                        MessageBox.Show("Uninstallation canceled: Uninstallation cannot be run without elevated administrator priviledges. Please uninstall through Control Panel.");
                        this.Close();
                        return;
                    }
                    catch (DirectoryNotFoundException ex) { //this is only thrown if forceUninstall flag in the UninstallITpipes static function is false.
                        MessageBox.Show("Uninstallation canceled: ITpipes installation directory does not exist");
                        this.Close();
                        return;
                    }
                    catch (System.Exception ex) {
                        MessageBox.Show(string.Format("Uninstallation Failed:\n\nException encountered:\n{0}", ex));
                    }

                    if (uninstallResult == true) {
                        MessageBox.Show("ITpipes successfully uninstalled");
                        this.Close();
                        return;
                    }
                    else {
                        MessageBox.Show("ITpipes uninstallation failed");
                    }
                }
                else {
                    this.Close();
                    return;
                }
            }
            
            else {
                MessageBox.Show(
                    "The uninstaller may only be run with an elevated Administrator account.\n\nPlease uninstall ITpipes through Control Panel.",
                    "Cannot Uninstall Without Administrator Account",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                this.Close();
            }
#endif
        }

        private void butUninstall_Click(object sender, RoutedEventArgs e) {
            if (doesProcessHaveAdminAccess()) {
                
                if (Uninstaller.Uninstaller.UninstallItpipes()) {
                    MessageBox.Show("ITpipes successfully uninstalled");
                    this.Close();
                }
                else {
                    MessageBox.Show("ITpipes uninstallation failed");
                    this.Close();
                }
            }
            else {
                MessageBox.Show(
                    "The uninstaller may only be run with an elevated Administrator account.\n\nPlease uninstall ITpipes through Control Panel.", 
                    "Cannot Uninstall Without Administrator Account", 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Warning);
                this.Close();
            }
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
}
