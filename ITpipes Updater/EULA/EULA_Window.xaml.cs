using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.ComponentModel;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Text.RegularExpressions;

namespace ITpipes_Updater.EULA {
    /// <summary>
    /// Interaction logic for EULA_Window.xaml
    /// </summary>
    public partial class EULA_Window : Window, INotifyPropertyChanged {

        #region Plaintext Eula

        public static readonly string EULA_PLAINTEXT = 
            "End-User: \nConsists of client and/or client employees.\n\nVendor: \nConsists of Infrastructure Technologies (I.T.), Authorized I.T. Dealer or Vendor.\n\nRepresentations: \nEnd-users shall limit their claims concerning the Software to the representations made by Vendor's published product literature.\n\nLimitation of liability for use of Software by End-Users: \nThe software is derived from third party, Vendor and no such third party warrants the software, or assumes any liability for any damages suffered or incurred by End-User including without limitation general, special or consequential damages arising from or in connection with the delivery, use of performance of software or this agreement.\n\nLimited warranty: \nVendor makes no warranty of any kind, expressed or implied, with regard to the equipment and/or software sold hereunder, except that Vendor warrants to End-User that the Equipment and Software, when installed and used in the United States of America or Canada shall be free from defects in materials or workmanship for a period of ninety (90) days from the date of original shipment. \n\nDuring the warranty period, Vendor or its designated service representative shall repair, or at its option, replace any such Equipment and/or Software which is returned to an authorized service center and that is confirmed by Vendor to be defective. If the Equipment and/or Software prove not to be defective or is not subject to warranty protection because of misuse, lapse of time or other reasons, end-Users shall be charged for repair or replacement at Vendor’s standard rates.\n\nThis warranty shall be null and void if the Equipment and/or Software has been damaged by accident, misuse or misapplication or has been modified or altered by End-User without Vendor’s express written acceptance of End-Users modifications for warranty purposes.\n\nSoftware License: \nThe Software license is non-transferable to other individuals or entities. The Software license cannot be sold or given to another person or entity. The Software can be purchased for another person or entity as a gift, but we must be informed in writing within 30 days of the original ship date of the Software from Vendor.\n\nWarranties; Limitations: \nThe foregoing warranties are in lieu of other warranties, expressed, or implied including without limitation, any warranty of merchantability or fitness for a particular purpose, or any warranty that equipment or software purchased hereunder is of a particular or merchantable equity.\n\nApplicable Law: \nAny and all warranties hereunder shall be constructed under the laws of the state of New Mexico. Any action for enforcement of warranties shall be proper only in the State of New Mexico. By its signature hereto, End-User agrees to the application of New Mexico law and venue for suit, only in Bernalillo County, New Mexico.\n\nTrademarks: \nNothing herein shall grant End-User any use, right, title or interest in the trade name Infrastructure Technologies or ITpipes or in other trademarks, service marks, words, symbols or other marks used, adopted or owned by Vendor from time to time, either alone or in association with other words or names. No rights are granted hereunder to use any trademark of Vendor in the name of the Software offered or furnished to the End-User.";

        #endregion Plaintext Eula

        public bool UserAcceptedEula = false;
        private Regex validPersonNameRegex = new Regex("[^a-zA-Z]");

        private string  _nameOfUserAcceptingEULA;
        
        public string  NameOfUserAcceptingEULA {
            get {
                return _nameOfUserAcceptingEULA;
            }
            set {
                _nameOfUserAcceptingEULA = value;

                if (isNameValid(value)) {
                    NameIsValid = true;
                }
                else {
                    NameIsValid = false;
                }

                OnPropertyChanged(this, new PropertyChangedEventArgs("NameOfUserAcceptingEULA"));
                
            }
        }

        private bool _nameIsValid;

        public bool NameIsValid {
            get {
                return _nameIsValid;
            }
            set {
                _nameIsValid = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("NameIsValid"));
            }
        }

        private bool _eulaWasAccepted;

        public bool EulaWasAccepted {
            get {
                return _eulaWasAccepted;
            }
            set {
                _eulaWasAccepted = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("EulaWasAccepted"));
            }
        }






        public EULA_Window() {

            this.DataContext = this;

            InitializeComponent();

            using (System.IO.Stream EulaStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ITpipes_Updater.Resources.EULA.rtf")) {

                curEulaRTB.Selection.Load(EulaStream, DataFormats.Rtf);
            }

            curEulaRTB.CaretPosition = curEulaRTB.CaretPosition.DocumentStart;
            FocusManager.SetFocusedElement(this, tboxSignature);
            Keyboard.Focus(tboxSignature);

        }

        private void butCancel_Click(object sender, RoutedEventArgs e) {
            UserAcceptedEula = false;
            this.Close();
        }




        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e) {
            if (PropertyChanged != null) {
                PropertyChanged(sender, e);
            }
        }

        private bool isNameValid(string input){

            string inputWithValidCharsOnly = validPersonNameRegex.Replace(input, "");

            if (inputWithValidCharsOnly.Length >= 5 && input.Trim().Contains(' ')) {
                return true;
            }

            return false;
        }

        private void butAcceptEULA_Click(object sender, RoutedEventArgs e) {
            if (NameIsValid) {
                EulaWasAccepted = true;
            }
            this.Close();
        }

        private void butCancel_Click_1(object sender, RoutedEventArgs e) {
            this.EulaWasAccepted = false;
            this.Close();
        }

        private void tboxSignature_PreviewKeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.Enter && isNameValid(this.NameOfUserAcceptingEULA)) {
                e.Handled = true;
                butAcceptEULA_Click(sender, new RoutedEventArgs());
            }
        }
    }
}
