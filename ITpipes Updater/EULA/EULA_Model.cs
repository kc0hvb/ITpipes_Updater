using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.ComponentModel;
using System.Windows.Documents;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ITpipes_Updater.EULA {
    class EULA_Model : INotifyPropertyChanged {

        private FlowDocument _eulaFlowDoc = new FlowDocument();

        private Regex validPersonNameRegex = new Regex("[^a-zA-Z]");

        public FlowDocument EulaFlowDoc {
            get {
                return _eulaFlowDoc;
            }
            set {
                _eulaFlowDoc = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("EulaFlowDoc"));
            }
        }

        private string _nameOfPersonAcceptingEula;

        public string NameOfPersonAcceptingEula {
            get {
                return _nameOfPersonAcceptingEula;
            }
            set {
                _nameOfPersonAcceptingEula = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("NameOfPersonAcceptingEula"));

                if (isPersonNameValid(value)) {
                    PersonNameIsValid = true;
                }
                else {
                    PersonNameIsValid = false;
                }
            }
        }

        private bool _personNameIsValid;

        public bool PersonNameIsValid {
            get {
                return _personNameIsValid;
            }
            set {
                _personNameIsValid = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("PersonNameIsValid"));
            }
        }


        private bool _userHasAcceptedEula;

        public bool UserHasAcceptedEula {
            get {
                return _userHasAcceptedEula;
            }
            set {
                _userHasAcceptedEula = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("UserHasAcceptedEula"));
            }
        }



        public EULA_Model() {
            loadEulaText();
        }

        private bool isPersonNameValid(string input) {

            string alphabeticalOnlyInput = validPersonNameRegex.Replace(input, "");

            if (alphabeticalOnlyInput.Length >= 5 && input.Trim().Contains(' ')) {
                return true;
            }

            return false;
        }

        private void loadEulaText() {

            TextRange newEulaTextRange = new TextRange(EulaFlowDoc.ContentStart, EulaFlowDoc.ContentEnd);

            using (Stream logFileStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ITpipes_Updater.Resources.EULA.rtf")) {

                newEulaTextRange.Load(logFileStream, DataFormats.Rtf);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e) {

            if (PropertyChanged != null) {
                PropertyChanged(sender, e);
            }

        }
    }

}
