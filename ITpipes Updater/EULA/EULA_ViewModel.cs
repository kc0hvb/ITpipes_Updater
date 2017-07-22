using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Data;
using System.Windows.Input;

namespace ITpipes_Updater.EULA {
    class EULA_ViewModel : INotifyPropertyChanged {

        System.Windows.Controls.RichTextBox eulaRTB = new System.Windows.Controls.RichTextBox();

        EULA_Model curEULA_Model = null;
        
        public ICommand CancelEulaCommand {
            get {
                return new GenericCommandCanAlwaysExecute(cancelEula);
            }
        }


        private void cancelEula() {
            this.curEULA_Model.UserHasAcceptedEula = false;
        }

        private bool _eulaWasAccepted = false;


        public bool EulaWasAccepted {
            get {
                return _eulaWasAccepted;
            }
            set {
                _eulaWasAccepted = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("EulaWasAccepted"));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e) {
            if (PropertyChanged != null) {
                PropertyChanged(sender, e);
            }
        }

    }

    public class GenericCommandCanAlwaysExecute : ICommand {

        private Action _action;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter) {
            return true;
        }

        public GenericCommandCanAlwaysExecute(Action action) {

            _action = action;

        }

        public void Execute(object parameter) {
            _action();
        }
    }

}
