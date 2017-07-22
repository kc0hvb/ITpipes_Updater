using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Data;
using System.Windows;
using System.ComponentModel;

namespace ITpipes_Updater.Object_Models {

    class Installer_Model : INotifyPropertyChanged {


        private InstallerMode _programMode;

        public InstallerMode ProgramMode {
            get {
                return _programMode;
            }
            set {
                _programMode = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("ProgramMode"));
            }
        }

        private string _itpipesDirectory = @"C:\Program Files\InspectIT";

        public string ITpipesDirectory {
            get {
                return _itpipesDirectory;
            }
            set {
                _itpipesDirectory = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("ITpipesDirectory"));
            }
        }

        private string _updaterStorageDirectory = @"C:\UpdateIT";

        public string UpdaterStorageDirectory {
            get {
                return _updaterStorageDirectory;
            }
            set {
                _updaterStorageDirectory = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("UpdaterStorageDirectory"));
            }
        }

        private string _appLaunchDir;

        public string AppLaunchDir {
            get {
                return _appLaunchDir;
            }
            set {
                _appLaunchDir = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("AppLaunchDir"));
            }
        }

        private BackgroundWorker totalUpdaterDirSizeCalcBW = new BackgroundWorker();

        private long _totalUpdaterDirSize;

        public long TotalUpdaterDirSize {
            get {
                return _totalUpdaterDirSize;
            }
            set {
                _totalUpdaterDirSize = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("TotalUpdaterDirSize"));
            }
        }

        private bool _isITpipesAlreadyInstalled;

        public bool IsITpipesAlreadyInstalled {
            get {
                return _isITpipesAlreadyInstalled;
            }
            set {
                _isITpipesAlreadyInstalled = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("IsITpipesAlreadyInstalled"));
            }
        }

        


        private string[] _itpipesUserFiles;

        public string[] ITpipesUserFiles {
            get {
                 return _itpipesUserFiles;
            }
            set {
                _itpipesUserFiles = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("ITpipesUserFiles"));
            }
        }










        public Installer_Model(string itpipesDir = @"C:Program Files\InspectIT") {

            if (Directory.Exists(itpipesDir)) {

            }

        }



        public event PropertyChangedEventHandler PropertyChanged;


        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e) {
            if (PropertyChanged != null) {
                PropertyChanged(sender, e);
            }
        }


    }

    



    public enum InstallerMode {
        Installer,
        Updater,
        Repair
    }


}
