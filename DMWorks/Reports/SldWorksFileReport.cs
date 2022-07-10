using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace DMWorks
{
    public class SldWorksFileReport: INotifyPropertyChanged {

        private string fileName;
        private string savedBy;
        private DateTime creationDate;
        private DateTime lastSaveDate;
        private bool dateViolation;
        private string warningNote;
        private string creationMatchId;
        private string saveMatchId;
        private string creationSaveMatchId;

        public SldWorksFileReport(){

        }

        public string FileName { get { return fileName; } set { fileName = value; } }
        public string SavedBy { get { return savedBy; } set { savedBy = value; } }
        public DateTime CreationDate { get { return creationDate; } set { creationDate = value; } }
        public DateTime LastSaveDate { get { return lastSaveDate; } set { lastSaveDate = value; } }
        public Boolean DateViolation { get { return dateViolation; } set { dateViolation = value; } }
        public string WarningNote { get { return warningNote; } set { warningNote = value; } }
        public string CreationMatchId { get { return creationMatchId; } set { creationMatchId = value; NotifyPropertyChanged(); } }
        public string SaveMatchId { get { return saveMatchId; } set { saveMatchId = value; NotifyPropertyChanged(); } }
        public string CreationSaveMatchId { get { return creationSaveMatchId; } set { creationSaveMatchId = value; NotifyPropertyChanged(); } }



        //Ref: https://docs.microsoft.com/en-us/dotnet/framework/winforms/how-to-implement-the-inotifypropertychanged-interface
        public event PropertyChangedEventHandler PropertyChanged;
        private void NotifyPropertyChanged([CallerMemberName] String propertyName = ""){
            if (PropertyChanged != null){
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
