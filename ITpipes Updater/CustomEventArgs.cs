using System;

namespace ITpipes_Updater
{
    public class ProgressReportEventArgs : EventArgs
    {
        public string StatusMessage { get; set; }
        public ProgressReportEventArgs(string message)
        {
            StatusMessage = message;
        }
    }
}
