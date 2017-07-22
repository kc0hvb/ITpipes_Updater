using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ITpipes_Updater
{
    public class InstallationFailedException : Exception
    {
        public InstallationFailedException(string message) : base(message) { }
    }
}
