using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Zinkuba.App.Folder
{
    public class MboxFolder
    {
        public String MboxPath
        {
            get { return _mboxPath; }
            set { _mboxPath = value;
                if (String.IsNullOrEmpty(FolderPath))
                {
                    FolderPath = Regex.Match(value, @".*\\(.*)\.mbox").Groups[1].Value;
                }
            }
        }

        public String FolderPath;
        private string _mboxPath;
    }
}
