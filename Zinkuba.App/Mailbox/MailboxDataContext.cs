using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Zinkuba.MailModule;
using Zinkuba.MailModule.API;

namespace Zinkuba.App.Mailbox
{
    public class MailboxDataContext : INotifyPropertyChanged
    {

        public event PropertyChangedEventHandler PropertyChanged;

        public virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
        private int _progress;
        
        public String ProgressText
        {
            get { return Exporter == null ? "Ready" : (Exporter.State == MessageProcessorStatus.Started ? "" + _progress + "%" : Exporter.State.ToString()); }
        }

        public int IgnoredMails { get; set; }
        public int FailedMails { get; set; }
        public int ExportedMails { get; set; }

        public int Progress
        {
            get { return _progress; }
            set
            {
                if (_progress != value)
                {
                    _progress = value;
                    OnPropertyChanged("Progress");
                    OnPropertyChanged("ProgressText");
                }
                OnPropertyChanged("ExportedMails");
                OnPropertyChanged("FailedMails");
                OnPropertyChanged("IgnoredMails");
            }
        }

        public MessagePipeline Exporter;
    }
}
