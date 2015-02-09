using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Controls;
using log4net;
using Zinkuba.App.Annotations;

namespace Zinkuba.App
{
        public partial class AuthenticatedMailboxTargetControl : UserControl
        {
            private static readonly ILog Logger = LogManager.GetLogger(typeof(AuthenticatedMailboxControl));
            public AuthenticatedMailboxTargetData AuthenticatedMailboxData;

            public AuthenticatedMailboxTargetControl(string id)
            {
                InitializeComponent();
                AuthenticatedMailboxData = new AuthenticatedMailboxTargetData(id);
                DataContext = AuthenticatedMailboxData;
            }

            public bool Validate()
            {
                if (Dispatcher.CheckAccess())
                {
                    return !String.IsNullOrWhiteSpace(UsernameField.Text) &&
                           !String.IsNullOrWhiteSpace(PasswordField.Password);
                }
                else
                {
                    bool result = false;
                    Dispatcher.Invoke(new Action(() => result = Validate()));
                    return result;
                }
            }

            private void PasswordField_LostFocus(object sender, System.Windows.RoutedEventArgs e)
            {
                AuthenticatedMailboxData.Password = PasswordField.Password;
            }
        }


    public class AuthenticatedMailboxTargetData : INotifyPropertyChanged
    {
        private string _password;
        private string _username;
        private string _id;

        public AuthenticatedMailboxTargetData(string id)
        {
            Id = id;
        }

        public String Id
        {
            get { return _id; }
            set
            {
                if (value == _id) return;
                _id = value;
                OnPropertyChanged();
            }
        }

        public string Username
        {
            get { return _username; }
            set
            {
                if (value == _username) return;
                _username = value;
                OnPropertyChanged();
            }
        }

        public String Password
        {
            get { return _password; }
            set
            {
                if (value == _password) return;
                _password = value;
                OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
