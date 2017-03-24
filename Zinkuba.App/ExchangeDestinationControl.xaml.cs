using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Zinkuba.App.Mailbox;
using Zinkuba.MailModule.API;
using Zinkuba.MailModule.MessageProcessor;

namespace Zinkuba.App
{
    /// <summary>
    /// Interaction logic for ExchangeDestinationControl.xaml
    /// </summary>
    public partial class ExchangeDestinationControl : UserControl, IDestinationManager
    {
        private readonly MainWindow _mainWindow;
        private Dictionary<string, AuthenticatedMailboxTargetControl> _destinations;
        private readonly object _destinationLock = new object();
        private ExchangeDestinationDataContext dataContext;

        public ExchangeDestinationControl(MainWindow mainWindow)
        {
            _mainWindow = mainWindow;
            InitializeComponent();
            _destinations = new Dictionary<String, AuthenticatedMailboxTargetControl>();
            dataContext = new ExchangeDestinationDataContext();
            DataContext = dataContext;
        }

        public IMessageDestination GetDestination(String id)
        {
            lock (_destinationLock)
            {
                if (_destinations.ContainsKey(id))
                {
                    var target = new ExchangeTarget(dataContext.Server,_destinations[id].AuthenticatedMailboxData.Username, _destinations[id].AuthenticatedMailboxData.Password 
                        ,_mainWindow.EmptyFolderCheckBox.IsChecked == true);
                    return target;
                }
            }
            return null;
        }

        public void AddDestination(string id)
        {
            lock (_destinationLock)
            {
                if (!_destinations.ContainsKey(id))
                {
                    var mailbox = new AuthenticatedMailboxTargetControl(id);
                    dataContext.Mailboxes.Add(mailbox);
                    _destinations.Add(id, mailbox);
                }
            }
        }

        public void RemoveDestination(string id)
        {
            lock (_destinationLock)
            {
                if (_destinations.ContainsKey(id))
                {
                    dataContext.Mailboxes.Remove(_destinations[id]);
                    _destinations.Remove(id);
                }
            }
        }
    }

    public class ExchangeDestinationDataContext
    {
        public ObservableCollection<AuthenticatedMailboxTargetControl> Mailboxes { get; set; }

        public String Server { get; set; }

        public ExchangeDestinationDataContext()
        {
            Mailboxes = new ObservableCollection<AuthenticatedMailboxTargetControl>();
        }
    }
}
