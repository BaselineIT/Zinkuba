using System;
using System.Collections.Generic;
using System.Windows.Controls;
using Zinkuba.MailModule.MessageProcessor;

namespace Zinkuba.App
{
    /// <summary>
    /// Interaction logic for PstDestinationControl.xaml
    /// </summary>
    public partial class PstDestinationControl : UserControl, IDestinationManager
    {
        private Dictionary<string, PstTarget> _destinations;
        private readonly object _destinationLock = new object();


        public PstDestinationControl()
        {
            InitializeComponent();
            SaveFolder.Text = Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Downloads");
            _destinations = new Dictionary<String,PstTarget>();
        }

        public IMessageDestination GetDestination(String id)
        {
            // we don't care about the id, we are id dependant
            return new PstTarget(SaveFolder.Text);
        }

        public void AddDestination(string id)
        {
        }

        public void RemoveDestination(string id)
        {
        }
    }
}
