using System;
using System.Collections.Generic;
using System.Windows.Documents;
using Zinkuba.MailModule.MessageProcessor;

namespace Zinkuba.App
{
    internal interface IDestinationManager
    {
        IMessageDestination GetDestination(String id);
        void AddDestination(String id);
        void RemoveDestination(String id);
    }
}