using System;
using System.Windows.Threading;
using Rendezz.UI;

namespace Zinkuba.App.Mailbox
{
    public class AuthenticatedMailboxList : ObservableListMirror<AuthenticatedMailbox, AuthenticatedMailboxControl>
    {
        private readonly Action<AuthenticatedMailbox> _removeMailboxAction;

        public AuthenticatedMailboxList(Action<AuthenticatedMailbox> removeMailboxAction, Dispatcher dispatcher)
            : base(dispatcher)
        {
            _removeMailboxAction = removeMailboxAction;
        }

        protected override AuthenticatedMailboxControl CreateNew(AuthenticatedMailbox sourceObject)
        {
            return new AuthenticatedMailboxControl(sourceObject)
            {
                RemoveMailboxFunction = _removeMailboxAction
            };
        }
    }
}