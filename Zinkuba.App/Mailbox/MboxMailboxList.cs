using System;
using System.Windows.Threading;
using Rendezz.UI;

namespace Zinkuba.App.Mailbox
{
    public class MboxMailboxList : ObservableListMirror<MboxMailbox, MboxMailboxControl>
    {
        private readonly Action<MboxMailbox> _removeMailboxAction;

        public MboxMailboxList(Action<MboxMailbox> removeMailboxAction, Dispatcher dispatcher)
            : base(dispatcher)
        {
            _removeMailboxAction = removeMailboxAction;
        }

        protected override MboxMailboxControl CreateNew(MboxMailbox sourceObject)
        {
            return new MboxMailboxControl(sourceObject)
            {
                RemoveMailboxFunction = _removeMailboxAction
            };
        }
    }
}