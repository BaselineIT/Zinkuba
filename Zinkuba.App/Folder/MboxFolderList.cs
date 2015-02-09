using System;
using System.Windows.Threading;
using Rendezz.UI;
using Zinkuba.App.Mailbox;

namespace Zinkuba.App.Folder
{
    public class MboxFolderList : ObservableListMirror<MboxFolder, MboxFolderControl>
    {
        private readonly Action<MboxFolder> _removeFolderAction;

        public MboxFolderList(Action<MboxFolder> removeFolderAction, Dispatcher dispatcher)
            : base(dispatcher)
        {
            _removeFolderAction = removeFolderAction;
        }

        protected override MboxFolderControl CreateNew(MboxFolder sourceObject)
        {
            return new MboxFolderControl(sourceObject)
            {
                RemoveFolderFunction = _removeFolderAction
            };
        }
    }
}