using System;
using Rendezz.UI;
using Zinkuba.App.MailAccount;

namespace Zinkuba.App
{
    public interface IMailAccountControl : IReflectedObject<IMailAccount>
    {
        IMailAccount Account { get; }
    }
}