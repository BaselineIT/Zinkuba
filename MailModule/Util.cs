using System;

namespace Zinkuba.MailModule
{
    public class Util
    {
        private readonly Random _random;
        private static Util _oneAndOnly;
        private static readonly object LockObj = new object();

        private Util()
        {
            _random = new Random();
        }

        public static Util Instance
        {
            get
            {
                lock (LockObj)
                {
                    if (_oneAndOnly == null) return _oneAndOnly = new Util();
                    return _oneAndOnly;
                }
            }
        }

        public Random Random
        {
            get { return _random; }
        }
    }
}
