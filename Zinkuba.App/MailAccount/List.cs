using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using Zinkuba.App.Mailbox;

namespace Zinkuba.App.MailAccount

{
    public delegate void CollectionChangedDelegate(NotifyCollectionChangedEventArgs e);

    public class ReturnTypeCollection<T>
    {
        private dynamic _underlyingCollection;

        public dynamic UnderlyingCollection
        {
            set { _underlyingCollection = value;
            _underlyingCollection.CollectionChanged += new NotifyCollectionChangedEventHandler(OnCollectionChanged);
            }
            get { return _underlyingCollection; }
        }

        public int Count { get { return UnderlyingCollection.Count; } }

        public T this[int index]
        {
            get
            {
                return UnderlyingCollection[index];
            }
            set
            {
                UnderlyingCollection[index] = value;
            }
        }

        public void Add<M>(T value) where M : class
        {
            UnderlyingCollection.Add(value as M);
        }

        public bool Remove<M>(T value) where M : class
        {
            return UnderlyingCollection.Remove(value as M);
        }

        public bool Contains<M>(T value) where M : class
        {
            return UnderlyingCollection.Contains(value as M);
        }

        public event EventHandler<NotifyCollectionChangedEventArgs> CollectionChanged;

        protected virtual void OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs notifyCollectionChangedEventArgs)
        {
            EventHandler<NotifyCollectionChangedEventArgs> handler = CollectionChanged;
            if (handler != null) handler(this, notifyCollectionChangedEventArgs);
        }

        public void Clear()
        {
            UnderlyingCollection.Clear();
        }
    }
}