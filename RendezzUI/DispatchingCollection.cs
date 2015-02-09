using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Windows.Threading;
using log4net;

namespace Rendezz.UI
{
    /// <summary>
    /// Implements a collection that wraps another collection and dispatches all collection change notifications to the
    /// <see cref="Dispatcher"/>s thread.
    /// </summary>
    /// <typeparam name="TCollection">
    /// The type of the underlying collection.
    /// </typeparam>
    /// <typeparam name="TItem">
    /// The type of items stored in the underlying collection (and consequently by this collection).
    /// </typeparam>
    public class DispatchingCollection<TCollection, TItem> : ICollection<TItem>, INotifyCollectionChanged
           where TCollection : ICollection<TItem>, INotifyCollectionChanged
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(DispatchingCollection<TCollection, TItem>));
        private readonly Dispatcher _dispatcher;
        private readonly TCollection _underlyingCollection;

        /// <summary>
        /// Gets the collection that is being wrapped by this collection.
        /// </summary>

        public TCollection UnderlyingCollection
        {
            get
            {
                return _underlyingCollection;
            }
        }

        /// <summary>
        /// Gets the <see cref="Dispatcher"/> to which collection change notifications will be marshalled.
        /// </summary>

        public Dispatcher Dispatcher
        {
            get
            {
                return _dispatcher;
            }
        }

        /// <summary>
        /// Gets the number of items in this collection.
        /// </summary>

        public int Count
        {
            get
            {
                return _underlyingCollection.Count;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this collection is read-only.
        /// </summary>
        public bool IsReadOnly
        {
            get
            {
                return _underlyingCollection.IsReadOnly;
            }
        }

        /// <summary>
        /// Occurs whenever this collection changes.
        /// </summary>
        /// <remarks>
        /// <para>
        /// Unlike the underlying collection's <c>CollectionChanged</c> event, this event is guaranteed to execute on the
        /// <see cref="Dispatcher"/>'s thread.
        /// </para>
        /// </remarks>

        public event NotifyCollectionChangedEventHandler CollectionChanged;

        /// <summary>
        /// Constructs an instance of <c>DispatchingCollection</c>.
        /// </summary>
        /// <param name="underlyingCollection">
        /// The collection being wrapped by this dispatching collection.
        /// </param>
        /// <param name="dispatcher">
        /// The <see cref="Dispatcher"/> to which <see cref="CollectionChanged"/> notifications will be marshalled.
        /// </param>

        public DispatchingCollection(TCollection underlyingCollection, Dispatcher dispatcher)
        {
            if (underlyingCollection == null)
            {
                throw new Exception("Incoming collection is null");
            }
            if (dispatcher == null)
            {
                throw new Exception("Incoming dispatcher is null");
            }
            _underlyingCollection = underlyingCollection;
            _dispatcher = dispatcher;
            _underlyingCollection.CollectionChanged += delegate(object sender, NotifyCollectionChangedEventArgs e)
            {
                OnCollectionChanged(e);
            };
        }

        private delegate void AddHandler(TItem item);

        /// <summary>
        /// Adds an item to this collection.
        /// </summary>
        /// <param name="item">
        /// The item to add.
        /// </param>

        public void Add(TItem item)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(DispatcherPriority.Send, new AddHandler(Add), item);
            }
            else
            {
                _underlyingCollection.Add(item);
            }
        }

        private delegate void ClearHandler();

        /// <summary>
        /// Clears this collection.
        /// </summary>

        public void Clear()
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(DispatcherPriority.Send, new ClearHandler(Clear));
            }
            else
            {
                _underlyingCollection.Clear();
            }
        }


        /// <summary>
        /// Determines whether an item is present in this collection.
        /// </summary>
        /// <param name="item">
        /// The item to check for.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if <paramref name="item"/> is contained in this collection, otherwise <see langword="false"/>.
        /// </returns>

        public bool Contains(TItem item)
        {
            return _underlyingCollection.Contains(item);
        }

        /// <summary>
        /// Copies the items in this collection to the specified array.
        /// </summary>
        /// <param name="array">
        /// The array to which items will be copied.
        /// </param>
        /// <param name="arrayIndex">
        /// The starting index.
        /// </param>

        public void CopyTo(TItem[] array, int arrayIndex)
        {
            _underlyingCollection.CopyTo(array, arrayIndex);
        }

        private delegate bool RemoveHandler(TItem item);

        /// <summary>
        /// Removes an item from this collection.
        /// </summary>
        /// <param name="item">
        /// The item to be removed.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if the item was successfully removed, otherwise <see langword="false"/>.
        /// </returns>

        public bool Remove(TItem item)
        {

            if (!Dispatcher.CheckAccess())
            {
                return (bool)Dispatcher.Invoke(DispatcherPriority.Send, new RemoveHandler(Remove), item);
            }

            else
            {
                return _underlyingCollection.Remove(item);
            }
        }

        /// <summary>
        /// Gets an enumerator that iterates over items in this collection.
        /// </summary>
        /// <returns>
        /// An <see cref="IEnumerator{TItem}"/> that iterates over the items in this collection.
        /// </returns>

        public IEnumerator<TItem> GetEnumerator()
        {
            return _underlyingCollection.GetEnumerator();
        }

        /// <summary>
        /// Gets an enumerator that iterates over items in this collection.
        /// </summary>
        /// <returns>
        /// An <see cref="IEnumerator"/> that iterates over the items in this collection.
        /// </returns>

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return (_underlyingCollection as System.Collections.ICollection).GetEnumerator();
        }

        private delegate void OnCollectionChangedHandler(NotifyCollectionChangedEventArgs e);

        private void OnCollectionChanged(NotifyCollectionChangedEventArgs e)
        {
            //marshal the event to the dispatcher's thread
            if (_dispatcher.CheckAccess())
            {
                NotifyCollectionChangedEventHandler handler = CollectionChanged;
                if (handler != null) handler(this, e);
            }
            else
            {
                _dispatcher.Invoke(DispatcherPriority.Send, new OnCollectionChangedHandler(OnCollectionChanged), e);
            }
        }
    }
}
