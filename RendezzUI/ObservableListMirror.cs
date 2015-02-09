using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Reflection;
using System.Windows.Threading;
using log4net;

namespace Rendezz.UI
{
    public abstract class ObservableListMirror<TIncomingType, TOutgoingType> : IEnumerable<TOutgoingType>, INotifyCollectionChanged where TOutgoingType : IReflectedObject<TIncomingType>
    {
        private static readonly ILog Logger =
            LogManager.GetLogger(typeof(ObservableListMirror<TIncomingType, TOutgoingType>));

        public ObservableCollection<TOutgoingType> List { get; private set; }
        protected readonly Dictionary<TIncomingType, TOutgoingType> ListLookup = new Dictionary<TIncomingType, TOutgoingType>();
        protected readonly Dispatcher Dispatcher;

        protected ObservableListMirror(Dispatcher dispatcher)
        {
            this.Dispatcher = dispatcher;
            // Create a local copy of the list
            List = new ObservableCollection<TOutgoingType>();
            List.CollectionChanged += (sender, args) => OnCollectionChanged(args);
        }

        public void SetList(ObservableCollection<TIncomingType> list)
        {
            List.Clear();
            foreach (var listElement in list)
            {
                AddElement(listElement);
            }
            list.CollectionChanged += ListChanged;
        }

        protected abstract TOutgoingType CreateNew(TIncomingType sourceObject);

        public TOutgoingType GetElement(TIncomingType lookup)
        {
            if (ListLookup.ContainsKey(lookup))
            {
                return ListLookup[lookup];
            }
            throw new Exception("Cannot find element identified by " + lookup);
        }

        private void AddElement(TIncomingType listElement)
        {
            if (ListLookup.ContainsKey(listElement))
            {
                Logger.Warn("Requested to add item " + listElement + ", but it exists already in the list");
            }
            else
            {
                try
                {
                    TOutgoingType newElement = CreateNew(listElement);
                    ListLookup.Add(listElement, newElement);
                    List.Add(newElement);
                }
                catch (Exception e)
                {
                    Logger.Error("Cannot create new listElement : " + e.Message, e);
                }
            }
        }

        private void ListChanged(object sender, NotifyCollectionChangedEventArgs eventArgs)
        {
            // Modifications to the local collection must be done in UI thread.. and so... :)
            if (Dispatcher.CheckAccess())
            {
                ProcessAction(eventArgs);
            }
            else
            {
                this.Dispatcher.Invoke(new Action(() => ProcessAction(eventArgs)));
            }
        }

        private void RemoveElement(TIncomingType listElement)
        {
            if (ListLookup.ContainsKey(listElement))
            {
                var element = ListLookup[listElement];
                List.Remove(element);
                ListLookup.Remove(listElement);
            }
            else
            {
                Logger.Warn("Requested to remove item " + listElement + ", but it doesn't exist");
            }
        }

        private void ProcessAction(NotifyCollectionChangedEventArgs eventArgs)
        {
            switch (eventArgs.Action)
            {
                case NotifyCollectionChangedAction.Add:
                    {
                        if (eventArgs.NewItems != null)
                        {
                            foreach (TIncomingType newItem in eventArgs.NewItems)
                            {
                                AddElement(newItem);
                            }
                        }
                    }
                    break;
                case NotifyCollectionChangedAction.Remove:
                    {
                        if (eventArgs.OldItems != null)
                        {
                            foreach (TIncomingType oldItem in eventArgs.OldItems)
                            {
                                RemoveElement(oldItem);
                            }
                        }

                    }
                    break;
                case NotifyCollectionChangedAction.Reset:
                    {
                        List.Clear();
                        ListLookup.Clear();
                    }
                    break;
                default:
                    {
                        Logger.Warn("List changed with unhandled action " +
                                    eventArgs.Action);
                    }
                    break;
            }
        }

        public IEnumerator<TOutgoingType> GetEnumerator()
        {
            return List.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public event NotifyCollectionChangedEventHandler CollectionChanged;

        protected virtual void OnCollectionChanged(NotifyCollectionChangedEventArgs e)
        {
            NotifyCollectionChangedEventHandler handler = CollectionChanged;
            if (handler != null) handler(this, e);
        }
    }
}
