using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using log4net;

namespace Zinkuba.MailModule
{
    public class PCQueue<TQueueObject,TStateObject> where TQueueObject : new() where TStateObject : IPCState, new()
    {
        private static readonly ILog Logger = LogManager.GetLogger(typeof(PCQueue<TQueueObject, TStateObject>));
        private readonly string _id;
        private readonly BlockingCollection<TQueueObject> _queue;
        private int activeThreadCount = 0;

        public Func<TQueueObject, TStateObject, TStateObject> ProduceMethod;
        public Func<TStateObject> InitialiseProducer;
        public Action<TStateObject, Exception> ShutdownProducer;

        private readonly Thread _produceThread;
        private readonly CancellationTokenSource _cancel;
        public int Produced { get; private set; }
        public int Consumed { get; private set; }
        public int Count { get { return _queue.Count; } }

        private TStateObject _lastState;
        private bool _isClosed;

        public PCQueue(String id)
        {
            _id = id;
            _queue = new BlockingCollection<TQueueObject>(new ConcurrentQueue<TQueueObject>());
            _produceThread = new Thread(RunProducer) { IsBackground = true, Name = _id + "-producer"};
            _cancel = new CancellationTokenSource();
        }

        public void Consume(TQueueObject element)
        {
            while (!_cancel.Token.IsCancellationRequested && !IsClosed)
            {
                if (!_queue.TryAdd(element, 1000, _cancel.Token)) continue;
                Consumed++;
                break;
            }
        }

        private void RunProducer()
        {
            _lastState = InitialiseProducer == null ? new TStateObject() : InitialiseProducer();
            try
            {
                while (!_cancel.Token.IsCancellationRequested && (!IsClosed || (Produced < Consumed)))
                {
                    TQueueObject queueElement;
                    try
                    {
                        if (!_queue.TryTake(out queueElement, 1000, _cancel.Token)) continue;
                        Produced++;
                        _lastState = ProduceMethod(queueElement, _lastState);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error("Caught Exception while trying to get next element from the queue", ex);
                    }
                }
            }
            catch (Exception ex)
            {
                Close(ex);
            }
        }

        private void Close(Exception exception)
        {
            if (!IsClosed)
            {
                IsClosed = true;
                Logger.Debug("Closing queue");
                if (exception != null)
                {
                    Logger.Debug("Cancelling queue processor");
                    _cancel.Cancel();
                }
                Logger.Debug("Flushing queue");
                if (_produceThread != null && Thread.CurrentThread != _produceThread)
                {
                    _produceThread.Join();
                    if(_produceThread.IsAlive) {
                        try
                        {
                            Logger.Debug("Killing Queue processor");
                            _produceThread.Interrupt();
                        }
                        catch
                        {
                        }
                    }
                }
                ShutdownProducer(_lastState, exception);
            }
        }

        public void Close()
        {
            Close(null);
        }

        public bool IsClosed
        {
            get { return _isClosed; }
            private set { _isClosed = value;}
        }

        public void Start()
        {
            if (ProduceMethod == null || InitialiseProducer == null || ShutdownProducer == null)
            {
                throw new Exception("Cannot start, ProduceMethod|InitialiseProducer|ShutdownProducer is not set");
            }
            _produceThread.Start();
        }
    }

}
