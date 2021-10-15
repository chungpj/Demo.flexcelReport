using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Threading;

namespace Report.Common
{
    [Serializable, DataContract]
    [ComVisible(false)]
    public sealed class SynchronizedDictionary<TKey, TValue> :
        IDictionary<TKey, TValue>, ICollection<KeyValuePair<TKey, TValue>>,
        IEnumerable<KeyValuePair<TKey, TValue>>, IDictionary, ICollection, IEnumerable//, ISerializable, IDeserializationCallback
    {
        private readonly IDictionary<TKey, TValue> source;
        private readonly ReaderWriterLockSlim locker = new ReaderWriterLockSlim();

        #region Constructors

        // Summary:
        //     Initializes a new instance of the System.Collections.Generic.SynchronizedDictionary<TKey,TValue>
        //     class that is empty, has the default initial capacity, and uses the default
        //     equality comparer for the key type.
        public SynchronizedDictionary()
        {
            this.locker = new ReaderWriterLockSlim(LockRecursionPolicy.SupportsRecursion);
            this.source = new Dictionary<TKey, TValue>();
        }
        //
        // Summary:
        //     Initializes a new instance of the System.Collections.Generic.SynchronizedDictionary<TKey,TValue>
        //     class that contains elements copied from the specified System.Collections.Generic.IDictionary<TKey,TValue>
        //     and uses the default equality comparer for the key type.
        //
        // Parameters:
        //   dictionary:
        //     The System.Collections.Generic.IDictionary<TKey,TValue> whose elements are
        //     copied to the new System.Collections.Generic.SynchronizedDictionary<TKey,TValue>.
        //
        // Exceptions:
        //   System.ArgumentNullException:
        //     dictionary is null.
        //
        //   System.ArgumentException:
        //     dictionary contains one or more duplicate keys.
        public SynchronizedDictionary(IDictionary<TKey, TValue> dictionary)
        {
            this.locker = new ReaderWriterLockSlim(LockRecursionPolicy.SupportsRecursion);
            this.source = new Dictionary<TKey, TValue>(dictionary);
        }
        //
        // Summary:
        //     Initializes a new instance of the System.Collections.Generic.SynchronizedDictionary<TKey,TValue>
        //     class that is empty, has the default initial capacity, and uses the specified
        //     System.Collections.Generic.IEqualityComparer<T>.
        //
        // Parameters:
        //   comparer:
        //     The System.Collections.Generic.IEqualityComparer<T> implementation to use
        //     when comparing keys, or null to use the default System.Collections.Generic.EqualityComparer<T>
        //     for the type of the key.
        public SynchronizedDictionary(IEqualityComparer<TKey> comparer)
        {
            this.locker = new ReaderWriterLockSlim(LockRecursionPolicy.SupportsRecursion);
            this.source = new Dictionary<TKey, TValue>(comparer);
        }
        //
        // Summary:
        //     Initializes a new instance of the System.Collections.Generic.SynchronizedDictionary<TKey,TValue>
        //     class that is empty, has the specified initial capacity, and uses the default
        //     equality comparer for the key type.
        //
        // Parameters:
        //   capacity:
        //     The initial number of elements that the System.Collections.Generic.SynchronizedDictionary<TKey,TValue>
        //     can contain.
        //
        // Exceptions:
        //   System.ArgumentOutOfRangeException:
        //     capacity is less than 0.
        public SynchronizedDictionary(int capacity)
        {
            this.locker = new ReaderWriterLockSlim(LockRecursionPolicy.SupportsRecursion);
            this.source = new Dictionary<TKey, TValue>(capacity);
        }
        //
        // Summary:
        //     Initializes a new instance of the System.Collections.Generic.SynchronizedDictionary<TKey,TValue>
        //     class that contains elements copied from the specified System.Collections.Generic.IDictionary<TKey,TValue>
        //     and uses the specified System.Collections.Generic.IEqualityComparer<T>.
        //
        // Parameters:
        //   dictionary:
        //     The System.Collections.Generic.IDictionary<TKey,TValue> whose elements are
        //     copied to the new System.Collections.Generic.SynchronizedDictionary<TKey,TValue>.
        //
        //   comparer:
        //     The System.Collections.Generic.IEqualityComparer<T> implementation to use
        //     when comparing keys, or null to use the default System.Collections.Generic.EqualityComparer<T>
        //     for the type of the key.
        //
        // Exceptions:
        //   System.ArgumentNullException:
        //     dictionary is null.
        //
        //   System.ArgumentException:
        //     dictionary contains one or more duplicate keys.
        public SynchronizedDictionary(IDictionary<TKey, TValue> dictionary, IEqualityComparer<TKey> comparer)
        {
            this.locker = new ReaderWriterLockSlim(LockRecursionPolicy.SupportsRecursion);
            this.source = new Dictionary<TKey, TValue>(dictionary, comparer);
        }
        //
        // Summary:
        //     Initializes a new instance of the System.Collections.Generic.SynchronizedDictionary<TKey,TValue>
        //     class that is empty, has the specified initial capacity, and uses the specified
        //     System.Collections.Generic.IEqualityComparer<T>.
        //
        // Parameters:
        //   capacity:
        //     The initial number of elements that the System.Collections.Generic.SynchronizedDictionary<TKey,TValue>
        //     can contain.
        //
        //   comparer:
        //     The System.Collections.Generic.IEqualityComparer<T> implementation to use
        //     when comparing keys, or null to use the default System.Collections.Generic.EqualityComparer<T>
        //     for the type of the key.
        //
        // Exceptions:
        //   System.ArgumentOutOfRangeException:
        //     capacity is less than 0.
        public SynchronizedDictionary(int capacity, IEqualityComparer<TKey> comparer)
        {
            this.locker = new ReaderWriterLockSlim(LockRecursionPolicy.SupportsRecursion);
            this.source = new Dictionary<TKey, TValue>(comparer);
        }
        /*
        //
        // Summary:
        //     Initializes a new instance of the System.Collections.Generic.SynchronizedDictionary<TKey,TValue>
        //     class with serialized data.
        //
        // Parameters:
        //   info:
        //     A System.Runtime.Serialization.SerializationInfo object containing the information
        //     required to serialize the System.Collections.Generic.SynchronizedDictionary<TKey,TValue>.
        //
        //   context:
        //     A System.Runtime.Serialization.StreamingContext structure containing the
        //     source and destination of the serialized stream associated with the System.Collections.Generic.SynchronizedDictionary<TKey,TValue>.
        protected SynchronizedDictionary(SerializationInfo info, StreamingContext context)
        {
            this.locker = new ReaderWriterLockSlim(LockRecursionPolicy.SupportsRecursion);
            this.source = new Dictionary<TKey, TValue>(info, context);
        }
*/

        private static void ThrowNotSupportedException()
        {
            throw new NotSupportedException("Synchoronized dictionary is not supported this action");
        }

        #endregion  Constructors

        #region IDictionary<TKey, TValue>'s members

        public void Add(TKey key, TValue value)
        {
            this.locker.EnterWriteLock();
            try
            {
                this.source.Add(key, value);
            }
            finally
            {
                this.locker.ExitWriteLock();
            }
        }

        //void IDictionary.Clear()
        public void Clear()
        {
            this.locker.EnterWriteLock();
            try
            {
                this.source.Clear();
            }
            finally
            {
                this.locker.ExitWriteLock();
            }
        }

        public int Count
        {
            get
            {
                this.locker.EnterReadLock();
                try
                {
                    return this.source.Count;
                }
                finally
                {
                    this.locker.ExitReadLock();
                }
            }
        }

        public bool ContainsKey(TKey key)
        {
            this.locker.EnterReadLock();
            try
            {
                return this.source.ContainsKey(key);
            }
            finally
            {
                this.locker.ExitReadLock();
            }
        }

        public bool ContainsValue(TValue value)
        {
            this.locker.EnterReadLock();
            try
            {
                return ((Dictionary<TKey, TValue>)this.source).ContainsValue(value);
            }
            finally
            {
                this.locker.ExitReadLock();
            }
        }

        public bool Remove(TKey key)
        {
            this.locker.EnterWriteLock();
            try
            {
                return this.source.Remove(key);
            }
            finally
            {
                this.locker.ExitWriteLock();
            }
        }

        public bool TryGetValue(TKey key, out TValue value)
        {
            this.locker.EnterReadLock();
            try
            {
                return this.source.TryGetValue(key, out value);
            }
            finally
            {
                this.locker.ExitReadLock();
            }
        }

        public TValue GetValueOrDefault(TKey key)
        {
            this.locker.EnterReadLock();
            try
            {
                return this.source.ContainsKey(key) ? this.source[key] : default(TValue);
            }
            finally
            {
                this.locker.ExitReadLock();
            }
        }

        public TValue GetValueOrAdd(TKey key, Func<TKey, TValue> valuer)
        {
            this.locker.EnterReadLock();
            try
            {
                if (this.source.ContainsKey(key))
                    return this.source[key];

                this.locker.ExitReadLock();
                this.locker.EnterWriteLock();
                try
                {
                    if (!this.source.ContainsKey(key))
                        this.source.Add(key, valuer(key));
                    return this.source[key];
                }
                finally
                {
                    this.locker.ExitWriteLock();
                }
            }
            finally
            {
                if (this.locker.IsReadLockHeld)
                    this.locker.ExitReadLock();
            }
        }

        public TValue this[TKey key]
        {
            get
            {
                this.locker.EnterReadLock();
                try
                {
                    return this.source[key];
                }
                finally
                {
                    this.locker.ExitReadLock();
                }
            }
            set
            {
                this.locker.EnterWriteLock();
                try
                {
                    this.source[key] = value;
                }
                finally
                {
                    this.locker.ExitWriteLock();
                }
            }
        }

        #region SynchronizedDictionary<TKey, TValue>.KeyEnumerator Members

        private SynchronizedKeyCollection keys;

        ICollection<TKey> IDictionary<TKey, TValue>.Keys
        {
            get { return this.Keys; }
        }

        public SynchronizedDictionary<TKey, TValue>.SynchronizedKeyCollection Keys
        {
            get
            {
                return this.keys != null ? this.keys : (this.keys = new SynchronizedKeyCollection(this));
            }
        }

        [Serializable, DataContract]
        [DebuggerDisplay("Count = {Count}")]
        public sealed class SynchronizedKeyCollection : ICollection<TKey>, IEnumerable<TKey>, ICollection, IEnumerable
        {
            private SynchronizedDictionary<TKey, TValue> _synchronizedDictionary;
            public SynchronizedKeyCollection(SynchronizedDictionary<TKey, TValue> synchronizedDictionary)
            {
                _synchronizedDictionary = synchronizedDictionary;
            }

            #region IEnumerable Members

            IEnumerator IEnumerable.GetEnumerator()
            {
                return new Enumerator(this);
            }

            #endregion

            #region IEnumerable<TKey> Members

            public IEnumerator<TKey> GetEnumerator()
            {
                return new Enumerator(this);
            }

            [Serializable, DataContract]
            public struct Enumerator : IEnumerator<TKey>, IDisposable, IEnumerator
            {
                private SynchronizedKeyCollection _KeyEnumerator;
                private IEnumerator<TKey> _enumerator;
                public Enumerator(SynchronizedKeyCollection keyEnumerator)
                {
                    keyEnumerator._synchronizedDictionary.locker.EnterReadLock();
                    _KeyEnumerator = keyEnumerator;
                    _enumerator = keyEnumerator._synchronizedDictionary.source.Keys.GetEnumerator();
                }

                #region IDisposable Members

                public void Dispose()
                {
                    _enumerator.Dispose();
                    _KeyEnumerator._synchronizedDictionary.locker.ExitReadLock();
                }

                #endregion

                #region IEnumerator Members

                object IEnumerator.Current
                {
                    get { return _enumerator.Current; }
                }

                public bool MoveNext()
                {
                    return _enumerator.MoveNext();
                }

                public void Reset()
                {
                    ((IEnumerator)_enumerator).Reset();
                }

                #endregion

                #region IEnumerator<TKey> Members

                public TKey Current
                {
                    get { return _enumerator.Current; }
                }

                #endregion
            }

            #endregion

            #region ICollection<TKey> Members

            void ICollection<TKey>.Add(TKey item)
            {
                ThrowNotSupportedException();
            }

            void ICollection<TKey>.Clear()
            {
                ThrowNotSupportedException();
            }

            bool ICollection<TKey>.Contains(TKey item)
            {
                return this._synchronizedDictionary.ContainsKey(item);
            }

            void ICollection<TKey>.CopyTo(TKey[] array, int arrayIndex)
            {
                this._synchronizedDictionary.locker.EnterReadLock();
                try
                {
                    this._synchronizedDictionary.source.Keys.CopyTo(array, arrayIndex);
                }
                finally
                {
                    this._synchronizedDictionary.locker.ExitReadLock();
                }
            }

            public int Count
            {
                get { return this._synchronizedDictionary.Count; }
            }

            bool ICollection<TKey>.IsReadOnly
            {
                get { return true; }
            }

            bool ICollection<TKey>.Remove(TKey item)
            {
                ThrowNotSupportedException();
                return false;
            }

            #endregion

            #region ICollection Members

            void ICollection.CopyTo(Array array, int index)
            {
                this._synchronizedDictionary.locker.EnterReadLock();
                try
                {
                    ((ICollection)this._synchronizedDictionary.source.Keys).CopyTo(array, index);
                }
                finally
                {
                    this._synchronizedDictionary.locker.ExitReadLock();
                }
            }

            bool ICollection.IsSynchronized
            {
                get { return true; }
            }

            object ICollection.SyncRoot
            {
                get { return ((ICollection)this._synchronizedDictionary).SyncRoot; }
            }

            #endregion
        }

        #endregion

        #region SynchronizedDictionary<TKey, TValue>.ValueEnumerator Members

        private SynchronizedValueCollection values;

        ICollection<TValue> IDictionary<TKey, TValue>.Values { get { return this.Values; } }

        public SynchronizedDictionary<TKey, TValue>.SynchronizedValueCollection Values
        {
            get { return this.values != null ? this.values : (this.values = new SynchronizedValueCollection(this)); }
        }

        [Serializable, DataContract]
        [DebuggerDisplay("Count = {Count}")]
        public sealed class SynchronizedValueCollection : ICollection<TValue>, IEnumerable<TValue>, ICollection, IEnumerable
        {
            private SynchronizedDictionary<TKey, TValue> _synchronizedDictionary;
            public SynchronizedValueCollection(SynchronizedDictionary<TKey, TValue> synchronizedDictionary)
            {
                _synchronizedDictionary = synchronizedDictionary;
            }

            #region IEnumerable Members

            IEnumerator IEnumerable.GetEnumerator()
            {
                return new Enumerator(this);
            }

            #endregion

            #region IEnumerable<TValue> Members

            public IEnumerator<TValue> GetEnumerator()
            {
                return new Enumerator(this);
            }

            [Serializable, DataContract]
            public struct Enumerator : IEnumerator<TValue>, IDisposable, IEnumerator
            {
                private SynchronizedValueCollection _valueEnumerator;
                private IEnumerator<TValue> _enumerator;
                public Enumerator(SynchronizedValueCollection valueEnumerator)
                {
                    valueEnumerator._synchronizedDictionary.locker.EnterReadLock();
                    _valueEnumerator = valueEnumerator;
                    _enumerator = valueEnumerator._synchronizedDictionary.source.Values.GetEnumerator();
                }

                #region IDisposable Members

                public void Dispose()
                {
                    _enumerator.Dispose();
                    _valueEnumerator._synchronizedDictionary.locker.ExitReadLock();
                }

                #endregion

                #region IEnumerator Members

                object IEnumerator.Current
                {
                    get { return _enumerator.Current; }
                }

                public bool MoveNext()
                {
                    return _enumerator.MoveNext();
                }

                public void Reset()
                {
                    ((IEnumerator)_enumerator).Reset();
                }

                #endregion

                #region IEnumerator<TValue> Members

                public TValue Current
                {
                    get { return _enumerator.Current; }
                }

                #endregion
            }

            #endregion

            #region ICollection<TValue> Members

            void ICollection<TValue>.Add(TValue item)
            {
                ThrowNotSupportedException();
            }

            void ICollection<TValue>.Clear()
            {
                ThrowNotSupportedException();
            }

            bool ICollection<TValue>.Contains(TValue item)
            {
                return this._synchronizedDictionary.ContainsValue(item);
            }

            void ICollection<TValue>.CopyTo(TValue[] array, int arrayIndex)
            {
                this._synchronizedDictionary.locker.EnterReadLock();
                try
                {
                    this._synchronizedDictionary.source.Values.CopyTo(array, arrayIndex);
                }
                finally
                {
                    this._synchronizedDictionary.locker.ExitReadLock();
                }
            }

            public int Count
            {
                get { return this._synchronizedDictionary.Count; }
            }

            bool ICollection<TValue>.IsReadOnly
            {
                get { return true; }
            }

            bool ICollection<TValue>.Remove(TValue item)
            {
                ThrowNotSupportedException();
                return false;
            }

            #endregion

            #region ICollection Members

            void ICollection.CopyTo(Array array, int index)
            {
                this._synchronizedDictionary.locker.EnterReadLock();
                try
                {
                    ((ICollection)this._synchronizedDictionary.source.Keys).CopyTo(array, index);
                }
                finally
                {
                    this._synchronizedDictionary.locker.ExitReadLock();
                }
            }

            bool ICollection.IsSynchronized
            {
                get { return true; }
            }

            object ICollection.SyncRoot
            {
                get { return ((ICollection)this._synchronizedDictionary).SyncRoot; }
            }

            #endregion
        }

        #endregion

        #endregion IDictionary<TKey, TValue>'s members

        #region ICollection<KeyValuePair<TKey, TValue>>'s members

        void ICollection<KeyValuePair<TKey, TValue>>.Add(KeyValuePair<TKey, TValue> item)
        {
            this.locker.EnterWriteLock();
            try
            {
                this.source.Add(item);
            }
            finally
            {
                this.locker.ExitWriteLock();
            }
        }

        bool ICollection<KeyValuePair<TKey, TValue>>.Contains(KeyValuePair<TKey, TValue> item)
        {
            this.locker.EnterReadLock();
            try
            {
                return ((ICollection<KeyValuePair<TKey, TValue>>)this.source).Contains(item);
            }
            finally
            {
                this.locker.ExitReadLock();
            }
        }

        void ICollection<KeyValuePair<TKey, TValue>>.CopyTo(KeyValuePair<TKey, TValue>[] array, int arrayIndex)
        {
            this.locker.EnterReadLock();
            try
            {
                ((ICollection<KeyValuePair<TKey, TValue>>)this.source).CopyTo(array, arrayIndex);
            }
            finally
            {
                this.locker.ExitReadLock();
            }
        }

        bool ICollection<KeyValuePair<TKey, TValue>>.IsReadOnly
        {
            get { return false; }
        }

        bool ICollection<KeyValuePair<TKey, TValue>>.Remove(KeyValuePair<TKey, TValue> item)
        {
            this.locker.EnterReadLock();
            try
            {
                return ((ICollection<KeyValuePair<TKey, TValue>>)this.source).Remove(item);
            }
            finally
            {
                this.locker.ExitReadLock();
            }
        }

        #endregion ICollection<KeyValuePair<TKey, TValue>>'s members

        #region IDictionary's members

        void IDictionary.Add(object key, object value)
        {
            this.locker.EnterWriteLock();
            try
            {
                ((IDictionary)this.source).Add(key, value);
            }
            finally
            {
                this.locker.ExitWriteLock();
            }
        }

        bool IDictionary.Contains(object key)
        {
            this.locker.EnterReadLock();
            try
            {
                return ((IDictionary)this.source).Contains(key);
            }
            finally
            {
                this.locker.ExitReadLock();
            }
        }

        IDictionaryEnumerator IDictionary.GetEnumerator()
        {
            return new SynchronizedEnumerator(this);
        }

        bool IDictionary.IsFixedSize { get { return false; } }

        bool IDictionary.IsReadOnly { get { return false; } }

        ICollection IDictionary.Keys
        {
            get { return this.Keys; }
        }

        void IDictionary.Remove(object key)
        {
            this.locker.EnterWriteLock();
            try
            {
                ((IDictionary)this.source).Remove(key);
            }
            finally
            {
                this.locker.ExitWriteLock();
            }
        }

        ICollection IDictionary.Values
        {
            get { return this.Values; }
        }

        object IDictionary.this[object key]
        {
            get { return this[(TKey)key]; }
            set { this[(TKey)key] = (TValue)value; }
        }

        #endregion IDictionary's members

        #region ICollection's members

        void ICollection.CopyTo(Array array, int index)
        {
            this.locker.EnterReadLock();
            try
            {
                ((ICollection)this.source).CopyTo(array, index);
            }
            finally
            {
                this.locker.ExitReadLock();
            }
        }

        bool ICollection.IsSynchronized
        {
            get { return true; }
        }

        object ICollection.SyncRoot
        {
            get { return ((ICollection)this.source).SyncRoot; }
        }

        #endregion ICollection's members

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        #endregion

        #region IEnumerable<KeyValuePair<TKey,TValue>> Members

        public IEnumerator<KeyValuePair<TKey, TValue>> GetEnumerator()
        {
            return new SynchronizedEnumerator(this);
        }

        [Serializable, DataContract]
        public struct SynchronizedEnumerator : IEnumerator<KeyValuePair<TKey, TValue>>, IDisposable, IDictionaryEnumerator, IEnumerator
        {
            private SynchronizedDictionary<TKey, TValue> _synchronizedDictionary;
            private IEnumerator<KeyValuePair<TKey, TValue>> _enumerator;

            public SynchronizedEnumerator(SynchronizedDictionary<TKey, TValue> synchronizedDictionary)
            {
                synchronizedDictionary.locker.EnterReadLock();
                _synchronizedDictionary = synchronizedDictionary;
                _enumerator = synchronizedDictionary.source.GetEnumerator();
            }

            #region IEnumerator<KeyValuePair<TKey, TValue>> Members

            public void Dispose()
            {
                _enumerator.Dispose();
                _synchronizedDictionary.locker.ExitReadLock();
            }

            KeyValuePair<TKey, TValue> IEnumerator<KeyValuePair<TKey, TValue>>.Current
            {
                get { return _enumerator.Current; }
            }

            public bool MoveNext()
            {
                return _enumerator.MoveNext();
            }

            #endregion

            #region IEnumerator Members

            object IEnumerator.Current
            {
                get { return ((IDictionaryEnumerator)_enumerator).Entry; }
            }

            public void Reset()
            {
                _enumerator.Reset();
            }

            #endregion

            #region IDictionaryEnumerator Members

            public DictionaryEntry Entry
            {
                get { return ((IDictionaryEnumerator)_enumerator).Entry; }
            }

            public object Key
            {
                get { return ((IDictionaryEnumerator)_enumerator).Key; }
            }

            public object Value
            {
                get { return ((IDictionaryEnumerator)_enumerator).Value; }
            }

            #endregion
        }

        #endregion

    }
}
