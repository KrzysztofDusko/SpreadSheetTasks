using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

//https://github.com/MarkPflug/Sylvan/blob/main/source/Sylvan.Common/StringPool.cs
namespace Benchmark;

/// <summary>
/// An IStringFactory implementation that provides string de-duping capabilities..
/// </summary>
public sealed partial class SimpleStringPool
{
    const int DefaultCapacity = 64;
    const int DefaultSizeLimit = 32;
    const int CollisionLimit = 8;

    // This is a greatly-simplified HashSet<string> that only allows additions.
    // and accepts char[] instead of string.

    // An extremely simple, and hopefully fast, hash algorithm.
    static uint GetHashCode(ReadOnlySpan<char> buffer)
    {
        uint hash = 0;
        int offset = 0;
        for (int i = 0; i < buffer.Length; i++)
        {
            hash = hash * 31 + buffer[offset++];
        }
        return hash;
    }

    //IEnumerable<(string str, int count)> GetUsage()
    //{
    //    for (int i = 0; i < this.buckets.Length; i++)
    //    {
    //        var b = buckets[i];
    //        if (b != 0)
    //        {
    //            var idx = b - 1;
    //            while ((uint)idx < entries.Length)
    //            {
    //                var e = this.entries[idx];
    //                yield return (e.str, e.count);
    //                idx = e.next;
    //            }
    //        }
    //    }
    //}

    readonly int stringSizeLimit;
    int[] buckets; // contains index into entries offset by -1. So that 0 (default) means empty bucket.
    Entry[] entries;

    int count;
    //long uniqueLen;
    //long dupeLen;
    //long skipLen;

    /// <summary>
    /// Creates a new StringPool instance.
    /// </summary>
    public SimpleStringPool() : this(DefaultSizeLimit) { }

    /// <summary>
    /// Creates a new StringPool instance.
    /// </summary>
    /// <param name="stringSizeLimit">The size limit beyond which strings will not be pooled.</param>
    /// <remarks>
    /// The <paramref name="stringSizeLimit"/> prevents pooling strings beyond a certain size. 
    /// Longer strings are typically less likely to be duplicated, and and carry extra cost for identifying uniqueness.
    /// </remarks>
    public SimpleStringPool(int stringSizeLimit)
    {
        int size = GetSize(DefaultCapacity);
        this.stringSizeLimit = stringSizeLimit;
        this.buckets = new int[size];
        this.entries = new Entry[size];
    }

    static int GetSize(int capacity)
    {
        var size = DefaultCapacity;
        while (size < capacity)
            size *= 2;
        return size;
    }

    /// <summary>
    /// Gets a string containing the characters in the input buffer.
    /// </summary>
    public string GetString(ReadOnlySpan<char> buffer)
    {
        if (buffer == null) throw new ArgumentNullException(nameof(buffer));
        if (buffer.Length == 0) return string.Empty;
        if (buffer.Length > stringSizeLimit)
        {
            //this.skipLen += buffer.Length;
            return new string(buffer);
        }

        var entries = this.entries;
        var hashCode = GetHashCode(buffer);

        uint collisionCount = 0;
        ref int bucket = ref GetBucket(hashCode);
        int i = bucket - 1;

        while ((uint)i < (uint)entries.Length)
        {
            ref var e = ref entries[i];
            if (e.hashCode == hashCode && MemoryExtensions.Equals(buffer, e.str.AsSpan(), StringComparison.Ordinal))
            {
                e.count++;
                //this.dupeLen += buffer.Length;
                return e.str;
            }

            i = e.next;

            collisionCount++;
            if (collisionCount > CollisionLimit)
            {
                // protects against malicious inputs
                // too many collisions give up and let the caller create the string.					
                return new string(buffer);
            }
        }

        int count = this.count;
        if (count == entries.Length)
        {
            entries = Resize();
            bucket = ref GetBucket(hashCode);
        }
        int index = count;
        this.count = count + 1;

        var stringValue = new string(buffer);

        ref Entry entry = ref entries![index];
        //this.uniqueLen += buffer.Length;
        entry.hashCode = hashCode;
        entry.count = 1;
        entry.next = bucket - 1;
        entry.str = stringValue;

        bucket = index + 1; // bucket is an int ref

        return stringValue;
    }

    Entry[] Resize()
    {
        var newSize = GetSize(this.count + 1);

        var entries = new Entry[newSize];

        int count = this.count;
        Array.Copy(this.entries, entries, count);

        this.buckets = new int[newSize];

        for (int i = 0; i < count; i++)
        {
            if (entries[i].next >= -1)
            {
                ref int bucket = ref GetBucket(entries[i].hashCode);
                entries[i].next = bucket - 1;
                bucket = i + 1;
            }
        }

        this.entries = entries;
        return entries;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    ref int GetBucket(uint hashCode)
    {
        int[] buckets = this.buckets;
        return ref buckets[hashCode & ((uint)buckets.Length - 1)];
    }

    struct Entry
    {
        public uint hashCode;
        public int next;
        public int count;
        public string str;
    }
}