using System;
using System.Collections;
using System.Collections.Generic;

namespace ExcelFlow.SourceGenerators;

/// <summary>
/// An immutable array with value (sequence) equality semantics.
/// Required for incremental generator models: the pipeline caches by Equals, so models
/// must compare by content, not by reference.
/// </summary>
internal readonly struct EquatableArray<T> : IEquatable<EquatableArray<T>>, IEnumerable<T>
    where T : IEquatable<T>
{
    public static readonly EquatableArray<T> Empty = new EquatableArray<T>(Array.Empty<T>());

    private readonly T[]? _array;

    public EquatableArray(T[] array) => _array = array;

    public int Count => _array?.Length ?? 0;

    public T this[int index] => _array![index];

    public bool Equals(EquatableArray<T> other)
    {
        if (ReferenceEquals(_array, other._array))
            return true;

        if (_array is null || other._array is null)
            return false;

        if (_array.Length != other._array.Length)
            return false;

        for (int i = 0; i < _array.Length; i++)
        {
            if (!_array[i].Equals(other._array[i]))
                return false;
        }

        return true;
    }

    public override bool Equals(object? obj) => obj is EquatableArray<T> other && Equals(other);

    public override int GetHashCode()
    {
        if (_array is null)
            return 0;

        unchecked
        {
            int hash = 17;

            foreach (T item in _array)
                hash = hash * 31 + (item is null ? 0 : item.GetHashCode());

            return hash;
        }
    }

    public IEnumerator<T> GetEnumerator() => ((IEnumerable<T>)(_array ?? Array.Empty<T>())).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
