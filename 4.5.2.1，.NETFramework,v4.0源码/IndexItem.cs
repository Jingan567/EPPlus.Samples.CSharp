using System;

internal struct IndexItem : IComparable<IndexItem>
{
	internal short Index;

	internal int IndexPointer { get; set; }

	public int CompareTo(IndexItem other)
	{
		return Index - other.Index;
	}
}
