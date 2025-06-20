using System;

internal class IndexBase : IComparable<IndexBase>
{
	internal short Index;

	public int CompareTo(IndexBase other)
	{
		return Index - other.Index;
	}
}
