namespace OfficeOpenXml.Utils;

internal class Argument<T> : IArgument<T>
{
	private T _value;

	T IArgument<T>.Value => _value;

	public Argument(T value)
	{
		_value = value;
	}
}
