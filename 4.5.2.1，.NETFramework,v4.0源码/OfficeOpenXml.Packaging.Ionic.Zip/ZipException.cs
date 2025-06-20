using System;

namespace OfficeOpenXml.Packaging.Ionic.Zip;

public class ZipException : Exception
{
	public ZipException()
	{
	}

	public ZipException(string message)
		: base(message)
	{
	}

	public ZipException(string message, Exception innerException)
		: base(message, innerException)
	{
	}
}
