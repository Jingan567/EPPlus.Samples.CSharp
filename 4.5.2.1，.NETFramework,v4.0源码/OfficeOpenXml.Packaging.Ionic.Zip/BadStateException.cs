using System;

namespace OfficeOpenXml.Packaging.Ionic.Zip;

public class BadStateException : ZipException
{
	public BadStateException()
	{
	}

	public BadStateException(string message)
		: base(message)
	{
	}

	public BadStateException(string message, Exception innerException)
		: base(message, innerException)
	{
	}
}
