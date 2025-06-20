using System;

namespace OfficeOpenXml.Packaging.Ionic.Zip;

public class BadReadException : ZipException
{
	public BadReadException()
	{
	}

	public BadReadException(string message)
		: base(message)
	{
	}

	public BadReadException(string message, Exception innerException)
		: base(message, innerException)
	{
	}
}
