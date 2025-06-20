using System;

namespace OfficeOpenXml.Packaging.Ionic.Zip;

public class BadPasswordException : ZipException
{
	public BadPasswordException()
	{
	}

	public BadPasswordException(string message)
		: base(message)
	{
	}

	public BadPasswordException(string message, Exception innerException)
		: base(message, innerException)
	{
	}
}
