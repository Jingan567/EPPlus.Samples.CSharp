namespace OfficeOpenXml.Packaging.Ionic.Zip;

public class BadCrcException : ZipException
{
	public BadCrcException()
	{
	}

	public BadCrcException(string message)
		: base(message)
	{
	}
}
