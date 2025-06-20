namespace OfficeOpenXml.Packaging.Ionic.Zip;

public class SfxGenerationException : ZipException
{
	public SfxGenerationException()
	{
	}

	public SfxGenerationException(string message)
		: base(message)
	{
	}
}
