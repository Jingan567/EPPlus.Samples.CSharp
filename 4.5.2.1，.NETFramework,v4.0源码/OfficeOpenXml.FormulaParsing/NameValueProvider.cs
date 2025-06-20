namespace OfficeOpenXml.FormulaParsing;

public class NameValueProvider : INameValueProvider
{
	public static INameValueProvider Empty => new NameValueProvider();

	private NameValueProvider()
	{
	}

	public bool IsNamedValue(string key, string worksheet)
	{
		return false;
	}

	public object GetNamedValue(string key)
	{
		return null;
	}

	public void Reload()
	{
	}
}
