using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing;

public class ParsingContext : IParsingLifetimeEventHandler
{
	public FormulaParser Parser { get; set; }

	public ExcelDataProvider ExcelDataProvider { get; set; }

	public RangeAddressFactory RangeAddressFactory { get; set; }

	public INameValueProvider NameValueProvider { get; set; }

	public ParsingConfiguration Configuration { get; set; }

	public ParsingScopes Scopes { get; private set; }

	public bool Debug => Configuration.Logger != null;

	private ParsingContext()
	{
	}

	public static ParsingContext Create()
	{
		ParsingContext obj = new ParsingContext
		{
			Configuration = ParsingConfiguration.Create()
		};
		obj.Scopes = new ParsingScopes(obj);
		return obj;
	}

	void IParsingLifetimeEventHandler.ParsingCompleted()
	{
	}
}
