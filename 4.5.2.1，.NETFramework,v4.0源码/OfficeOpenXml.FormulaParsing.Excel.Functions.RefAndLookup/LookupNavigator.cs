using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public abstract class LookupNavigator
{
	protected readonly LookupDirection Direction;

	protected readonly LookupArguments Arguments;

	protected readonly ParsingContext ParsingContext;

	public abstract int Index { get; }

	public abstract object CurrentValue { get; }

	public LookupNavigator(LookupDirection direction, LookupArguments arguments, ParsingContext parsingContext)
	{
		Require.That(arguments).Named("arguments").IsNotNull();
		Require.That(parsingContext).Named("parsingContext").IsNotNull();
		Require.That(parsingContext.ExcelDataProvider).Named("parsingContext.ExcelDataProvider").IsNotNull();
		Direction = direction;
		Arguments = arguments;
		ParsingContext = parsingContext;
	}

	public abstract bool MoveNext();

	public abstract object GetLookupValue();
}
