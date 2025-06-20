using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public abstract class LookupFunction : ExcelFunction
{
	private readonly ValueMatcher _valueMatcher;

	private readonly CompileResultFactory _compileResultFactory;

	public override bool IsLookupFuction => true;

	public LookupFunction()
		: this(new LookupValueMatcher(), new CompileResultFactory())
	{
	}

	public LookupFunction(ValueMatcher valueMatcher, CompileResultFactory compileResultFactory)
	{
		_valueMatcher = valueMatcher;
		_compileResultFactory = compileResultFactory;
	}

	protected int IsMatch(object o1, object o2)
	{
		return _valueMatcher.IsMatch(o1, o2);
	}

	protected LookupDirection GetLookupDirection(RangeAddress rangeAddress)
	{
		int num = rangeAddress.ToRow - rangeAddress.FromRow;
		if (rangeAddress.ToCol - rangeAddress.FromCol <= num)
		{
			return LookupDirection.Vertical;
		}
		return LookupDirection.Horizontal;
	}

	protected CompileResult Lookup(LookupNavigator navigator, LookupArguments lookupArgs)
	{
		object obj = null;
		object obj2 = null;
		int? num = null;
		if (lookupArgs.SearchedValue == null)
		{
			return new CompileResult(eErrorType.NA);
		}
		do
		{
			int num2 = IsMatch(navigator.CurrentValue, lookupArgs.SearchedValue);
			if (num2 != 0)
			{
				if (obj != null && navigator.CurrentValue == null)
				{
					break;
				}
				if (lookupArgs.RangeLookup)
				{
					if (obj == null && num2 > 0)
					{
						return new CompileResult(eErrorType.NA);
					}
					if (obj != null && num2 > 0 && num < 0)
					{
						return _compileResultFactory.Create(obj2);
					}
					num = num2;
					obj = navigator.CurrentValue;
					obj2 = navigator.GetLookupValue();
				}
				continue;
			}
			return _compileResultFactory.Create(navigator.GetLookupValue());
		}
		while (navigator.MoveNext());
		if (!lookupArgs.RangeLookup)
		{
			return new CompileResult(eErrorType.NA);
		}
		return _compileResultFactory.Create(obj2);
	}
}
