using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class ExcelAddressExpression : AtomicExpression
{
	private readonly ExcelDataProvider _excelDataProvider;

	private readonly ParsingContext _parsingContext;

	private readonly RangeAddressFactory _rangeAddressFactory;

	private readonly bool _negate;

	public bool ResolveAsRange { get; set; }

	public override bool IsGroupedExpression => false;

	public ExcelAddressExpression(string expression, ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
		: this(expression, excelDataProvider, parsingContext, new RangeAddressFactory(excelDataProvider), negate: false)
	{
	}

	public ExcelAddressExpression(string expression, ExcelDataProvider excelDataProvider, ParsingContext parsingContext, bool negate)
		: this(expression, excelDataProvider, parsingContext, new RangeAddressFactory(excelDataProvider), negate)
	{
	}

	public ExcelAddressExpression(string expression, ExcelDataProvider excelDataProvider, ParsingContext parsingContext, RangeAddressFactory rangeAddressFactory, bool negate)
		: base(expression)
	{
		Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
		Require.That(parsingContext).Named("parsingContext").IsNotNull();
		Require.That(rangeAddressFactory).Named("rangeAddressFactory").IsNotNull();
		_excelDataProvider = excelDataProvider;
		_parsingContext = parsingContext;
		_rangeAddressFactory = rangeAddressFactory;
		_negate = negate;
	}

	public override CompileResult Compile()
	{
		if (ParentIsLookupFunction)
		{
			return new CompileResult(base.ExpressionString, DataType.ExcelAddress);
		}
		return CompileRangeValues();
	}

	private CompileResult CompileRangeValues()
	{
		ParsingScope current = _parsingContext.Scopes.Current;
		ExcelDataProvider.IRangeInfo range = _excelDataProvider.GetRange(current.Address.Worksheet, current.Address.FromRow, current.Address.FromCol, base.ExpressionString);
		if (range == null)
		{
			return CompileResult.Empty;
		}
		if (ResolveAsRange || range.Address.Rows > 1 || range.Address.Columns > 1)
		{
			return new CompileResult(range, DataType.Enumerable);
		}
		return CompileSingleCell(range);
	}

	private CompileResult CompileSingleCell(ExcelDataProvider.IRangeInfo result)
	{
		ExcelDataProvider.ICellInfo cellInfo = result.FirstOrDefault();
		if (cellInfo == null)
		{
			return CompileResult.Empty;
		}
		CompileResult compileResult = new CompileResultFactory().Create(cellInfo.Value);
		if (_negate && compileResult.IsNumeric)
		{
			compileResult = new CompileResult(compileResult.ResultNumeric * -1.0, compileResult.DataType);
		}
		compileResult.IsHiddenCell = cellInfo.IsHiddenRow;
		return compileResult;
	}
}
