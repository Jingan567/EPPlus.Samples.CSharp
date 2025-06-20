using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Mina : ExcelFunction
{
	private readonly DoubleEnumerableArgConverter _argConverter;

	public Mina()
		: this(new DoubleEnumerableArgConverter())
	{
	}

	public Mina(DoubleEnumerableArgConverter argConverter)
	{
		Require.That(argConverter).Named("argConverter").IsNotNull();
		_argConverter = argConverter;
	}

	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		IEnumerable<ExcelDoubleCellValue> source = _argConverter.ConvertArgsIncludingOtherTypes(arguments);
		return CreateResult(source.Min(), DataType.Decimal);
	}
}
