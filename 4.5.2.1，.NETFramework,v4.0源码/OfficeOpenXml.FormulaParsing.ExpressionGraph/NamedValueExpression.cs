using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class NamedValueExpression : AtomicExpression
{
	private readonly ParsingContext _parsingContext;

	public NamedValueExpression(string expression, ParsingContext parsingContext)
		: base(expression)
	{
		_parsingContext = parsingContext;
	}

	public override CompileResult Compile()
	{
		ParsingScope current = _parsingContext.Scopes.Current;
		ExcelDataProvider.INameInfo name = _parsingContext.ExcelDataProvider.GetName(current.Address.Worksheet, base.ExpressionString);
		if (name == null)
		{
			throw new ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Name));
		}
		if (name.Value == null)
		{
			return null;
		}
		if (name.Value is ExcelDataProvider.IRangeInfo)
		{
			ExcelDataProvider.IRangeInfo rangeInfo = (ExcelDataProvider.IRangeInfo)name.Value;
			if (rangeInfo.IsMulti)
			{
				return new CompileResult(name.Value, DataType.Enumerable);
			}
			if (rangeInfo.IsEmpty)
			{
				return null;
			}
			return new CompileResultFactory().Create(rangeInfo.First().Value);
		}
		return new CompileResultFactory().Create(name.Value);
	}
}
