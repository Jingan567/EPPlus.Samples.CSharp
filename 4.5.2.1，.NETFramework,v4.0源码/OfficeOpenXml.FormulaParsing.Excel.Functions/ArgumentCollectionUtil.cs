using System;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class ArgumentCollectionUtil
{
	private readonly DoubleEnumerableArgConverter _doubleEnumerableArgConverter;

	private readonly ObjectEnumerableArgConverter _objectEnumerableArgConverter;

	public ArgumentCollectionUtil()
		: this(new DoubleEnumerableArgConverter(), new ObjectEnumerableArgConverter())
	{
	}

	public ArgumentCollectionUtil(DoubleEnumerableArgConverter doubleEnumerableArgConverter, ObjectEnumerableArgConverter objectEnumerableArgConverter)
	{
		_doubleEnumerableArgConverter = doubleEnumerableArgConverter;
		_objectEnumerableArgConverter = objectEnumerableArgConverter;
	}

	public virtual IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerable(bool ignoreHidden, bool ignoreErrors, IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		return _doubleEnumerableArgConverter.ConvertArgs(ignoreHidden, ignoreErrors, arguments, context);
	}

	public virtual IEnumerable<object> ArgsToObjectEnumerable(bool ignoreHidden, IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		return _objectEnumerableArgConverter.ConvertArgs(ignoreHidden, arguments, context);
	}

	public virtual double CalculateCollection(IEnumerable<FunctionArgument> collection, double result, Func<FunctionArgument, double, double> action)
	{
		foreach (FunctionArgument item in collection)
		{
			result = ((!(item.Value is IEnumerable<FunctionArgument>)) ? action(item, result) : CalculateCollection((IEnumerable<FunctionArgument>)item.Value, result, action));
		}
		return result;
	}
}
