using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

public class Value : ExcelFunction
{
	private readonly string _groupSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;

	private readonly string _decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;

	private readonly string _timeSeparator = CultureInfo.CurrentCulture.DateTimeFormat.TimeSeparator;

	private readonly string _shortTimePattern = CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern;

	private readonly DateValue _dateValueFunc = new DateValue();

	private readonly TimeValue _timeValueFunc = new TimeValue();

	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 1);
		string text = ArgToString(arguments, 0).TrimEnd(' ');
		double result = 0.0;
		if (Regex.IsMatch(text, $"^[\\d]*({Regex.Escape(_groupSeparator)}?[\\d]*)?({Regex.Escape(_decimalSeparator)}[\\d]*)?[ ?% ?]?$"))
		{
			if (text.EndsWith("%"))
			{
				text = text.TrimEnd('%');
				result = double.Parse(text) / 100.0;
			}
			else
			{
				result = double.Parse(text);
			}
			return CreateResult(result, DataType.Decimal);
		}
		if (double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out result))
		{
			return CreateResult(result, DataType.Decimal);
		}
		string text2 = Regex.Escape(_timeSeparator);
		if (Regex.IsMatch(text, "^[\\d]{1,2}" + text2 + "[\\d]{2}(" + text2 + "[\\d]{2})?$"))
		{
			CompileResult compileResult = _timeValueFunc.Execute(text);
			if (compileResult.DataType == DataType.Date)
			{
				return compileResult;
			}
		}
		CompileResult compileResult2 = _dateValueFunc.Execute(text);
		if (compileResult2.DataType == DataType.Date)
		{
			return compileResult2;
		}
		return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
	}
}
