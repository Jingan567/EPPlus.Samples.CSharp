using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

public abstract class DateParsingFunction : ExcelFunction
{
	protected System.DateTime ParseDate(IEnumerable<FunctionArgument> arguments, object dateObj)
	{
		System.DateTime minValue = System.DateTime.MinValue;
		if (dateObj is string)
		{
			return System.DateTime.Parse(dateObj.ToString(), CultureInfo.InvariantCulture);
		}
		return System.DateTime.FromOADate(ArgToDecimal(arguments, 0));
	}
}
