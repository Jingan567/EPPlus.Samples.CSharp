using System;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class DoubleArgumentParser : ArgumentParser
{
	public override object Parse(object obj)
	{
		OfficeOpenXml.FormulaParsing.Utilities.Require.That(obj).Named("argument").IsNotNull();
		if (obj is ExcelDataProvider.IRangeInfo)
		{
			return ((ExcelDataProvider.IRangeInfo)obj).FirstOrDefault()?.ValueDouble ?? 0.0;
		}
		if (obj is double)
		{
			return obj;
		}
		if (obj.IsNumeric())
		{
			return ConvertUtil.GetValueDouble(obj);
		}
		string s = ((obj != null) ? obj.ToString() : string.Empty);
		try
		{
			if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
			{
				return result;
			}
			return System.DateTime.Parse(s, CultureInfo.CurrentCulture, DateTimeStyles.None).ToOADate();
		}
		catch
		{
			throw new ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Value));
		}
	}
}
