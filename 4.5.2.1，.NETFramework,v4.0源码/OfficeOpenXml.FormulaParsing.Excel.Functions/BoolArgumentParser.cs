using System;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class BoolArgumentParser : ArgumentParser
{
	public override object Parse(object obj)
	{
		if (obj is ExcelDataProvider.IRangeInfo)
		{
			obj = ((ExcelDataProvider.IRangeInfo)obj).FirstOrDefault()?.Value;
		}
		if (obj == null)
		{
			return false;
		}
		if (obj is bool)
		{
			return (bool)obj;
		}
		if (obj.IsNumeric())
		{
			return Convert.ToBoolean(obj);
		}
		bool.TryParse(obj.ToString(), out var result);
		return result;
	}
}
