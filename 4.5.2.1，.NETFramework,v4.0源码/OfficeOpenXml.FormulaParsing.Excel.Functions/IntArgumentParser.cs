using System;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class IntArgumentParser : ArgumentParser
{
	public override object Parse(object obj)
	{
		Require.That(obj).Named("argument").IsNotNull();
		if (obj is ExcelDataProvider.IRangeInfo)
		{
			ExcelDataProvider.ICellInfo cellInfo = ((ExcelDataProvider.IRangeInfo)obj).FirstOrDefault();
			return (cellInfo != null) ? Convert.ToInt32(cellInfo.ValueDouble) : 0;
		}
		Type type = obj.GetType();
		if (type == typeof(int))
		{
			return (int)obj;
		}
		if (type == typeof(double) || type == typeof(decimal))
		{
			return Convert.ToInt32(obj);
		}
		if (!int.TryParse(obj.ToString(), out var result))
		{
			throw new ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Value));
		}
		return result;
	}
}
