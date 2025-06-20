using System;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class CompileResultFactory
{
	public virtual CompileResult Create(object obj)
	{
		if (obj is ExcelDataProvider.INameInfo)
		{
			obj = ((ExcelDataProvider.INameInfo)obj).Value;
		}
		if (obj is ExcelDataProvider.IRangeInfo)
		{
			obj = ((ExcelDataProvider.IRangeInfo)obj).GetOffset(0, 0);
		}
		if (obj == null)
		{
			return new CompileResult(null, DataType.Empty);
		}
		if (obj.GetType().Equals(typeof(string)))
		{
			return new CompileResult(obj, DataType.String);
		}
		if (obj.GetType().Equals(typeof(double)) || obj is decimal)
		{
			return new CompileResult(obj, DataType.Decimal);
		}
		if (obj.GetType().Equals(typeof(int)) || obj is long || obj is short)
		{
			return new CompileResult(obj, DataType.Integer);
		}
		if (obj.GetType().Equals(typeof(bool)))
		{
			return new CompileResult(obj, DataType.Boolean);
		}
		if (obj.GetType().Equals(typeof(ExcelErrorValue)))
		{
			return new CompileResult(obj, DataType.ExcelError);
		}
		if (obj.GetType().Equals(typeof(DateTime)))
		{
			return new CompileResult(((DateTime)obj).ToOADate(), DataType.Date);
		}
		throw new ArgumentException("Non supported type " + obj.GetType().FullName);
	}
}
