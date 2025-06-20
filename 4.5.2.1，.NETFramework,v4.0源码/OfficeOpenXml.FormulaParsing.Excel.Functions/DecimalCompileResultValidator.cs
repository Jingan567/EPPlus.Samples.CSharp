using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class DecimalCompileResultValidator : CompileResultValidator
{
	public override void Validate(object obj)
	{
		double valueDouble = ConvertUtil.GetValueDouble(obj);
		if (double.IsNaN(valueDouble) || double.IsInfinity(valueDouble))
		{
			throw new ExcelErrorValueException(eErrorType.Num);
		}
	}
}
