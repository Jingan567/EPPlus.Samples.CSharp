using System;

namespace OfficeOpenXml.FormulaParsing.Exceptions;

public class ExcelErrorValueException : Exception
{
	public ExcelErrorValue ErrorValue { get; private set; }

	public ExcelErrorValueException(ExcelErrorValue error)
		: this(error.ToString(), error)
	{
	}

	public ExcelErrorValueException(string message, ExcelErrorValue error)
		: base(message)
	{
		ErrorValue = error;
	}

	public ExcelErrorValueException(eErrorType errorType)
		: this(ExcelErrorValue.Create(errorType))
	{
	}
}
