using System;

namespace OfficeOpenXml.FormulaParsing.Exceptions;

public class CircularReferenceException : Exception
{
	public CircularReferenceException(string message)
		: base(message)
	{
	}
}
