using System;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.Exceptions;

public class UnrecognizedTokenException : Exception
{
	public UnrecognizedTokenException(Token token)
		: base("Unrecognized token: " + token.Value)
	{
	}
}
