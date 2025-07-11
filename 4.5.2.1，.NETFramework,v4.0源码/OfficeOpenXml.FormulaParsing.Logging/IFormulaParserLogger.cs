using System;

namespace OfficeOpenXml.FormulaParsing.Logging;

public interface IFormulaParserLogger : IDisposable
{
	void Log(ParsingContext context, Exception ex);

	void Log(ParsingContext context, string message);

	void Log(string message);

	void LogCellCounted();

	void LogFunction(string func);

	void LogFunction(string func, long milliseconds);
}
