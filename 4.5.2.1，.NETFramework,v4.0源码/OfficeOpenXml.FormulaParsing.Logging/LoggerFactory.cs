using System.IO;

namespace OfficeOpenXml.FormulaParsing.Logging;

public static class LoggerFactory
{
	public static IFormulaParserLogger CreateTextFileLogger(FileInfo file)
	{
		return new TextFileLogger(file);
	}
}
