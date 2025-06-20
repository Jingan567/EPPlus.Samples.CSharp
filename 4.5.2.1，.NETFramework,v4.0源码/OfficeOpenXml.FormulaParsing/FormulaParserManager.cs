using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing;

public class FormulaParserManager
{
	private readonly FormulaParser _parser;

	internal FormulaParserManager(FormulaParser parser)
	{
		Require.That(parser).Named("parser").IsNotNull();
		_parser = parser;
	}

	public void LoadFunctionModule(IFunctionModule module)
	{
		_parser.Configure(delegate(ParsingConfiguration x)
		{
			x.FunctionRepository.LoadModule(module);
		});
	}

	public void AddOrReplaceFunction(string functionName, ExcelFunction functionImpl)
	{
		_parser.Configure(delegate(ParsingConfiguration x)
		{
			x.FunctionRepository.AddOrReplaceFunction(functionName, functionImpl);
		});
	}

	public void CopyFunctionsFrom(ExcelWorkbook otherWorkbook)
	{
		foreach (KeyValuePair<string, ExcelFunction> implementedFunction in otherWorkbook.FormulaParserManager.GetImplementedFunctions())
		{
			AddOrReplaceFunction(implementedFunction.Key, implementedFunction.Value);
		}
	}

	public IEnumerable<string> GetImplementedFunctionNames()
	{
		List<string> list = _parser.FunctionNames.ToList();
		list.Sort((string x, string y) => string.Compare(x, y, StringComparison.Ordinal));
		return list;
	}

	public IEnumerable<KeyValuePair<string, ExcelFunction>> GetImplementedFunctions()
	{
		List<KeyValuePair<string, ExcelFunction>> functions = new List<KeyValuePair<string, ExcelFunction>>();
		_parser.Configure(delegate(ParsingConfiguration parsingConfiguration)
		{
			foreach (string functionName in parsingConfiguration.FunctionRepository.FunctionNames)
			{
				functions.Add(new KeyValuePair<string, ExcelFunction>(functionName, parsingConfiguration.FunctionRepository.GetFunction(functionName)));
			}
		});
		return functions;
	}

	public object Parse(string formula)
	{
		return _parser.Parse(formula);
	}

	public void AttachLogger(IFormulaParserLogger logger)
	{
		_parser.Configure(delegate(ParsingConfiguration c)
		{
			c.AttachLogger(logger);
		});
	}

	public void AttachLogger(FileInfo logfile)
	{
		_parser.Configure(delegate(ParsingConfiguration c)
		{
			c.AttachLogger(LoggerFactory.CreateTextFileLogger(logfile));
		});
	}

	public void DetachLogger()
	{
		_parser.Configure(delegate(ParsingConfiguration c)
		{
			c.DetachLogger();
		});
	}
}
