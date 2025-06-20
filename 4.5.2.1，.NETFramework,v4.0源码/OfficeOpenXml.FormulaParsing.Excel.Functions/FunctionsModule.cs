using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public abstract class FunctionsModule : IFunctionModule
{
	private readonly Dictionary<string, ExcelFunction> _functions = new Dictionary<string, ExcelFunction>();

	private readonly Dictionary<Type, FunctionCompiler> _customCompilers = new Dictionary<Type, FunctionCompiler>();

	public IDictionary<string, ExcelFunction> Functions => _functions;

	public IDictionary<Type, FunctionCompiler> CustomCompilers => _customCompilers;
}
