using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public interface IFunctionModule
{
	IDictionary<string, ExcelFunction> Functions { get; }

	IDictionary<Type, FunctionCompiler> CustomCompilers { get; }
}
