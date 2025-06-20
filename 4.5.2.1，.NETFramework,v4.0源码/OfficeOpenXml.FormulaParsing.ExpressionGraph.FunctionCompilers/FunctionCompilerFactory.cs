using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

public class FunctionCompilerFactory
{
	private readonly Dictionary<Type, FunctionCompiler> _specialCompilers = new Dictionary<Type, FunctionCompiler>();

	public FunctionCompilerFactory(FunctionRepository repository)
	{
		_specialCompilers.Add(typeof(If), new IfFunctionCompiler(repository.GetFunction("if")));
		_specialCompilers.Add(typeof(IfError), new IfErrorFunctionCompiler(repository.GetFunction("iferror")));
		_specialCompilers.Add(typeof(IfNa), new IfNaFunctionCompiler(repository.GetFunction("ifna")));
		foreach (Type key in repository.CustomCompilers.Keys)
		{
			_specialCompilers.Add(key, repository.CustomCompilers[key]);
		}
	}

	private FunctionCompiler GetCompilerByType(ExcelFunction function)
	{
		Type type = function.GetType();
		if (_specialCompilers.ContainsKey(type))
		{
			return _specialCompilers[type];
		}
		return new DefaultCompiler(function);
	}

	public virtual FunctionCompiler Create(ExcelFunction function)
	{
		if (function.IsLookupFuction)
		{
			return new LookupFunctionCompiler(function);
		}
		if (function.IsErrorHandlingFunction)
		{
			return new ErrorHandlingFunctionCompiler(function);
		}
		return GetCompilerByType(function);
	}
}
