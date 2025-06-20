using System;
using System.Collections.Generic;
using System.Globalization;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class FunctionRepository : IFunctionNameProvider
{
	private Dictionary<Type, FunctionCompiler> _customCompilers = new Dictionary<Type, FunctionCompiler>();

	private Dictionary<string, ExcelFunction> _functions = new Dictionary<string, ExcelFunction>(StringComparer.Ordinal);

	public Dictionary<Type, FunctionCompiler> CustomCompilers => _customCompilers;

	public IEnumerable<string> FunctionNames => _functions.Keys;

	private FunctionRepository()
	{
	}

	public static FunctionRepository Create()
	{
		FunctionRepository functionRepository = new FunctionRepository();
		functionRepository.LoadModule(new BuiltInFunctions());
		return functionRepository;
	}

	public virtual void LoadModule(IFunctionModule module)
	{
		foreach (string key2 in module.Functions.Keys)
		{
			string key = key2.ToLower(CultureInfo.InvariantCulture);
			_functions[key] = module.Functions[key2];
		}
		foreach (Type key3 in module.CustomCompilers.Keys)
		{
			CustomCompilers[key3] = module.CustomCompilers[key3];
		}
	}

	public virtual ExcelFunction GetFunction(string name)
	{
		if (!_functions.ContainsKey(name.ToLower(CultureInfo.InvariantCulture)))
		{
			return null;
		}
		return _functions[name.ToLower(CultureInfo.InvariantCulture)];
	}

	public virtual void Clear()
	{
		_functions.Clear();
	}

	public bool IsFunctionName(string name)
	{
		return _functions.ContainsKey(name.ToLower(CultureInfo.InvariantCulture));
	}

	public void AddOrReplaceFunction(string functionName, ExcelFunction functionImpl)
	{
		Require.That(functionName).Named("functionName").IsNotNullOrEmpty();
		Require.That(functionImpl).Named("functionImpl").IsNotNull();
		string key = functionName.ToLower(CultureInfo.InvariantCulture);
		if (_functions.ContainsKey(key))
		{
			_functions.Remove(key);
		}
		_functions[key] = functionImpl;
	}
}
