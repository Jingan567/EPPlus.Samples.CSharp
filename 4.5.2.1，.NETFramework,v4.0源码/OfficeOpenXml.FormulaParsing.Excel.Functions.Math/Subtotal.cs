using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

public class Subtotal : ExcelFunction
{
	private Dictionary<int, HiddenValuesHandlingFunction> _functions = new Dictionary<int, HiddenValuesHandlingFunction>();

	public Subtotal()
	{
		Initialize();
	}

	private void Initialize()
	{
		_functions[1] = new Average();
		_functions[2] = new Count();
		_functions[3] = new CountA();
		_functions[4] = new Max();
		_functions[5] = new Min();
		_functions[6] = new Product();
		_functions[7] = new Stdev();
		_functions[8] = new StdevP();
		_functions[9] = new Sum();
		_functions[10] = new Var();
		_functions[11] = new VarP();
		AddHiddenValueHandlingFunction(new Average(), 101);
		AddHiddenValueHandlingFunction(new Count(), 102);
		AddHiddenValueHandlingFunction(new CountA(), 103);
		AddHiddenValueHandlingFunction(new Max(), 104);
		AddHiddenValueHandlingFunction(new Min(), 105);
		AddHiddenValueHandlingFunction(new Product(), 106);
		AddHiddenValueHandlingFunction(new Stdev(), 107);
		AddHiddenValueHandlingFunction(new StdevP(), 108);
		AddHiddenValueHandlingFunction(new Sum(), 109);
		AddHiddenValueHandlingFunction(new Var(), 110);
		AddHiddenValueHandlingFunction(new VarP(), 111);
	}

	private void AddHiddenValueHandlingFunction(HiddenValuesHandlingFunction func, int funcNum)
	{
		func.IgnoreHiddenValues = true;
		_functions[funcNum] = func;
	}

	public override void BeforeInvoke(ParsingContext context)
	{
		base.BeforeInvoke(context);
		context.Scopes.Current.IsSubtotal = true;
	}

	public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
	{
		ValidateArguments(arguments, 2);
		int funcNum = ArgToInt(arguments, 0);
		if (context.Scopes.Current.Parent != null && context.Scopes.Current.Parent.IsSubtotal)
		{
			return CreateResult(0.0, DataType.Decimal);
		}
		IEnumerable<FunctionArgument> arguments2 = arguments.Skip(1);
		CompileResult compileResult = GetFunctionByCalcType(funcNum).Execute(arguments2, context);
		compileResult.IsResultOfSubtotal = true;
		return compileResult;
	}

	private ExcelFunction GetFunctionByCalcType(int funcNum)
	{
		if (!_functions.ContainsKey(funcNum))
		{
			ThrowExcelErrorValueException(eErrorType.Value);
		}
		return _functions[funcNum];
	}
}
