using System;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class ConstantExpression : AtomicExpression
{
	private readonly Func<CompileResult> _factoryMethod;

	public ConstantExpression(string title, Func<CompileResult> factoryMethod)
		: base(title)
	{
		_factoryMethod = factoryMethod;
	}

	public override CompileResult Compile()
	{
		return _factoryMethod();
	}
}
