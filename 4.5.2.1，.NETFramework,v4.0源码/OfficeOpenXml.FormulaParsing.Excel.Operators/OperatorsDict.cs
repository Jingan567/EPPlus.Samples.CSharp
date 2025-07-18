using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Operators;

public class OperatorsDict : Dictionary<string, IOperator>
{
	private static IDictionary<string, IOperator> _instance;

	public static IDictionary<string, IOperator> Instance
	{
		get
		{
			if (_instance == null)
			{
				_instance = new OperatorsDict();
			}
			return _instance;
		}
	}

	public OperatorsDict()
	{
		Add("+", Operator.Plus);
		Add("-", Operator.Minus);
		Add("*", Operator.Multiply);
		Add("/", Operator.Divide);
		Add("^", Operator.Exp);
		Add("=", Operator.Eq);
		Add(">", Operator.GreaterThan);
		Add(">=", Operator.GreaterThanOrEqual);
		Add("<", Operator.LessThan);
		Add("<=", Operator.LessThanOrEqual);
		Add("<>", Operator.NotEqualsTo);
		Add("&", Operator.Concat);
	}
}
