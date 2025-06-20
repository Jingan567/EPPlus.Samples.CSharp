using System;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class ArgumentParserFactory
{
	public virtual ArgumentParser CreateArgumentParser(DataType dataType)
	{
		return dataType switch
		{
			DataType.Integer => new IntArgumentParser(), 
			DataType.Boolean => new BoolArgumentParser(), 
			DataType.Decimal => new DoubleArgumentParser(), 
			_ => throw new InvalidOperationException("non supported argument parser type " + dataType), 
		};
	}
}
