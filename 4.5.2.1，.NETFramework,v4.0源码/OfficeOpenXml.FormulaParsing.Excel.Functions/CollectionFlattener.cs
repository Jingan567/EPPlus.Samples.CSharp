using System;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public abstract class CollectionFlattener<T>
{
	public virtual IEnumerable<T> FuncArgsToFlatEnumerable(IEnumerable<FunctionArgument> arguments, Action<FunctionArgument, IList<T>> convertFunc)
	{
		List<T> list = new List<T>();
		FuncArgsToFlatEnumerable(arguments, list, convertFunc);
		return list;
	}

	private void FuncArgsToFlatEnumerable(IEnumerable<FunctionArgument> arguments, List<T> argList, Action<FunctionArgument, IList<T>> convertFunc)
	{
		foreach (FunctionArgument argument in arguments)
		{
			if (argument.Value is IEnumerable<FunctionArgument>)
			{
				FuncArgsToFlatEnumerable((IEnumerable<FunctionArgument>)argument.Value, argList, convertFunc);
			}
			else
			{
				convertFunc(argument, argList);
			}
		}
	}
}
