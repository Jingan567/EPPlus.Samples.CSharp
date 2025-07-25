using System;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

public static class LookupNavigatorFactory
{
	public static LookupNavigator Create(LookupDirection direction, LookupArguments args, ParsingContext parsingContext)
	{
		if (args.ArgumentDataType == LookupArguments.LookupArgumentDataType.ExcelRange)
		{
			return new ExcelLookupNavigator(direction, args, parsingContext);
		}
		if (args.ArgumentDataType == LookupArguments.LookupArgumentDataType.DataArray)
		{
			return new ArrayLookupNavigator(direction, args, parsingContext);
		}
		throw new NotSupportedException("Invalid argument datatype");
	}
}
