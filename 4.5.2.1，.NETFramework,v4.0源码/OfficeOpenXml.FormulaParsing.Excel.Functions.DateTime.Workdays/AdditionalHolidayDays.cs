using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;

public class AdditionalHolidayDays
{
	private readonly FunctionArgument _holidayArg;

	private readonly List<System.DateTime> _holidayDates = new List<System.DateTime>();

	public IEnumerable<System.DateTime> AdditionalDates => _holidayDates;

	public AdditionalHolidayDays(FunctionArgument holidayArg)
	{
		_holidayArg = holidayArg;
		Initialize();
	}

	private void Initialize()
	{
		if (_holidayArg.Value is IEnumerable<FunctionArgument> source)
		{
			foreach (System.DateTime item in from arg in source
				where ConvertUtil.IsNumeric(arg.Value)
				select ConvertUtil.GetValueDouble(arg.Value) into dateSerial
				select System.DateTime.FromOADate(dateSerial))
			{
				_holidayDates.Add(item);
			}
		}
		if (_holidayArg.Value is ExcelDataProvider.IRangeInfo source2)
		{
			foreach (System.DateTime item2 in from cell in source2
				where ConvertUtil.IsNumeric(cell.Value)
				select ConvertUtil.GetValueDouble(cell.Value) into dateSerial
				select System.DateTime.FromOADate(dateSerial))
			{
				_holidayDates.Add(item2);
			}
		}
		if (ConvertUtil.IsNumeric(_holidayArg.Value))
		{
			_holidayDates.Add(System.DateTime.FromOADate(ConvertUtil.GetValueDouble(_holidayArg.Value)));
		}
	}
}
