using System;

namespace OfficeOpenXml.Table.PivotTable;

[Flags]
public enum eDateGroupBy
{
	Years = 1,
	Quarters = 2,
	Months = 4,
	Days = 8,
	Hours = 0x10,
	Minutes = 0x20,
	Seconds = 0x40
}
