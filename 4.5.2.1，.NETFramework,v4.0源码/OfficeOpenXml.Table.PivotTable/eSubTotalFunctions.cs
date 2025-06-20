using System;

namespace OfficeOpenXml.Table.PivotTable;

[Flags]
public enum eSubTotalFunctions
{
	None = 1,
	Count = 2,
	CountA = 4,
	Avg = 8,
	Default = 0x10,
	Min = 0x20,
	Max = 0x40,
	Product = 0x80,
	StdDev = 0x100,
	StdDevP = 0x200,
	Sum = 0x400,
	Var = 0x800,
	VarP = 0x1000
}
