using System;

namespace OfficeOpenXml.FormulaParsing.Excel;

[Flags]
public enum ExcelCellState
{
	HiddenCell = 1,
	ContainsError = 2,
	IsResultOfSubtotal = 4
}
