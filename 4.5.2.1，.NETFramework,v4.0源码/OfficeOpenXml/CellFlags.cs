using System;

namespace OfficeOpenXml;

[Flags]
internal enum CellFlags
{
	RichText = 2,
	SharedFormula = 4,
	ArrayFormula = 8
}
