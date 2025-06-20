using System;

namespace OfficeOpenXml.VBA;

public class ExcelVbaReferenceControl : ExcelVbaReference
{
	public string LibIdExternal { get; set; }

	public string LibIdTwiddled { get; set; }

	public Guid OriginalTypeLib { get; set; }

	internal uint Cookie { get; set; }

	public ExcelVbaReferenceControl()
	{
		base.ReferenceRecordID = 47;
	}
}
