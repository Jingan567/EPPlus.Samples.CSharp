namespace OfficeOpenXml.VBA;

public class ExcelVbaReferenceProject : ExcelVbaReference
{
	public string LibIdRelative { get; set; }

	public uint MajorVersion { get; set; }

	public ushort MinorVersion { get; set; }

	public ExcelVbaReferenceProject()
	{
		base.ReferenceRecordID = 14;
	}
}
