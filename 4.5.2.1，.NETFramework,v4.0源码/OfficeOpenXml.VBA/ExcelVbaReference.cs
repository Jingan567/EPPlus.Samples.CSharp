namespace OfficeOpenXml.VBA;

public class ExcelVbaReference
{
	public int ReferenceRecordID { get; internal set; }

	public string Name { get; set; }

	public string Libid { get; set; }

	public ExcelVbaReference()
	{
		ReferenceRecordID = 13;
	}

	public override string ToString()
	{
		return Name;
	}
}
