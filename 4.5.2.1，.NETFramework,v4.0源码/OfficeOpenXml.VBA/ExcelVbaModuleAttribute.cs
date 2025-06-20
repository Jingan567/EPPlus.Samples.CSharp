namespace OfficeOpenXml.VBA;

public class ExcelVbaModuleAttribute
{
	public string Name { get; internal set; }

	public eAttributeDataType DataType { get; internal set; }

	public string Value { get; set; }

	internal ExcelVbaModuleAttribute()
	{
	}

	public override string ToString()
	{
		return Name;
	}
}
