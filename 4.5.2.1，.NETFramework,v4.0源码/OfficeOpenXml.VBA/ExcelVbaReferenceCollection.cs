namespace OfficeOpenXml.VBA;

public class ExcelVbaReferenceCollection : ExcelVBACollectionBase<ExcelVbaReference>
{
	internal ExcelVbaReferenceCollection()
	{
	}

	public void Add(ExcelVbaReference Item)
	{
		_list.Add(Item);
	}
}
