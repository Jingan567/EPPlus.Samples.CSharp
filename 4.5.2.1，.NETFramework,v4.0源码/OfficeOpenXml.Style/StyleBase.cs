namespace OfficeOpenXml.Style;

public abstract class StyleBase
{
	protected ExcelStyles _styles;

	internal XmlHelper.ChangedEventHandler _ChangedEvent;

	protected int _positionID;

	protected string _address;

	internal int Index { get; set; }

	internal abstract string Id { get; }

	internal StyleBase(ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string Address)
	{
		_styles = styles;
		_ChangedEvent = ChangedEvent;
		_address = Address;
		_positionID = PositionID;
	}

	internal virtual void SetIndex(int index)
	{
		Index = index;
	}
}
