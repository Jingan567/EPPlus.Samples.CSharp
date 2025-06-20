using System;

namespace OfficeOpenXml.Style;

internal class StyleChangeEventArgs : EventArgs
{
	internal eStyleClass StyleClass;

	internal eStyleProperty StyleProperty;

	internal object Value;

	internal string Address;

	internal int PositionID { get; set; }

	internal StyleChangeEventArgs(eStyleClass styleclass, eStyleProperty styleProperty, object value, int positionID, string address)
	{
		StyleClass = styleclass;
		StyleProperty = styleProperty;
		Value = value;
		Address = address;
		PositionID = positionID;
	}
}
