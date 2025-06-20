namespace OfficeOpenXml.FormulaParsing.ExcelUtilities;

public class RangeAddress
{
	private static RangeAddress _empty = new RangeAddress();

	internal string Address { get; set; }

	public string Worksheet { get; internal set; }

	public int FromCol { get; internal set; }

	public int ToCol { get; internal set; }

	public int FromRow { get; internal set; }

	public int ToRow { get; internal set; }

	public static RangeAddress Empty => _empty;

	public RangeAddress()
	{
		Address = string.Empty;
	}

	public override string ToString()
	{
		return Address;
	}

	public bool CollidesWith(RangeAddress other)
	{
		if (other.Worksheet != Worksheet)
		{
			return false;
		}
		if (other.FromRow > ToRow || other.FromCol > ToCol || FromRow > other.ToRow || FromCol > other.ToCol)
		{
			return false;
		}
		return true;
	}
}
