namespace OfficeOpenXml.Style;

public sealed class ExcelNumberFormat : StyleBase
{
	public int NumFmtID => base.Index;

	public string Format
	{
		get
		{
			for (int i = 0; i < _styles.NumberFormats.Count; i++)
			{
				if (base.Index == _styles.NumberFormats[i].NumFmtId)
				{
					return _styles.NumberFormats[i].Format;
				}
			}
			return "general";
		}
		set
		{
			_ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Numberformat, eStyleProperty.Format, string.IsNullOrEmpty(value) ? "General" : value, _positionID, _address));
		}
	}

	internal override string Id => Format;

	public bool BuildIn { get; private set; }

	internal ExcelNumberFormat(ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string Address, int index)
		: base(styles, ChangedEvent, PositionID, Address)
	{
		base.Index = index;
	}

	internal static string GetFromBuildInFromID(int _numFmtId)
	{
		return _numFmtId switch
		{
			0 => "General", 
			1 => "0", 
			2 => "0.00", 
			3 => "#,##0", 
			4 => "#,##0.00", 
			9 => "0%", 
			10 => "0.00%", 
			11 => "0.00E+00", 
			12 => "# ?/?", 
			13 => "# ??/??", 
			14 => "mm-dd-yy", 
			15 => "d-mmm-yy", 
			16 => "d-mmm", 
			17 => "mmm-yy", 
			18 => "h:mm AM/PM", 
			19 => "h:mm:ss AM/PM", 
			20 => "h:mm", 
			21 => "h:mm:ss", 
			22 => "m/d/yy h:mm", 
			37 => "#,##0 ;(#,##0)", 
			38 => "#,##0 ;[Red](#,##0)", 
			39 => "#,##0.00;(#,##0.00)", 
			40 => "#,##0.00;[Red](#,##0.00)", 
			45 => "mm:ss", 
			46 => "[h]:mm:ss", 
			47 => "mmss.0", 
			48 => "##0.0", 
			49 => "@", 
			_ => string.Empty, 
		};
	}

	internal static int GetFromBuildIdFromFormat(string format)
	{
		switch (format)
		{
		case "General":
		case "":
			return 0;
		case "0":
			return 1;
		case "0.00":
			return 2;
		case "#,##0":
			return 3;
		case "#,##0.00":
			return 4;
		case "0%":
			return 9;
		case "0.00%":
			return 10;
		case "0.00E+00":
			return 11;
		case "# ?/?":
			return 12;
		case "# ??/??":
			return 13;
		case "mm-dd-yy":
			return 14;
		case "d-mmm-yy":
			return 15;
		case "d-mmm":
			return 16;
		case "mmm-yy":
			return 17;
		case "h:mm AM/PM":
			return 18;
		case "h:mm:ss AM/PM":
			return 19;
		case "h:mm":
			return 20;
		case "h:mm:ss":
			return 21;
		case "m/d/yy h:mm":
			return 22;
		case "#,##0 ;(#,##0)":
			return 37;
		case "#,##0 ;[Red](#,##0)":
			return 38;
		case "#,##0.00;(#,##0.00)":
			return 39;
		case "#,##0.00;[Red](#,##0.00)":
			return 40;
		case "mm:ss":
			return 45;
		case "[h]:mm:ss":
			return 46;
		case "mmss.0":
			return 47;
		case "##0.0":
			return 48;
		case "@":
			return 49;
		default:
			return int.MinValue;
		}
	}
}
