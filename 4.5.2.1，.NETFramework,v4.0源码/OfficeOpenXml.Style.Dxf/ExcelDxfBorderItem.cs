namespace OfficeOpenXml.Style.Dxf;

public class ExcelDxfBorderItem : DxfStyleBase<ExcelDxfBorderItem>
{
	public ExcelBorderStyle? Style { get; set; }

	public ExcelDxfColor Color { get; internal set; }

	protected internal override string Id => GetAsString(Style) + "|" + ((Color == null) ? "" : Color.Id);

	protected internal override bool HasValue
	{
		get
		{
			if (!Style.HasValue)
			{
				return Color.HasValue;
			}
			return true;
		}
	}

	internal ExcelDxfBorderItem(ExcelStyles styles)
		: base(styles)
	{
		Color = new ExcelDxfColor(styles);
	}

	protected internal override void CreateNodes(XmlHelper helper, string path)
	{
		SetValueEnum(helper, path + "/@style", Style);
		SetValueColor(helper, path + "/d:color", Color);
	}

	protected internal override ExcelDxfBorderItem Clone()
	{
		return new ExcelDxfBorderItem(_styles)
		{
			Style = Style,
			Color = Color
		};
	}
}
