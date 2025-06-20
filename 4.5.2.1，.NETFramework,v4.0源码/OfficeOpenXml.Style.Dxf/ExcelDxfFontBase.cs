namespace OfficeOpenXml.Style.Dxf;

public class ExcelDxfFontBase : DxfStyleBase<ExcelDxfFontBase>
{
	public bool? Bold { get; set; }

	public bool? Italic { get; set; }

	public bool? Strike { get; set; }

	public ExcelDxfColor Color { get; set; }

	public ExcelUnderLineType? Underline { get; set; }

	protected internal override string Id => GetAsString(Bold) + "|" + GetAsString(Italic) + "|" + GetAsString(Strike) + "|" + ((Color == null) ? "" : Color.Id) + "|" + GetAsString(Underline);

	protected internal override bool HasValue
	{
		get
		{
			if (!Bold.HasValue && !Italic.HasValue && !Strike.HasValue && !Underline.HasValue)
			{
				return Color.HasValue;
			}
			return true;
		}
	}

	public ExcelDxfFontBase(ExcelStyles styles)
		: base(styles)
	{
		Color = new ExcelDxfColor(styles);
	}

	protected internal override void CreateNodes(XmlHelper helper, string path)
	{
		helper.CreateNode(path);
		SetValueBool(helper, path + "/d:b/@val", Bold);
		SetValueBool(helper, path + "/d:i/@val", Italic);
		SetValueBool(helper, path + "/d:strike", Strike);
		SetValue(helper, path + "/d:u/@val", Underline);
		SetValueColor(helper, path + "/d:color", Color);
	}

	protected internal override ExcelDxfFontBase Clone()
	{
		return new ExcelDxfFontBase(_styles)
		{
			Bold = Bold,
			Color = Color.Clone(),
			Italic = Italic,
			Strike = Strike,
			Underline = Underline
		};
	}
}
