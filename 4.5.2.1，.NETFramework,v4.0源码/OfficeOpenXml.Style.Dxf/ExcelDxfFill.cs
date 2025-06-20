namespace OfficeOpenXml.Style.Dxf;

public class ExcelDxfFill : DxfStyleBase<ExcelDxfFill>
{
	public ExcelFillStyle? PatternType { get; set; }

	public ExcelDxfColor PatternColor { get; internal set; }

	public ExcelDxfColor BackgroundColor { get; internal set; }

	protected internal override string Id => GetAsString(PatternType) + "|" + ((PatternColor == null) ? "" : PatternColor.Id) + "|" + ((BackgroundColor == null) ? "" : BackgroundColor.Id);

	protected internal override bool HasValue
	{
		get
		{
			if (!PatternType.HasValue && !PatternColor.HasValue)
			{
				return BackgroundColor.HasValue;
			}
			return true;
		}
	}

	public ExcelDxfFill(ExcelStyles styles)
		: base(styles)
	{
		PatternColor = new ExcelDxfColor(styles);
		BackgroundColor = new ExcelDxfColor(styles);
	}

	protected internal override void CreateNodes(XmlHelper helper, string path)
	{
		helper.CreateNode(path);
		SetValueEnum(helper, path + "/d:patternFill/@patternType", PatternType);
		SetValueColor(helper, path + "/d:patternFill/d:fgColor", PatternColor);
		SetValueColor(helper, path + "/d:patternFill/d:bgColor", BackgroundColor);
	}

	protected internal override ExcelDxfFill Clone()
	{
		return new ExcelDxfFill(_styles)
		{
			PatternType = PatternType,
			PatternColor = PatternColor.Clone(),
			BackgroundColor = BackgroundColor.Clone()
		};
	}
}
