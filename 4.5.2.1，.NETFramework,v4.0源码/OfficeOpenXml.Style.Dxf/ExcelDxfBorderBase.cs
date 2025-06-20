namespace OfficeOpenXml.Style.Dxf;

public class ExcelDxfBorderBase : DxfStyleBase<ExcelDxfBorderBase>
{
	public ExcelDxfBorderItem Left { get; internal set; }

	public ExcelDxfBorderItem Right { get; internal set; }

	public ExcelDxfBorderItem Top { get; internal set; }

	public ExcelDxfBorderItem Bottom { get; internal set; }

	protected internal override string Id => Top.Id + Bottom.Id + Left.Id + Right.Id;

	protected internal override bool HasValue
	{
		get
		{
			if (!Left.HasValue && !Right.HasValue && !Top.HasValue)
			{
				return Bottom.HasValue;
			}
			return true;
		}
	}

	internal ExcelDxfBorderBase(ExcelStyles styles)
		: base(styles)
	{
		Left = new ExcelDxfBorderItem(_styles);
		Right = new ExcelDxfBorderItem(_styles);
		Top = new ExcelDxfBorderItem(_styles);
		Bottom = new ExcelDxfBorderItem(_styles);
	}

	protected internal override void CreateNodes(XmlHelper helper, string path)
	{
		Left.CreateNodes(helper, path + "/d:left");
		Right.CreateNodes(helper, path + "/d:right");
		Top.CreateNodes(helper, path + "/d:top");
		Bottom.CreateNodes(helper, path + "/d:bottom");
	}

	protected internal override ExcelDxfBorderBase Clone()
	{
		return new ExcelDxfBorderBase(_styles)
		{
			Bottom = Bottom.Clone(),
			Top = Top.Clone(),
			Left = Left.Clone(),
			Right = Right.Clone()
		};
	}
}
