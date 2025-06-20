namespace OfficeOpenXml.Style.Dxf;

public class ExcelDxfNumberFormat : DxfStyleBase<ExcelDxfNumberFormat>
{
	private int _numFmtID = int.MinValue;

	private string _format = "";

	public int NumFmtID
	{
		get
		{
			return _numFmtID;
		}
		internal set
		{
			_numFmtID = value;
		}
	}

	public string Format
	{
		get
		{
			return _format;
		}
		set
		{
			_format = value;
			NumFmtID = ExcelNumberFormat.GetFromBuildIdFromFormat(value);
		}
	}

	protected internal override string Id => Format;

	protected internal override bool HasValue => !string.IsNullOrEmpty(Format);

	public ExcelDxfNumberFormat(ExcelStyles styles)
		: base(styles)
	{
	}

	protected internal override void CreateNodes(XmlHelper helper, string path)
	{
		if (NumFmtID < 0 && !string.IsNullOrEmpty(Format))
		{
			NumFmtID = _styles._nextDfxNumFmtID++;
		}
		helper.CreateNode(path);
		SetValue(helper, path + "/@numFmtId", NumFmtID);
		SetValue(helper, path + "/@formatCode", Format);
	}

	protected internal override ExcelDxfNumberFormat Clone()
	{
		return new ExcelDxfNumberFormat(_styles)
		{
			NumFmtID = NumFmtID,
			Format = Format
		};
	}
}
