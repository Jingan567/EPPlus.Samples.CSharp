using System.Xml;

namespace OfficeOpenXml.Style.XmlAccess;

public sealed class ExcelNamedStyleXml : StyleXmlHelper
{
	private ExcelStyles _styles;

	private int _styleXfId;

	private const string idPath = "@xfId";

	private int _xfId = int.MinValue;

	private const string buildInIdPath = "@builtinId";

	private const string customBuiltinPath = "@customBuiltin";

	private const string namePath = "@name";

	private string _name;

	private ExcelStyle _style;

	internal override string Id => Name;

	public int StyleXfId
	{
		get
		{
			return _styleXfId;
		}
		set
		{
			_styleXfId = value;
		}
	}

	internal int XfId
	{
		get
		{
			return _xfId;
		}
		set
		{
			_xfId = value;
		}
	}

	public int BuildInId { get; set; }

	public bool CustomBuildin { get; set; }

	public string Name
	{
		get
		{
			return _name;
		}
		internal set
		{
			_name = value;
		}
	}

	public ExcelStyle Style
	{
		get
		{
			return _style;
		}
		internal set
		{
			_style = value;
		}
	}

	internal ExcelNamedStyleXml(XmlNamespaceManager nameSpaceManager, ExcelStyles styles)
		: base(nameSpaceManager)
	{
		_styles = styles;
		BuildInId = int.MinValue;
	}

	internal ExcelNamedStyleXml(XmlNamespaceManager NameSpaceManager, XmlNode topNode, ExcelStyles styles)
		: base(NameSpaceManager, topNode)
	{
		StyleXfId = GetXmlNodeInt("@xfId");
		Name = GetXmlNodeString("@name");
		BuildInId = GetXmlNodeInt("@builtinId");
		CustomBuildin = GetXmlNodeBool("@customBuiltin");
		_styles = styles;
		_style = new ExcelStyle(styles, styles.NamedStylePropertyChange, -1, Name, _styleXfId);
	}

	internal override XmlNode CreateXmlNode(XmlNode topNode)
	{
		base.TopNode = topNode;
		SetXmlNodeString("@name", _name);
		SetXmlNodeString("@xfId", _styles.CellStyleXfs[StyleXfId].newID.ToString());
		if (BuildInId >= 0)
		{
			SetXmlNodeString("@builtinId", BuildInId.ToString());
		}
		if (CustomBuildin)
		{
			SetXmlNodeBool("@customBuiltin", value: true);
		}
		return base.TopNode;
	}
}
