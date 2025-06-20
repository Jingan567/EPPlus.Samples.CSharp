using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Style.XmlAccess;

public class ExcelFillXml : StyleXmlHelper
{
	private const string fillPatternTypePath = "d:patternFill/@patternType";

	protected ExcelFillStyle _fillPatternType;

	protected ExcelColorXml _patternColor;

	private const string _patternColorPath = "d:patternFill/d:bgColor";

	protected ExcelColorXml _backgroundColor;

	private const string _backgroundColorPath = "d:patternFill/d:fgColor";

	internal override string Id => string.Concat(PatternType, PatternColor.Id, BackgroundColor.Id);

	public ExcelFillStyle PatternType
	{
		get
		{
			return _fillPatternType;
		}
		set
		{
			_fillPatternType = value;
		}
	}

	public ExcelColorXml PatternColor
	{
		get
		{
			return _patternColor;
		}
		internal set
		{
			_patternColor = value;
		}
	}

	public ExcelColorXml BackgroundColor
	{
		get
		{
			return _backgroundColor;
		}
		internal set
		{
			_backgroundColor = value;
		}
	}

	internal ExcelFillXml(XmlNamespaceManager nameSpaceManager)
		: base(nameSpaceManager)
	{
		_fillPatternType = ExcelFillStyle.None;
		_backgroundColor = new ExcelColorXml(base.NameSpaceManager);
		_patternColor = new ExcelColorXml(base.NameSpaceManager);
	}

	internal ExcelFillXml(XmlNamespaceManager nsm, XmlNode topNode)
		: base(nsm, topNode)
	{
		PatternType = GetPatternType(GetXmlNodeString("d:patternFill/@patternType"));
		_backgroundColor = new ExcelColorXml(nsm, topNode.SelectSingleNode("d:patternFill/d:fgColor", nsm));
		_patternColor = new ExcelColorXml(nsm, topNode.SelectSingleNode("d:patternFill/d:bgColor", nsm));
	}

	private ExcelFillStyle GetPatternType(string patternType)
	{
		if (patternType == "")
		{
			return ExcelFillStyle.None;
		}
		patternType = patternType.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + patternType.Substring(1, patternType.Length - 1);
		try
		{
			return (ExcelFillStyle)Enum.Parse(typeof(ExcelFillStyle), patternType);
		}
		catch
		{
			return ExcelFillStyle.None;
		}
	}

	internal virtual ExcelFillXml Copy()
	{
		return new ExcelFillXml(base.NameSpaceManager)
		{
			PatternType = _fillPatternType,
			BackgroundColor = _backgroundColor.Copy(),
			PatternColor = _patternColor.Copy()
		};
	}

	internal override XmlNode CreateXmlNode(XmlNode topNode)
	{
		base.TopNode = topNode;
		SetXmlNodeString("d:patternFill/@patternType", SetPatternString(_fillPatternType));
		if (PatternType != 0)
		{
			topNode.SelectSingleNode("d:patternFill/@patternType", base.NameSpaceManager);
			if (BackgroundColor.Exists)
			{
				CreateNode("d:patternFill/d:fgColor");
				BackgroundColor.CreateXmlNode(topNode.SelectSingleNode("d:patternFill/d:fgColor", base.NameSpaceManager));
				if (PatternColor.Exists)
				{
					CreateNode("d:patternFill/d:bgColor");
					PatternColor.CreateXmlNode(topNode.SelectSingleNode("d:patternFill/d:bgColor", base.NameSpaceManager));
				}
			}
		}
		return topNode;
	}

	private string SetPatternString(ExcelFillStyle pattern)
	{
		string name = Enum.GetName(typeof(ExcelFillStyle), pattern);
		return name.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + name.Substring(1, name.Length - 1);
	}
}
