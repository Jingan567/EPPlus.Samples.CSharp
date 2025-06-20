using System.Xml;

namespace OfficeOpenXml.Style.XmlAccess;

public sealed class ExcelBorderXml : StyleXmlHelper
{
	private const string leftPath = "d:left";

	private ExcelBorderItemXml _left;

	private const string rightPath = "d:right";

	private ExcelBorderItemXml _right;

	private const string topPath = "d:top";

	private ExcelBorderItemXml _top;

	private const string bottomPath = "d:bottom";

	private ExcelBorderItemXml _bottom;

	private const string diagonalPath = "d:diagonal";

	private ExcelBorderItemXml _diagonal;

	private const string diagonalUpPath = "@diagonalUp";

	private bool _diagonalUp;

	private const string diagonalDownPath = "@diagonalDown";

	private bool _diagonalDown;

	internal override string Id => Left.Id + Right.Id + Top.Id + Bottom.Id + Diagonal.Id + DiagonalUp + DiagonalDown;

	public ExcelBorderItemXml Left
	{
		get
		{
			return _left;
		}
		internal set
		{
			_left = value;
		}
	}

	public ExcelBorderItemXml Right
	{
		get
		{
			return _right;
		}
		internal set
		{
			_right = value;
		}
	}

	public ExcelBorderItemXml Top
	{
		get
		{
			return _top;
		}
		internal set
		{
			_top = value;
		}
	}

	public ExcelBorderItemXml Bottom
	{
		get
		{
			return _bottom;
		}
		internal set
		{
			_bottom = value;
		}
	}

	public ExcelBorderItemXml Diagonal
	{
		get
		{
			return _diagonal;
		}
		internal set
		{
			_diagonal = value;
		}
	}

	public bool DiagonalUp
	{
		get
		{
			return _diagonalUp;
		}
		internal set
		{
			_diagonalUp = value;
		}
	}

	public bool DiagonalDown
	{
		get
		{
			return _diagonalDown;
		}
		internal set
		{
			_diagonalDown = value;
		}
	}

	internal ExcelBorderXml(XmlNamespaceManager nameSpaceManager)
		: base(nameSpaceManager)
	{
	}

	internal ExcelBorderXml(XmlNamespaceManager nsm, XmlNode topNode)
		: base(nsm, topNode)
	{
		_left = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode("d:left", nsm));
		_right = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode("d:right", nsm));
		_top = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode("d:top", nsm));
		_bottom = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode("d:bottom", nsm));
		_diagonal = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode("d:diagonal", nsm));
		_diagonalUp = GetBoolValue(topNode, "@diagonalUp");
		_diagonalDown = GetBoolValue(topNode, "@diagonalDown");
	}

	internal ExcelBorderXml Copy()
	{
		return new ExcelBorderXml(base.NameSpaceManager)
		{
			Bottom = _bottom.Copy(),
			Diagonal = _diagonal.Copy(),
			Left = _left.Copy(),
			Right = _right.Copy(),
			Top = _top.Copy(),
			DiagonalUp = _diagonalUp,
			DiagonalDown = _diagonalDown
		};
	}

	internal override XmlNode CreateXmlNode(XmlNode topNode)
	{
		base.TopNode = topNode;
		CreateNode("d:left");
		topNode.AppendChild(_left.CreateXmlNode(base.TopNode.SelectSingleNode("d:left", base.NameSpaceManager)));
		CreateNode("d:right");
		topNode.AppendChild(_right.CreateXmlNode(base.TopNode.SelectSingleNode("d:right", base.NameSpaceManager)));
		CreateNode("d:top");
		topNode.AppendChild(_top.CreateXmlNode(base.TopNode.SelectSingleNode("d:top", base.NameSpaceManager)));
		CreateNode("d:bottom");
		topNode.AppendChild(_bottom.CreateXmlNode(base.TopNode.SelectSingleNode("d:bottom", base.NameSpaceManager)));
		CreateNode("d:diagonal");
		topNode.AppendChild(_diagonal.CreateXmlNode(base.TopNode.SelectSingleNode("d:diagonal", base.NameSpaceManager)));
		if (_diagonalUp)
		{
			SetXmlNodeString("@diagonalUp", "1");
		}
		if (_diagonalDown)
		{
			SetXmlNodeString("@diagonalDown", "1");
		}
		return topNode;
	}
}
