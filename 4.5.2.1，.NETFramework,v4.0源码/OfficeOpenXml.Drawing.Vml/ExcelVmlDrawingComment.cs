using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml;

public class ExcelVmlDrawingComment : ExcelVmlDrawingBase, IRangeID
{
	private const string VERTICAL_ALIGNMENT_PATH = "x:ClientData/x:TextVAlign";

	private const string HORIZONTAL_ALIGNMENT_PATH = "x:ClientData/x:TextHAlign";

	private const string VISIBLE_PATH = "x:ClientData/x:Visible";

	private const string BACKGROUNDCOLOR_PATH = "@fillcolor";

	private const string BACKGROUNDCOLOR2_PATH = "v:fill/@color2";

	private const string LINESTYLE_PATH = "v:stroke/@dashstyle";

	private const string ENDCAP_PATH = "v:stroke/@endcap";

	private const string LINECOLOR_PATH = "@strokecolor";

	private const string LINEWIDTH_PATH = "@strokeweight";

	private const string TEXTBOX_STYLE_PATH = "v:textbox/@style";

	private const string LOCKED_PATH = "x:ClientData/x:Locked";

	private const string LOCK_TEXT_PATH = "x:ClientData/x:LockText";

	private ExcelVmlDrawingPosition _from;

	private ExcelVmlDrawingPosition _to;

	private const string ROW_PATH = "x:ClientData/x:Row";

	private const string COLUMN_PATH = "x:ClientData/x:Column";

	private const string STYLE_PATH = "@style";

	internal ExcelRangeBase Range { get; set; }

	public string Address
	{
		get
		{
			return Range.Address;
		}
		internal set
		{
			Range.Address = value;
		}
	}

	public eTextAlignVerticalVml VerticalAlignment
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("x:ClientData/x:TextVAlign");
			if (!(xmlNodeString == "Center"))
			{
				if (xmlNodeString == "Bottom")
				{
					return eTextAlignVerticalVml.Bottom;
				}
				return eTextAlignVerticalVml.Top;
			}
			return eTextAlignVerticalVml.Center;
		}
		set
		{
			switch (value)
			{
			case eTextAlignVerticalVml.Center:
				SetXmlNodeString("x:ClientData/x:TextVAlign", "Center");
				break;
			case eTextAlignVerticalVml.Bottom:
				SetXmlNodeString("x:ClientData/x:TextVAlign", "Bottom");
				break;
			default:
				DeleteNode("x:ClientData/x:TextVAlign");
				break;
			}
		}
	}

	public eTextAlignHorizontalVml HorizontalAlignment
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("x:ClientData/x:TextHAlign");
			if (!(xmlNodeString == "Center"))
			{
				if (xmlNodeString == "Right")
				{
					return eTextAlignHorizontalVml.Right;
				}
				return eTextAlignHorizontalVml.Left;
			}
			return eTextAlignHorizontalVml.Center;
		}
		set
		{
			switch (value)
			{
			case eTextAlignHorizontalVml.Center:
				SetXmlNodeString("x:ClientData/x:TextHAlign", "Center");
				break;
			case eTextAlignHorizontalVml.Right:
				SetXmlNodeString("x:ClientData/x:TextHAlign", "Right");
				break;
			default:
				DeleteNode("x:ClientData/x:TextHAlign");
				break;
			}
		}
	}

	public bool Visible
	{
		get
		{
			return base.TopNode.SelectSingleNode("x:ClientData/x:Visible", base.NameSpaceManager) != null;
		}
		set
		{
			if (value)
			{
				CreateNode("x:ClientData/x:Visible");
				Style = SetStyle(Style, "visibility", "visible");
			}
			else
			{
				DeleteNode("x:ClientData/x:Visible");
				Style = SetStyle(Style, "visibility", "hidden");
			}
		}
	}

	public Color BackgroundColor
	{
		get
		{
			string text = GetXmlNodeString("@fillcolor");
			if (text == "")
			{
				return Color.FromArgb(255, 255, 225);
			}
			if (text.StartsWith("#"))
			{
				text = text.Substring(1, text.Length - 1);
			}
			if (int.TryParse(text, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var result))
			{
				return Color.FromArgb(result);
			}
			return Color.Empty;
		}
		set
		{
			string value2 = "#" + value.ToArgb().ToString("X").Substring(2, 6);
			SetXmlNodeString("@fillcolor", value2);
		}
	}

	public eLineStyleVml LineStyle
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("v:stroke/@dashstyle");
			if (xmlNodeString == "")
			{
				return eLineStyleVml.Solid;
			}
			if (xmlNodeString == "1 1")
			{
				xmlNodeString = GetXmlNodeString("v:stroke/@endcap");
				return (eLineStyleVml)Enum.Parse(typeof(eLineStyleVml), xmlNodeString, ignoreCase: true);
			}
			return (eLineStyleVml)Enum.Parse(typeof(eLineStyleVml), xmlNodeString, ignoreCase: true);
		}
		set
		{
			if (value == eLineStyleVml.Round || value == eLineStyleVml.Square)
			{
				SetXmlNodeString("v:stroke/@dashstyle", "1 1");
				if (value == eLineStyleVml.Round)
				{
					SetXmlNodeString("v:stroke/@endcap", "round");
				}
				else
				{
					DeleteNode("v:stroke/@endcap");
				}
			}
			else
			{
				string text = value.ToString();
				text = text.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + text.Substring(1, text.Length - 1);
				SetXmlNodeString("v:stroke/@dashstyle", text);
				DeleteNode("v:stroke/@endcap");
			}
		}
	}

	public Color LineColor
	{
		get
		{
			string text = GetXmlNodeString("@strokecolor");
			if (text == "")
			{
				return Color.Black;
			}
			if (text.StartsWith("#"))
			{
				text = text.Substring(1, text.Length - 1);
			}
			if (int.TryParse(text, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out var result))
			{
				return Color.FromArgb(result);
			}
			return Color.Empty;
		}
		set
		{
			string value2 = "#" + value.ToArgb().ToString("X").Substring(2, 6);
			SetXmlNodeString("@strokecolor", value2);
		}
	}

	public float LineWidth
	{
		get
		{
			string text = GetXmlNodeString("@strokeweight");
			if (text == "")
			{
				return 0.75f;
			}
			if (text.EndsWith("pt"))
			{
				text = text.Substring(0, text.Length - 2);
			}
			if (float.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
			{
				return result;
			}
			return 0f;
		}
		set
		{
			SetXmlNodeString("@strokeweight", value.ToString(CultureInfo.InvariantCulture) + "pt");
		}
	}

	public bool AutoFit
	{
		get
		{
			GetStyle(GetXmlNodeString("v:textbox/@style"), "mso-fit-shape-to-text", out var value);
			return value == "t";
		}
		set
		{
			SetXmlNodeString("v:textbox/@style", SetStyle(GetXmlNodeString("v:textbox/@style"), "mso-fit-shape-to-text", value ? "t" : ""));
		}
	}

	public bool Locked
	{
		get
		{
			return GetXmlNodeBool("x:ClientData/x:Locked", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("x:ClientData/x:Locked", value, removeIf: false);
		}
	}

	public bool LockText
	{
		get
		{
			return GetXmlNodeBool("x:ClientData/x:LockText", blankValue: false);
		}
		set
		{
			SetXmlNodeBool("x:ClientData/x:LockText", value, removeIf: false);
		}
	}

	public ExcelVmlDrawingPosition From
	{
		get
		{
			if (_from == null)
			{
				_from = new ExcelVmlDrawingPosition(base.NameSpaceManager, base.TopNode.SelectSingleNode("x:ClientData", base.NameSpaceManager), 0);
			}
			return _from;
		}
	}

	public ExcelVmlDrawingPosition To
	{
		get
		{
			if (_to == null)
			{
				_to = new ExcelVmlDrawingPosition(base.NameSpaceManager, base.TopNode.SelectSingleNode("x:ClientData", base.NameSpaceManager), 4);
			}
			return _to;
		}
	}

	internal int Row
	{
		get
		{
			return GetXmlNodeInt("x:ClientData/x:Row");
		}
		set
		{
			SetXmlNodeString("x:ClientData/x:Row", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	internal int Column
	{
		get
		{
			return GetXmlNodeInt("x:ClientData/x:Column");
		}
		set
		{
			SetXmlNodeString("x:ClientData/x:Column", value.ToString(CultureInfo.InvariantCulture));
		}
	}

	internal string Style
	{
		get
		{
			return GetXmlNodeString("@style");
		}
		set
		{
			SetXmlNodeString("@style", value);
		}
	}

	ulong IRangeID.RangeID
	{
		get
		{
			return ExcelCellBase.GetCellID(Range.Worksheet.SheetID, Range.Start.Row, Range.Start.Column);
		}
		set
		{
		}
	}

	internal ExcelVmlDrawingComment(XmlNode topNode, ExcelRangeBase range, XmlNamespaceManager ns)
		: base(topNode, ns)
	{
		Range = range;
		base.SchemaNodeOrder = new string[17]
		{
			"fill", "stroke", "shadow", "path", "textbox", "ClientData", "MoveWithCells", "SizeWithCells", "Anchor", "Locked",
			"AutoFill", "LockText", "TextHAlign", "TextVAlign", "Row", "Column", "Visible"
		};
	}
}
