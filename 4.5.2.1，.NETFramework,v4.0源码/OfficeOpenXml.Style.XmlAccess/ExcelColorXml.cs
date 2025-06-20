using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Style.XmlAccess;

public sealed class ExcelColorXml : StyleXmlHelper
{
	private bool _auto;

	private string _theme;

	private decimal _tint;

	private string _rgb;

	private int _indexed;

	private bool _exists;

	internal override string Id => _auto.ToString() + "|" + _theme + "|" + _tint + "|" + _rgb + "|" + _indexed;

	public bool Auto
	{
		get
		{
			return _auto;
		}
		set
		{
			_auto = value;
			_exists = true;
			Clear();
		}
	}

	public string Theme => _theme;

	public decimal Tint
	{
		get
		{
			if (_tint == decimal.MinValue)
			{
				return 0m;
			}
			return _tint;
		}
		set
		{
			_tint = value;
			_exists = true;
		}
	}

	public string Rgb
	{
		get
		{
			return _rgb;
		}
		set
		{
			_rgb = value;
			_exists = true;
			_indexed = int.MinValue;
			_auto = false;
		}
	}

	public int Indexed
	{
		get
		{
			if (_indexed != int.MinValue)
			{
				return _indexed;
			}
			return 0;
		}
		set
		{
			if (value < 0 || value > 65)
			{
				throw new ArgumentOutOfRangeException("Index out of range");
			}
			Clear();
			_indexed = value;
			_exists = true;
		}
	}

	internal bool Exists => _exists;

	internal ExcelColorXml(XmlNamespaceManager nameSpaceManager)
		: base(nameSpaceManager)
	{
		_auto = false;
		_theme = "";
		_tint = default(decimal);
		_rgb = "";
		_indexed = int.MinValue;
	}

	internal ExcelColorXml(XmlNamespaceManager nsm, XmlNode topNode)
		: base(nsm, topNode)
	{
		if (topNode == null)
		{
			_exists = false;
			return;
		}
		_exists = true;
		_auto = GetXmlNodeBool("@auto");
		_theme = GetXmlNodeString("@theme");
		_tint = GetXmlNodeDecimalNull("@tint") ?? decimal.MinValue;
		_rgb = GetXmlNodeString("@rgb");
		_indexed = GetXmlNodeIntNull("@indexed") ?? int.MinValue;
	}

	internal void Clear()
	{
		_theme = "";
		_tint = decimal.MinValue;
		_indexed = int.MinValue;
		_rgb = "";
		_auto = false;
	}

	public void SetColor(Color color)
	{
		Clear();
		_rgb = color.ToArgb().ToString("X");
	}

	internal ExcelColorXml Copy()
	{
		return new ExcelColorXml(base.NameSpaceManager)
		{
			_indexed = _indexed,
			_tint = _tint,
			_rgb = _rgb,
			_theme = _theme,
			_auto = _auto,
			_exists = _exists
		};
	}

	internal override XmlNode CreateXmlNode(XmlNode topNode)
	{
		base.TopNode = topNode;
		if (_rgb != "")
		{
			SetXmlNodeString("@rgb", _rgb);
		}
		else if (_indexed >= 0)
		{
			SetXmlNodeString("@indexed", _indexed.ToString());
		}
		else if (_auto)
		{
			SetXmlNodeBool("@auto", _auto);
		}
		else
		{
			SetXmlNodeString("@theme", _theme.ToString());
		}
		if (_tint != decimal.MinValue)
		{
			SetXmlNodeString("@tint", _tint.ToString(CultureInfo.InvariantCulture));
		}
		return base.TopNode;
	}
}
