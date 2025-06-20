using System;

namespace OfficeOpenXml;

public class ExcelHyperLink : Uri
{
	private string _referenceAddress;

	private string _display = "";

	private int _colSpann;

	private int _rowSpann;

	public string ReferenceAddress
	{
		get
		{
			return _referenceAddress;
		}
		set
		{
			_referenceAddress = value;
		}
	}

	public string Display
	{
		get
		{
			return _display;
		}
		set
		{
			_display = value;
		}
	}

	public string ToolTip { get; set; }

	public int ColSpann
	{
		get
		{
			return _colSpann;
		}
		set
		{
			_colSpann = value;
		}
	}

	public int RowSpann
	{
		get
		{
			return _rowSpann;
		}
		set
		{
			_rowSpann = value;
		}
	}

	public Uri OriginalUri { get; internal set; }

	internal string RId { get; set; }

	public ExcelHyperLink(string uriString)
		: base(uriString)
	{
		OriginalUri = this;
	}

	[Obsolete("base constructor 'System.Uri.Uri(string, bool)' is obsolete: 'The constructor has been deprecated. Please use new ExcelHyperLink(string). The dontEscape parameter is deprecated and is always false.")]
	public ExcelHyperLink(string uriString, bool dontEscape)
		: base(uriString, dontEscape)
	{
		OriginalUri = this;
	}

	public ExcelHyperLink(string uriString, UriKind uriKind)
		: base(uriString, uriKind)
	{
		OriginalUri = this;
	}

	public ExcelHyperLink(string referenceAddress, string display)
		: base("xl://internal")
	{
		_referenceAddress = referenceAddress;
		_display = display;
	}
}
