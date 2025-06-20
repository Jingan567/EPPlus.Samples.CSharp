using System;
using System.Collections;
using System.Xml;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public sealed class ExcelHeaderFooter : XmlHelper
{
	public const string PageNumber = "&P";

	public const string NumberOfPages = "&N";

	public const string FontColor = "&K";

	public const string SheetName = "&A";

	public const string FilePath = "&Z";

	public const string FileName = "&F";

	public const string CurrentDate = "&D";

	public const string CurrentTime = "&T";

	public const string Image = "&G";

	public const string OutlineStyle = "&O";

	public const string ShadowStyle = "&H";

	internal ExcelHeaderFooterText _oddHeader;

	internal ExcelHeaderFooterText _oddFooter;

	internal ExcelHeaderFooterText _evenHeader;

	internal ExcelHeaderFooterText _evenFooter;

	internal ExcelHeaderFooterText _firstHeader;

	internal ExcelHeaderFooterText _firstFooter;

	private ExcelWorksheet _ws;

	private const string alignWithMarginsPath = "@alignWithMargins";

	private const string differentOddEvenPath = "@differentOddEven";

	private const string differentFirstPath = "@differentFirst";

	private const string scaleWithDocPath = "@scaleWithDoc";

	private ExcelVmlDrawingPictureCollection _vmlDrawingsHF;

	public bool AlignWithMargins
	{
		get
		{
			return GetXmlNodeBool("@alignWithMargins");
		}
		set
		{
			SetXmlNodeString("@alignWithMargins", value ? "1" : "0");
		}
	}

	public bool differentOddEven
	{
		get
		{
			return GetXmlNodeBool("@differentOddEven");
		}
		set
		{
			SetXmlNodeString("@differentOddEven", value ? "1" : "0");
		}
	}

	public bool differentFirst
	{
		get
		{
			return GetXmlNodeBool("@differentFirst");
		}
		set
		{
			SetXmlNodeString("@differentFirst", value ? "1" : "0");
		}
	}

	public bool ScaleWithDocument
	{
		get
		{
			return GetXmlNodeBool("@scaleWithDoc");
		}
		set
		{
			SetXmlNodeBool("@scaleWithDoc", value);
		}
	}

	public ExcelHeaderFooterText OddHeader
	{
		get
		{
			if (_oddHeader == null)
			{
				_oddHeader = new ExcelHeaderFooterText(base.TopNode.SelectSingleNode("d:oddHeader", base.NameSpaceManager), _ws, "H");
			}
			return _oddHeader;
		}
	}

	public ExcelHeaderFooterText OddFooter
	{
		get
		{
			if (_oddFooter == null)
			{
				_oddFooter = new ExcelHeaderFooterText(base.TopNode.SelectSingleNode("d:oddFooter", base.NameSpaceManager), _ws, "F");
			}
			return _oddFooter;
		}
	}

	public ExcelHeaderFooterText EvenHeader
	{
		get
		{
			if (_evenHeader == null)
			{
				_evenHeader = new ExcelHeaderFooterText(base.TopNode.SelectSingleNode("d:evenHeader", base.NameSpaceManager), _ws, "HEVEN");
				differentOddEven = true;
			}
			return _evenHeader;
		}
	}

	public ExcelHeaderFooterText EvenFooter
	{
		get
		{
			if (_evenFooter == null)
			{
				_evenFooter = new ExcelHeaderFooterText(base.TopNode.SelectSingleNode("d:evenFooter", base.NameSpaceManager), _ws, "FEVEN");
				differentOddEven = true;
			}
			return _evenFooter;
		}
	}

	public ExcelHeaderFooterText FirstHeader
	{
		get
		{
			if (_firstHeader == null)
			{
				_firstHeader = new ExcelHeaderFooterText(base.TopNode.SelectSingleNode("d:firstHeader", base.NameSpaceManager), _ws, "HFIRST");
				differentFirst = true;
			}
			return _firstHeader;
		}
	}

	public ExcelHeaderFooterText FirstFooter
	{
		get
		{
			if (_firstFooter == null)
			{
				_firstFooter = new ExcelHeaderFooterText(base.TopNode.SelectSingleNode("d:firstFooter", base.NameSpaceManager), _ws, "FFIRST");
				differentFirst = true;
			}
			return _firstFooter;
		}
	}

	public ExcelVmlDrawingPictureCollection Pictures
	{
		get
		{
			if (_vmlDrawingsHF == null)
			{
				XmlNode xmlNode = _ws.WorksheetXml.SelectSingleNode("d:worksheet/d:legacyDrawingHF/@r:id", base.NameSpaceManager);
				if (xmlNode == null)
				{
					_vmlDrawingsHF = new ExcelVmlDrawingPictureCollection(_ws._package, _ws, null);
				}
				else if (_ws.Part.RelationshipExists(xmlNode.Value))
				{
					ZipPackageRelationship relationship = _ws.Part.GetRelationship(xmlNode.Value);
					Uri uri = UriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
					_vmlDrawingsHF = new ExcelVmlDrawingPictureCollection(_ws._package, _ws, uri);
					_vmlDrawingsHF.RelId = relationship.Id;
				}
			}
			return _vmlDrawingsHF;
		}
	}

	internal ExcelHeaderFooter(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelWorksheet ws)
		: base(nameSpaceManager, topNode)
	{
		_ws = ws;
		base.SchemaNodeOrder = new string[7] { "headerFooter", "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter" };
	}

	internal void Save()
	{
		if (_oddHeader != null)
		{
			SetXmlNodeString("d:oddHeader", GetText(OddHeader));
		}
		if (_oddFooter != null)
		{
			SetXmlNodeString("d:oddFooter", GetText(OddFooter));
		}
		if (differentOddEven)
		{
			if (_evenHeader != null)
			{
				SetXmlNodeString("d:evenHeader", GetText(EvenHeader));
			}
			if (_evenFooter != null)
			{
				SetXmlNodeString("d:evenFooter", GetText(EvenFooter));
			}
		}
		if (differentFirst)
		{
			if (_firstHeader != null)
			{
				SetXmlNodeString("d:firstHeader", GetText(FirstHeader));
			}
			if (_firstFooter != null)
			{
				SetXmlNodeString("d:firstFooter", GetText(FirstFooter));
			}
		}
	}

	internal void SaveHeaderFooterImages()
	{
		if (_vmlDrawingsHF == null)
		{
			return;
		}
		if (_vmlDrawingsHF.Count == 0)
		{
			if (_vmlDrawingsHF.Uri != null)
			{
				_ws.Part.DeleteRelationship(_vmlDrawingsHF.RelId);
				_ws._package.Package.DeletePart(_vmlDrawingsHF.Uri);
			}
			return;
		}
		if (_vmlDrawingsHF.Uri == null)
		{
			_vmlDrawingsHF.Uri = XmlHelper.GetNewUri(_ws._package.Package, "/xl/drawings/vmlDrawing{0}.vml");
		}
		if (_vmlDrawingsHF.Part == null)
		{
			_vmlDrawingsHF.Part = _ws._package.Package.CreatePart(_vmlDrawingsHF.Uri, "application/vnd.openxmlformats-officedocument.vmlDrawing", _ws._package.Compression);
			ZipPackageRelationship zipPackageRelationship = _ws.Part.CreateRelationship(UriHelper.GetRelativeUri(_ws.WorksheetUri, _vmlDrawingsHF.Uri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing");
			_ws.SetHFLegacyDrawingRel(zipPackageRelationship.Id);
			_vmlDrawingsHF.RelId = zipPackageRelationship.Id;
			foreach (ExcelVmlDrawingPicture item in (IEnumerable)_vmlDrawingsHF)
			{
				zipPackageRelationship = _vmlDrawingsHF.Part.CreateRelationship(UriHelper.GetRelativeUri(_vmlDrawingsHF.Uri, item.ImageUri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
				item.RelId = zipPackageRelationship.Id;
			}
		}
		_vmlDrawingsHF.VmlDrawingXml.Save(_vmlDrawingsHF.Part.GetStream());
	}

	private string GetText(ExcelHeaderFooterText headerFooter)
	{
		string text = "";
		if (headerFooter.LeftAlignedText != null)
		{
			text = text + "&L" + headerFooter.LeftAlignedText;
		}
		if (headerFooter.CenteredText != null)
		{
			text = text + "&C" + headerFooter.CenteredText;
		}
		if (headerFooter.RightAlignedText != null)
		{
			text = text + "&R" + headerFooter.RightAlignedText;
		}
		return text;
	}
}
