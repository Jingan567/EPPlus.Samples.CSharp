using System;
using System.Globalization;
using System.IO;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public sealed class OfficeProperties : XmlHelper
{
	private XmlDocument _xmlPropertiesCore;

	private XmlDocument _xmlPropertiesExtended;

	private XmlDocument _xmlPropertiesCustom;

	private Uri _uriPropertiesCore = new Uri("/docProps/core.xml", UriKind.Relative);

	private Uri _uriPropertiesExtended = new Uri("/docProps/app.xml", UriKind.Relative);

	private Uri _uriPropertiesCustom = new Uri("/docProps/custom.xml", UriKind.Relative);

	private XmlHelper _coreHelper;

	private XmlHelper _extendedHelper;

	private XmlHelper _customHelper;

	private ExcelPackage _package;

	private const string TitlePath = "dc:title";

	private const string SubjectPath = "dc:subject";

	private const string AuthorPath = "dc:creator";

	private const string CommentsPath = "dc:description";

	private const string KeywordsPath = "cp:keywords";

	private const string LastModifiedByPath = "cp:lastModifiedBy";

	private const string LastPrintedPath = "cp:lastPrinted";

	private const string CreatedPath = "dcterms:created";

	private const string CategoryPath = "cp:category";

	private const string ContentStatusPath = "cp:contentStatus";

	private const string ApplicationPath = "xp:Properties/xp:Application";

	private const string HyperlinkBasePath = "xp:Properties/xp:HyperlinkBase";

	private const string AppVersionPath = "xp:Properties/xp:AppVersion";

	private const string CompanyPath = "xp:Properties/xp:Company";

	private const string ManagerPath = "xp:Properties/xp:Manager";

	private const string ModifiedPath = "dcterms:modified";

	private const string LinksUpToDatePath = "xp:Properties/xp:LinksUpToDate";

	private const string HyperlinksChangedPath = "xp:Properties/xp:HyperlinksChanged";

	private const string ScaleCropPath = "xp:Properties/xp:ScaleCrop";

	private const string SharedDocPath = "xp:Properties/xp:SharedDoc";

	public XmlDocument CorePropertiesXml
	{
		get
		{
			if (_xmlPropertiesCore == null)
			{
				string startXml = string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><cp:coreProperties xmlns:cp=\"{0}\" xmlns:dc=\"{1}\" xmlns:dcterms=\"{2}\" xmlns:dcmitype=\"{3}\" xmlns:xsi=\"{4}\"></cp:coreProperties>", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties", "http://purl.org/dc/elements/1.1/", "http://purl.org/dc/terms/", "http://purl.org/dc/dcmitype/", "http://www.w3.org/2001/XMLSchema-instance");
				_xmlPropertiesCore = GetXmlDocument(startXml, _uriPropertiesCore, "application/vnd.openxmlformats-package.core-properties+xml", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
			}
			return _xmlPropertiesCore;
		}
	}

	public string Title
	{
		get
		{
			return _coreHelper.GetXmlNodeString("dc:title");
		}
		set
		{
			_coreHelper.SetXmlNodeString("dc:title", value);
		}
	}

	public string Subject
	{
		get
		{
			return _coreHelper.GetXmlNodeString("dc:subject");
		}
		set
		{
			_coreHelper.SetXmlNodeString("dc:subject", value);
		}
	}

	public string Author
	{
		get
		{
			return _coreHelper.GetXmlNodeString("dc:creator");
		}
		set
		{
			_coreHelper.SetXmlNodeString("dc:creator", value);
		}
	}

	public string Comments
	{
		get
		{
			return _coreHelper.GetXmlNodeString("dc:description");
		}
		set
		{
			_coreHelper.SetXmlNodeString("dc:description", value);
		}
	}

	public string Keywords
	{
		get
		{
			return _coreHelper.GetXmlNodeString("cp:keywords");
		}
		set
		{
			_coreHelper.SetXmlNodeString("cp:keywords", value);
		}
	}

	public string LastModifiedBy
	{
		get
		{
			return _coreHelper.GetXmlNodeString("cp:lastModifiedBy");
		}
		set
		{
			_coreHelper.SetXmlNodeString("cp:lastModifiedBy", value);
		}
	}

	public string LastPrinted
	{
		get
		{
			return _coreHelper.GetXmlNodeString("cp:lastPrinted");
		}
		set
		{
			_coreHelper.SetXmlNodeString("cp:lastPrinted", value);
		}
	}

	public DateTime Created
	{
		get
		{
			if (!DateTime.TryParse(_coreHelper.GetXmlNodeString("dcterms:created"), out var result))
			{
				return DateTime.MinValue;
			}
			return result;
		}
		set
		{
			string value2 = value.ToUniversalTime().ToString("s", CultureInfo.InvariantCulture) + "Z";
			_coreHelper.SetXmlNodeString("dcterms:created", value2);
			_coreHelper.SetXmlNodeString("dcterms:created/@xsi:type", "dcterms:W3CDTF");
		}
	}

	public string Category
	{
		get
		{
			return _coreHelper.GetXmlNodeString("cp:category");
		}
		set
		{
			_coreHelper.SetXmlNodeString("cp:category", value);
		}
	}

	public string Status
	{
		get
		{
			return _coreHelper.GetXmlNodeString("cp:contentStatus");
		}
		set
		{
			_coreHelper.SetXmlNodeString("cp:contentStatus", value);
		}
	}

	public XmlDocument ExtendedPropertiesXml
	{
		get
		{
			if (_xmlPropertiesExtended == null)
			{
				_xmlPropertiesExtended = GetXmlDocument(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><Properties xmlns:vt=\"{0}\" xmlns=\"{1}\"></Properties>", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"), _uriPropertiesExtended, "application/vnd.openxmlformats-officedocument.extended-properties+xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
			}
			return _xmlPropertiesExtended;
		}
	}

	public string Application
	{
		get
		{
			return _extendedHelper.GetXmlNodeString("xp:Properties/xp:Application");
		}
		set
		{
			_extendedHelper.SetXmlNodeString("xp:Properties/xp:Application", value);
		}
	}

	public Uri HyperlinkBase
	{
		get
		{
			return new Uri(_extendedHelper.GetXmlNodeString("xp:Properties/xp:HyperlinkBase"), UriKind.Absolute);
		}
		set
		{
			_extendedHelper.SetXmlNodeString("xp:Properties/xp:HyperlinkBase", value.AbsoluteUri);
		}
	}

	public string AppVersion
	{
		get
		{
			return _extendedHelper.GetXmlNodeString("xp:Properties/xp:AppVersion");
		}
		set
		{
			_extendedHelper.SetXmlNodeString("xp:Properties/xp:AppVersion", value);
		}
	}

	public string Company
	{
		get
		{
			return _extendedHelper.GetXmlNodeString("xp:Properties/xp:Company");
		}
		set
		{
			_extendedHelper.SetXmlNodeString("xp:Properties/xp:Company", value);
		}
	}

	public string Manager
	{
		get
		{
			return _extendedHelper.GetXmlNodeString("xp:Properties/xp:Manager");
		}
		set
		{
			_extendedHelper.SetXmlNodeString("xp:Properties/xp:Manager", value);
		}
	}

	public DateTime Modified
	{
		get
		{
			if (!DateTime.TryParse(_coreHelper.GetXmlNodeString("dcterms:modified"), out var result))
			{
				return DateTime.MinValue;
			}
			return result;
		}
		set
		{
			string value2 = value.ToUniversalTime().ToString("s", CultureInfo.InvariantCulture) + "Z";
			_coreHelper.SetXmlNodeString("dcterms:modified", value2);
			_coreHelper.SetXmlNodeString("dcterms:modified/@xsi:type", "dcterms:W3CDTF");
		}
	}

	public bool LinksUpToDate
	{
		get
		{
			return _extendedHelper.GetXmlNodeBool("xp:Properties/xp:LinksUpToDate");
		}
		set
		{
			_extendedHelper.SetXmlNodeBool("xp:Properties/xp:LinksUpToDate", value);
		}
	}

	public bool HyperlinksChanged
	{
		get
		{
			return _extendedHelper.GetXmlNodeBool("xp:Properties/xp:HyperlinksChanged");
		}
		set
		{
			_extendedHelper.SetXmlNodeBool("xp:Properties/xp:HyperlinksChanged", value);
		}
	}

	public bool ScaleCrop
	{
		get
		{
			return _extendedHelper.GetXmlNodeBool("xp:Properties/xp:ScaleCrop");
		}
		set
		{
			_extendedHelper.SetXmlNodeBool("xp:Properties/xp:ScaleCrop", value);
		}
	}

	public bool SharedDoc
	{
		get
		{
			return _extendedHelper.GetXmlNodeBool("xp:Properties/xp:SharedDoc");
		}
		set
		{
			_extendedHelper.SetXmlNodeBool("xp:Properties/xp:SharedDoc", value);
		}
	}

	public XmlDocument CustomPropertiesXml
	{
		get
		{
			if (_xmlPropertiesCustom == null)
			{
				_xmlPropertiesCustom = GetXmlDocument(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><Properties xmlns:vt=\"{0}\" xmlns=\"{1}\"></Properties>", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"), _uriPropertiesCustom, "application/vnd.openxmlformats-officedocument.custom-properties+xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
			}
			return _xmlPropertiesCustom;
		}
	}

	internal OfficeProperties(ExcelPackage package, XmlNamespaceManager ns)
		: base(ns)
	{
		_package = package;
		_coreHelper = XmlHelperFactory.Create(ns, CorePropertiesXml.SelectSingleNode("cp:coreProperties", base.NameSpaceManager));
		_extendedHelper = XmlHelperFactory.Create(ns, ExtendedPropertiesXml);
		_customHelper = XmlHelperFactory.Create(ns, CustomPropertiesXml);
	}

	private XmlDocument GetXmlDocument(string startXml, Uri uri, string contentType, string relationship)
	{
		XmlDocument xmlDocument;
		if (_package.Package.PartExists(uri))
		{
			xmlDocument = _package.GetXmlFromUri(uri);
		}
		else
		{
			xmlDocument = new XmlDocument();
			xmlDocument.LoadXml(startXml);
			StreamWriter writer = new StreamWriter(_package.Package.CreatePart(uri, contentType).GetStream(FileMode.Create, FileAccess.Write));
			xmlDocument.Save(writer);
			_package.Package.Flush();
			_package.Package.CreateRelationship(UriHelper.GetRelativeUri(new Uri("/xl", UriKind.Relative), uri), TargetMode.Internal, relationship);
			_package.Package.Flush();
		}
		return xmlDocument;
	}

	public string GetExtendedPropertyValue(string propertyName)
	{
		string result = null;
		string xpath = $"xp:Properties/xp:{propertyName}";
		XmlNode xmlNode = ExtendedPropertiesXml.SelectSingleNode(xpath, base.NameSpaceManager);
		if (xmlNode != null)
		{
			result = xmlNode.InnerText;
		}
		return result;
	}

	public void SetExtendedPropertyValue(string propertyName, string value)
	{
		string path = $"xp:Properties/xp:{propertyName}";
		_extendedHelper.SetXmlNodeString(path, value);
	}

	public object GetCustomPropertyValue(string propertyName)
	{
		string xpath = $"ctp:Properties/ctp:property[@name='{propertyName}']";
		if (CustomPropertiesXml.SelectSingleNode(xpath, base.NameSpaceManager) is XmlElement xmlElement)
		{
			string innerText = xmlElement.LastChild.InnerText;
			switch (xmlElement.LastChild.LocalName)
			{
			case "filetime":
			{
				if (DateTime.TryParse(innerText, out var result))
				{
					return result;
				}
				return null;
			}
			case "i4":
			{
				if (int.TryParse(innerText, NumberStyles.Number, CultureInfo.InvariantCulture, out var result3))
				{
					return result3;
				}
				return null;
			}
			case "r8":
			{
				if (double.TryParse(innerText, NumberStyles.Number, CultureInfo.InvariantCulture, out var result2))
				{
					return result2;
				}
				return null;
			}
			case "bool":
				if (innerText == "true")
				{
					return true;
				}
				if (innerText == "false")
				{
					return false;
				}
				return null;
			default:
				return innerText;
			}
		}
		return null;
	}

	public void SetCustomPropertyValue(string propertyName, object value)
	{
		XmlNode xmlNode = CustomPropertiesXml.SelectSingleNode("ctp:Properties", base.NameSpaceManager);
		string xpath = $"ctp:Properties/ctp:property[@name='{propertyName}']";
		XmlElement xmlElement = CustomPropertiesXml.SelectSingleNode(xpath, base.NameSpaceManager) as XmlElement;
		if (xmlElement == null)
		{
			XmlNode xmlNode2 = CustomPropertiesXml.SelectSingleNode("ctp:Properties/ctp:property[not(@pid <= preceding-sibling::ctp:property/@pid) and not(@pid <= following-sibling::ctp:property/@pid)]", base.NameSpaceManager);
			int num;
			if (xmlNode2 == null)
			{
				num = 2;
			}
			else
			{
				if (!int.TryParse(xmlNode2.Attributes["pid"].Value, out num))
				{
					num = 2;
				}
				num++;
			}
			xmlElement = CustomPropertiesXml.CreateElement("property", "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties");
			xmlElement.SetAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
			xmlElement.SetAttribute("pid", num.ToString());
			xmlElement.SetAttribute("name", propertyName);
			xmlNode.AppendChild(xmlElement);
		}
		else
		{
			while (xmlElement.ChildNodes.Count > 0)
			{
				xmlElement.RemoveChild(xmlElement.ChildNodes[0]);
			}
		}
		XmlElement xmlElement2;
		if (value is bool)
		{
			xmlElement2 = CustomPropertiesXml.CreateElement("vt", "bool", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
			xmlElement2.InnerText = value.ToString().ToLower(CultureInfo.InvariantCulture);
		}
		else if (value is DateTime)
		{
			xmlElement2 = CustomPropertiesXml.CreateElement("vt", "filetime", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
			xmlElement2.InnerText = ((DateTime)value).AddHours(-1.0).ToString("yyyy-MM-ddTHH:mm:ssZ");
		}
		else if (value is short || value is int)
		{
			xmlElement2 = CustomPropertiesXml.CreateElement("vt", "i4", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
			xmlElement2.InnerText = value.ToString();
		}
		else if (value is double || value is decimal || value is float || value is long)
		{
			xmlElement2 = CustomPropertiesXml.CreateElement("vt", "r8", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
			if (value is double)
			{
				xmlElement2.InnerText = ((double)value).ToString(CultureInfo.InvariantCulture);
			}
			else if (value is float)
			{
				xmlElement2.InnerText = ((float)value).ToString(CultureInfo.InvariantCulture);
			}
			else if (value is decimal)
			{
				xmlElement2.InnerText = ((decimal)value).ToString(CultureInfo.InvariantCulture);
			}
			else
			{
				xmlElement2.InnerText = value.ToString();
			}
		}
		else
		{
			xmlElement2 = CustomPropertiesXml.CreateElement("vt", "lpwstr", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
			xmlElement2.InnerText = value.ToString();
		}
		xmlElement.AppendChild(xmlElement2);
	}

	internal void Save()
	{
		if (_xmlPropertiesCore != null)
		{
			_package.SavePart(_uriPropertiesCore, _xmlPropertiesCore);
		}
		if (_xmlPropertiesExtended != null)
		{
			_package.SavePart(_uriPropertiesExtended, _xmlPropertiesExtended);
		}
		if (_xmlPropertiesCustom != null)
		{
			_package.SavePart(_uriPropertiesCustom, _xmlPropertiesCustom);
		}
	}
}
