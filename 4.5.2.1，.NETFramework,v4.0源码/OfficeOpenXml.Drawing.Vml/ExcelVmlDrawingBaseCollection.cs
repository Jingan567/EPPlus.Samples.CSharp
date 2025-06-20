using System;
using System.Xml;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.Drawing.Vml;

public class ExcelVmlDrawingBaseCollection
{
	internal XmlDocument VmlDrawingXml { get; set; }

	internal Uri Uri { get; set; }

	internal string RelId { get; set; }

	internal ZipPackagePart Part { get; set; }

	internal XmlNamespaceManager NameSpaceManager { get; set; }

	internal ExcelVmlDrawingBaseCollection(ExcelPackage pck, ExcelWorksheet ws, Uri uri)
	{
		VmlDrawingXml = new XmlDocument();
		VmlDrawingXml.PreserveWhitespace = false;
		NameTable nameTable = new NameTable();
		NameSpaceManager = new XmlNamespaceManager(nameTable);
		NameSpaceManager.AddNamespace("v", "urn:schemas-microsoft-com:vml");
		NameSpaceManager.AddNamespace("o", "urn:schemas-microsoft-com:office:office");
		NameSpaceManager.AddNamespace("x", "urn:schemas-microsoft-com:office:excel");
		Uri = uri;
		if (uri == null)
		{
			Part = null;
			return;
		}
		Part = pck.Package.GetPart(uri);
		XmlHelper.LoadXmlSafe(VmlDrawingXml, Part.GetStream());
	}
}
