using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing.Vml;

public class ExcelVmlDrawingPictureCollection : ExcelVmlDrawingBaseCollection, IEnumerable
{
	internal List<ExcelVmlDrawingPicture> _images;

	private ExcelPackage _pck;

	private ExcelWorksheet _ws;

	private int _nextID;

	public ExcelVmlDrawingPicture this[int Index] => _images[Index];

	public int Count => _images.Count;

	internal ExcelVmlDrawingPictureCollection(ExcelPackage pck, ExcelWorksheet ws, Uri uri)
		: base(pck, ws, uri)
	{
		_pck = pck;
		_ws = ws;
		if (uri == null)
		{
			base.VmlDrawingXml.LoadXml(CreateVmlDrawings());
			_images = new List<ExcelVmlDrawingPicture>();
		}
		else
		{
			AddDrawingsFromXml();
		}
	}

	private void AddDrawingsFromXml()
	{
		XmlNodeList xmlNodeList = base.VmlDrawingXml.SelectNodes("//v:shape", base.NameSpaceManager);
		_images = new List<ExcelVmlDrawingPicture>();
		foreach (XmlNode item in xmlNodeList)
		{
			ExcelVmlDrawingPicture excelVmlDrawingPicture = new ExcelVmlDrawingPicture(item, base.NameSpaceManager, _ws);
			ZipPackageRelationship relationship = base.Part.GetRelationship(excelVmlDrawingPicture.RelId);
			excelVmlDrawingPicture.ImageUri = UriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
			_images.Add(excelVmlDrawingPicture);
		}
	}

	private string CreateVmlDrawings()
	{
		return string.Concat(string.Concat(string.Concat(string.Concat(string.Concat(string.Concat(string.Concat(string.Format("<xml xmlns:v=\"{0}\" xmlns:o=\"{1}\" xmlns:x=\"{2}\">", "urn:schemas-microsoft-com:vml", "urn:schemas-microsoft-com:office:office", "urn:schemas-microsoft-com:office:excel") + "<o:shapelayout v:ext=\"edit\">", "<o:idmap v:ext=\"edit\" data=\"1\"/>"), "</o:shapelayout>"), "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">"), "<v:stroke joinstyle=\"miter\" />"), "<v:path gradientshapeok=\"t\" o:connecttype=\"rect\" />"), "</v:shapetype>"), "</xml>");
	}

	internal ExcelVmlDrawingPicture Add(string id, Uri uri, string name, double width, double height)
	{
		ExcelVmlDrawingPicture excelVmlDrawingPicture = new ExcelVmlDrawingPicture(AddImage(id, uri, name, width, height), base.NameSpaceManager, _ws);
		excelVmlDrawingPicture.ImageUri = uri;
		_images.Add(excelVmlDrawingPicture);
		return excelVmlDrawingPicture;
	}

	private XmlNode AddImage(string id, Uri targeUri, string Name, double width, double height)
	{
		XmlElement xmlElement = base.VmlDrawingXml.CreateElement("v", "shape", "urn:schemas-microsoft-com:vml");
		base.VmlDrawingXml.DocumentElement.AppendChild(xmlElement);
		xmlElement.SetAttribute("id", id);
		xmlElement.SetAttribute("o:type", "#_x0000_t75");
		xmlElement.SetAttribute("style", $"position:absolute;margin-left:0;margin-top:0;width:{width.ToString(CultureInfo.InvariantCulture)}pt;height:{height.ToString(CultureInfo.InvariantCulture)}pt;z-index:1");
		xmlElement.InnerXml = $"<v:imagedata o:relid=\"\" o:title=\"{Name}\"/><o:lock v:ext=\"edit\" rotation=\"t\"/>";
		return xmlElement;
	}

	internal string GetNewId()
	{
		if (_nextID == 0)
		{
			foreach (ExcelVmlDrawingComment item in (IEnumerable)this)
			{
				if (item.Id.Length > 3 && item.Id.StartsWith("vml") && int.TryParse(item.Id.Substring(3, item.Id.Length - 3), NumberStyles.Number, CultureInfo.InvariantCulture, out var result) && result > _nextID)
				{
					_nextID = result;
				}
			}
		}
		_nextID++;
		return "vml" + _nextID;
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _images.GetEnumerator();
	}
}
