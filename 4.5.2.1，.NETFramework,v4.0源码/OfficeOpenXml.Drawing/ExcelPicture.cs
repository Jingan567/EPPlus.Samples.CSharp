using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing;

public sealed class ExcelPicture : ExcelDrawing
{
	private Image _image;

	private ImageFormat _imageFormat = ImageFormat.Jpeg;

	internal ZipPackagePart Part;

	private ExcelDrawingFill _fill;

	private ExcelDrawingBorder _border;

	private Uri _hyperlink;

	internal string ImageHash { get; set; }

	public Image Image
	{
		get
		{
			return _image;
		}
		set
		{
			if (value != null)
			{
				_image = value;
				try
				{
					string value2 = SavePicture(value);
					base.TopNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", base.NameSpaceManager).Value = value2;
				}
				catch (Exception ex)
				{
					throw new Exception("Can't save image - " + ex.Message, ex);
				}
			}
		}
	}

	public ImageFormat ImageFormat
	{
		get
		{
			return _imageFormat;
		}
		internal set
		{
			_imageFormat = value;
		}
	}

	internal string ContentType { get; set; }

	internal Uri UriPic { get; set; }

	internal ZipPackageRelationship RelPic { get; set; }

	internal ZipPackageRelationship HypRel { get; set; }

	internal new string Id => base.Name;

	public ExcelDrawingFill Fill
	{
		get
		{
			if (_fill == null)
			{
				_fill = new ExcelDrawingFill(base.NameSpaceManager, base.TopNode, "xdr:pic/xdr:spPr");
			}
			return _fill;
		}
	}

	public ExcelDrawingBorder Border
	{
		get
		{
			if (_border == null)
			{
				_border = new ExcelDrawingBorder(base.NameSpaceManager, base.TopNode, "xdr:pic/xdr:spPr/a:ln");
			}
			return _border;
		}
	}

	public Uri Hyperlink => _hyperlink;

	internal ExcelPicture(ExcelDrawings drawings, XmlNode node)
		: base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
	{
		XmlNode xmlNode = node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip", drawings.NameSpaceManager);
		if (xmlNode == null)
		{
			return;
		}
		RelPic = drawings.Part.GetRelationship(xmlNode.Attributes["r:embed"].Value);
		UriPic = UriHelper.ResolvePartUri(drawings.UriDrawing, RelPic.TargetUri);
		Part = drawings.Part.Package.GetPart(UriPic);
		FileInfo fileInfo = new FileInfo(UriPic.OriginalString);
		ContentType = GetContentType(fileInfo.Extension);
		_image = Image.FromStream(Part.GetStream());
		byte[] image = (byte[])new ImageConverter().ConvertTo(_image, typeof(byte[]));
		ExcelPackage.ImageInfo imageInfo = _drawings._package.LoadImage(image, UriPic, Part);
		ImageHash = imageInfo.Hash;
		string xmlNodeString = GetXmlNodeString("xdr:pic/xdr:nvPicPr/xdr:cNvPr/a:hlinkClick/@r:id");
		if (!string.IsNullOrEmpty(xmlNodeString))
		{
			HypRel = drawings.Part.GetRelationship(xmlNodeString);
			if (HypRel.TargetUri.IsAbsoluteUri)
			{
				_hyperlink = new ExcelHyperLink(HypRel.TargetUri.AbsoluteUri);
			}
			else
			{
				_hyperlink = new ExcelHyperLink(HypRel.TargetUri.OriginalString, UriKind.Relative);
			}
			((ExcelHyperLink)_hyperlink).ToolTip = GetXmlNodeString("xdr:pic/xdr:nvPicPr/xdr:cNvPr/a:hlinkClick/@tooltip");
		}
	}

	internal ExcelPicture(ExcelDrawings drawings, XmlNode node, Image image, Uri hyperlink)
		: base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
	{
		XmlElement xmlElement = node.OwnerDocument.CreateElement("xdr", "pic", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
		node.InsertAfter(xmlElement, node.SelectSingleNode("xdr:to", base.NameSpaceManager));
		_hyperlink = hyperlink;
		xmlElement.InnerXml = PicStartXml();
		node.InsertAfter(node.OwnerDocument.CreateElement("xdr", "clientData", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"), xmlElement);
		ZipPackage package = drawings.Worksheet._package.Package;
		_image = image;
		string value = SavePicture(image);
		node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", base.NameSpaceManager).Value = value;
		_height = image.Height;
		_width = image.Width;
		SetPosDefaults(image);
		package.Flush();
	}

	internal ExcelPicture(ExcelDrawings drawings, XmlNode node, FileInfo imageFile, Uri hyperlink)
		: base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
	{
		XmlElement xmlElement = node.OwnerDocument.CreateElement("xdr", "pic", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
		node.InsertAfter(xmlElement, node.SelectSingleNode("xdr:to", base.NameSpaceManager));
		_hyperlink = hyperlink;
		xmlElement.InnerXml = PicStartXml();
		node.InsertAfter(node.OwnerDocument.CreateElement("xdr", "clientData", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"), xmlElement);
		ZipPackage package = drawings.Worksheet._package.Package;
		ContentType = GetContentType(imageFile.Extension);
		FileStream fileStream = new FileStream(imageFile.FullName, FileMode.Open, FileAccess.Read);
		_image = Image.FromStream(fileStream);
		byte[] array = (byte[])new ImageConverter().ConvertTo(_image, typeof(byte[]));
		fileStream.Close();
		UriPic = XmlHelper.GetNewUri(package, "/xl/media/{0}" + imageFile.Name);
		ExcelPackage.ImageInfo imageInfo = _drawings._package.AddImage(array, UriPic, ContentType);
		string text;
		if (!drawings._hashes.ContainsKey(imageInfo.Hash))
		{
			Part = imageInfo.Part;
			RelPic = drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(drawings.UriDrawing, imageInfo.Uri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
			text = RelPic.Id;
			_drawings._hashes.Add(imageInfo.Hash, text);
			AddNewPicture(array, text);
		}
		else
		{
			text = drawings._hashes[imageInfo.Hash];
			ZipPackageRelationship relationship = _drawings.Part.GetRelationship(text);
			UriPic = UriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
		}
		ImageHash = imageInfo.Hash;
		_height = Image.Height;
		_width = Image.Width;
		SetPosDefaults(Image);
		node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", base.NameSpaceManager).Value = text;
		package.Flush();
	}

	internal static string GetContentType(string extension)
	{
		switch (extension.ToLower(CultureInfo.InvariantCulture))
		{
		case ".bmp":
			return "image/bmp";
		case ".jpg":
		case ".jpeg":
			return "image/jpeg";
		case ".gif":
			return "image/gif";
		case ".png":
			return "image/png";
		case ".cgm":
			return "image/cgm";
		case ".emf":
			return "image/x-emf";
		case ".eps":
			return "image/x-eps";
		case ".pcx":
			return "image/x-pcx";
		case ".tga":
			return "image/x-tga";
		case ".tif":
		case ".tiff":
			return "image/x-tiff";
		case ".wmf":
			return "image/x-wmf";
		default:
			return "image/jpeg";
		}
	}

	internal static ImageFormat GetImageFormat(string contentType)
	{
		return contentType.ToLower(CultureInfo.InvariantCulture) switch
		{
			"image/bmp" => ImageFormat.Bmp, 
			"image/jpeg" => ImageFormat.Jpeg, 
			"image/gif" => ImageFormat.Gif, 
			"image/png" => ImageFormat.Png, 
			"image/x-emf" => ImageFormat.Emf, 
			"image/x-tiff" => ImageFormat.Tiff, 
			"image/x-wmf" => ImageFormat.Wmf, 
			_ => ImageFormat.Jpeg, 
		};
	}

	private void AddNewPicture(byte[] img, string relID)
	{
		new ExcelDrawings.ImageCompare
		{
			image = img,
			relID = relID
		};
	}

	private string SavePicture(Image image)
	{
		byte[] image2 = (byte[])new ImageConverter().ConvertTo(image, typeof(byte[]));
		ExcelPackage.ImageInfo imageInfo = _drawings._package.AddImage(image2);
		ImageHash = imageInfo.Hash;
		if (_drawings._hashes.ContainsKey(imageInfo.Hash))
		{
			string text = _drawings._hashes[imageInfo.Hash];
			ZipPackageRelationship relationship = _drawings.Part.GetRelationship(text);
			UriPic = UriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
			return text;
		}
		UriPic = imageInfo.Uri;
		ImageHash = imageInfo.Hash;
		RelPic = _drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(_drawings.UriDrawing, UriPic), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
		_drawings._hashes.Add(imageInfo.Hash, RelPic.Id);
		return RelPic.Id;
	}

	private void SetPosDefaults(Image image)
	{
		base.EditAs = eEditAs.OneCell;
		SetPixelWidth(image.Width, image.HorizontalResolution);
		SetPixelHeight(image.Height, image.VerticalResolution);
	}

	private string PicStartXml()
	{
		StringBuilder stringBuilder = new StringBuilder();
		stringBuilder.Append("<xdr:nvPicPr>");
		if (_hyperlink == null)
		{
			stringBuilder.AppendFormat("<xdr:cNvPr id=\"{0}\" descr=\"\" />", _id);
		}
		else
		{
			HypRel = _drawings.Part.CreateRelationship(_hyperlink, TargetMode.External, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
			stringBuilder.AppendFormat("<xdr:cNvPr id=\"{0}\" descr=\"\">", _id);
			if (HypRel != null)
			{
				if (_hyperlink is ExcelHyperLink)
				{
					stringBuilder.AppendFormat("<a:hlinkClick xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{0}\" tooltip=\"{1}\"/>", HypRel.Id, ((ExcelHyperLink)_hyperlink).ToolTip);
				}
				else
				{
					stringBuilder.AppendFormat("<a:hlinkClick xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{0}\" />", HypRel.Id);
				}
			}
			stringBuilder.Append("</xdr:cNvPr>");
		}
		stringBuilder.Append("<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"\" cstate=\"print\" /><a:stretch><a:fillRect /> </a:stretch> </xdr:blipFill> <xdr:spPr> <a:xfrm> <a:off x=\"0\" y=\"0\" />  <a:ext cx=\"0\" cy=\"0\" /> </a:xfrm> <a:prstGeom prst=\"rect\"> <a:avLst /> </a:prstGeom> </xdr:spPr>");
		return stringBuilder.ToString();
	}

	public override void SetSize(int Percent)
	{
		if (Image == null)
		{
			base.SetSize(Percent);
			return;
		}
		_width = Image.Width;
		_height = Image.Height;
		_width = (int)((decimal)_width * ((decimal)Percent / 100m));
		_height = (int)((decimal)_height * ((decimal)Percent / 100m));
		SetPixelWidth(_width, Image.HorizontalResolution);
		SetPixelHeight(_height, Image.VerticalResolution);
	}

	internal override void DeleteMe()
	{
		_drawings._package.RemoveImage(ImageHash);
		base.DeleteMe();
	}

	public override void Dispose()
	{
		base.Dispose();
		_hyperlink = null;
		_image.Dispose();
		_image = null;
	}
}
