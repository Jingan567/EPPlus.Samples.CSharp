using System;
using System.Drawing;
using System.IO;
using System.Xml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml;

public class ExcelBackgroundImage : XmlHelper
{
	private ExcelWorksheet _workSheet;

	private const string BACKGROUNDPIC_PATH = "d:picture/@r:id";

	public Image Image
	{
		get
		{
			string xmlNodeString = GetXmlNodeString("d:picture/@r:id");
			if (!string.IsNullOrEmpty(xmlNodeString))
			{
				ZipPackageRelationship relationship = _workSheet.Part.GetRelationship(xmlNodeString);
				return Image.FromStream(_workSheet.Part.Package.GetPart(UriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri)).GetStream());
			}
			return null;
		}
		set
		{
			DeletePrevImage();
			if (value == null)
			{
				DeleteAllNode("d:picture/@r:id");
				return;
			}
			byte[] image = (byte[])new ImageConverter().ConvertTo(value, typeof(byte[]));
			ExcelPackage.ImageInfo imageInfo = _workSheet.Workbook._package.AddImage(image);
			ZipPackageRelationship zipPackageRelationship = _workSheet.Part.CreateRelationship(imageInfo.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
			SetXmlNodeString("d:picture/@r:id", zipPackageRelationship.Id);
		}
	}

	internal ExcelBackgroundImage(XmlNamespaceManager nsm, XmlNode topNode, ExcelWorksheet workSheet)
		: base(nsm, topNode)
	{
		_workSheet = workSheet;
	}

	public void SetFromFile(FileInfo PictureFile)
	{
		DeletePrevImage();
		byte[] array;
		try
		{
			array = File.ReadAllBytes(PictureFile.FullName);
			Image.FromFile(PictureFile.FullName);
		}
		catch (Exception innerException)
		{
			throw new InvalidDataException("File is not a supported image-file or is corrupt", innerException);
		}
		string contentType = ExcelPicture.GetContentType(PictureFile.Extension);
		Uri newUri = XmlHelper.GetNewUri(_workSheet._package.Package, "/xl/media/" + PictureFile.Name.Substring(0, PictureFile.Name.Length - PictureFile.Extension.Length) + "{0}" + PictureFile.Extension);
		ExcelPackage.ImageInfo imageInfo = _workSheet.Workbook._package.AddImage(array, newUri, contentType);
		if (_workSheet.Part.Package.PartExists(newUri) && imageInfo.RefCount == 1)
		{
			_workSheet.Part.Package.DeletePart(newUri);
		}
		_workSheet.Part.Package.CreatePart(newUri, contentType, CompressionLevel.Level0).GetStream(FileMode.Create, FileAccess.Write).Write(array, 0, array.Length);
		ZipPackageRelationship zipPackageRelationship = _workSheet.Part.CreateRelationship(newUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
		SetXmlNodeString("d:picture/@r:id", zipPackageRelationship.Id);
	}

	private void DeletePrevImage()
	{
		string xmlNodeString = GetXmlNodeString("d:picture/@r:id");
		if (xmlNodeString != "")
		{
			byte[] image = (byte[])new ImageConverter().ConvertTo(Image, typeof(byte[]));
			ExcelPackage.ImageInfo imageInfo = _workSheet.Workbook._package.GetImageInfo(image);
			_workSheet.Part.DeleteRelationship(xmlNodeString);
			if (imageInfo != null && imageInfo.RefCount == 1 && _workSheet.Part.Package.PartExists(imageInfo.Uri))
			{
				_workSheet.Part.Package.DeletePart(imageInfo.Uri);
			}
		}
	}
}
