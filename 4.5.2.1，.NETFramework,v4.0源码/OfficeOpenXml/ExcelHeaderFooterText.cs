using System;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Xml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Vml;

namespace OfficeOpenXml;

public class ExcelHeaderFooterText
{
	private ExcelWorksheet _ws;

	private string _hf;

	public string LeftAlignedText;

	public string CenteredText;

	public string RightAlignedText;

	internal ExcelHeaderFooterText(XmlNode TextNode, ExcelWorksheet ws, string hf)
	{
		_ws = ws;
		_hf = hf;
		if (TextNode == null || string.IsNullOrEmpty(TextNode.InnerText))
		{
			return;
		}
		string innerText = TextNode.InnerText;
		string code = innerText.Substring(0, 2);
		int num = 2;
		for (int i = num; i < innerText.Length - 2; i++)
		{
			string text = innerText.Substring(i, 2);
			if (text == "&C" || text == "&R")
			{
				SetText(code, innerText.Substring(num, i - num));
				num = i + 2;
				i = num;
				code = text;
			}
		}
		SetText(code, innerText.Substring(num, innerText.Length - num));
	}

	private void SetText(string code, string text)
	{
		if (!(code == "&L"))
		{
			if (code == "&C")
			{
				CenteredText = text;
			}
			else
			{
				RightAlignedText = text;
			}
		}
		else
		{
			LeftAlignedText = text;
		}
	}

	public ExcelVmlDrawingPicture InsertPicture(Image Picture, PictureAlignment Alignment)
	{
		string id = ValidateImage(Alignment);
		byte[] image = (byte[])new ImageConverter().ConvertTo(Picture, typeof(byte[]));
		ExcelPackage.ImageInfo ii = _ws.Workbook._package.AddImage(image);
		return AddImage(Picture, id, ii);
	}

	public ExcelVmlDrawingPicture InsertPicture(FileInfo PictureFile, PictureAlignment Alignment)
	{
		string id = ValidateImage(Alignment);
		Image image;
		try
		{
			if (!PictureFile.Exists)
			{
				throw new FileNotFoundException($"{PictureFile.FullName} is missing");
			}
			image = Image.FromFile(PictureFile.FullName);
		}
		catch (Exception innerException)
		{
			throw new InvalidDataException("File is not a supported image-file or is corrupt", innerException);
		}
		string contentType = ExcelPicture.GetContentType(PictureFile.Extension);
		Uri newUri = XmlHelper.GetNewUri(_ws._package.Package, "/xl/media/" + PictureFile.Name.Substring(0, PictureFile.Name.Length - PictureFile.Extension.Length) + "{0}" + PictureFile.Extension);
		byte[] image2 = (byte[])new ImageConverter().ConvertTo(image, typeof(byte[]));
		ExcelPackage.ImageInfo ii = _ws.Workbook._package.AddImage(image2, newUri, contentType);
		return AddImage(image, id, ii);
	}

	private ExcelVmlDrawingPicture AddImage(Image Picture, string id, ExcelPackage.ImageInfo ii)
	{
		double width = (float)(Picture.Width * 72) / Picture.HorizontalResolution;
		double height = (float)(Picture.Height * 72) / Picture.VerticalResolution;
		return _ws.HeaderFooter.Pictures.Add(id, ii.Uri, "", width, height);
	}

	private string ValidateImage(PictureAlignment Alignment)
	{
		string text = Alignment.ToString()[0] + _hf;
		foreach (ExcelVmlDrawingPicture item in (IEnumerable)_ws.HeaderFooter.Pictures)
		{
			if (item.Id == text)
			{
				throw new InvalidOperationException("A picture already exists in this section");
			}
		}
		switch (Alignment)
		{
		case PictureAlignment.Left:
			LeftAlignedText += "&G";
			break;
		case PictureAlignment.Centered:
			CenteredText += "&G";
			break;
		default:
			RightAlignedText += "&G";
			break;
		}
		return text;
	}
}
