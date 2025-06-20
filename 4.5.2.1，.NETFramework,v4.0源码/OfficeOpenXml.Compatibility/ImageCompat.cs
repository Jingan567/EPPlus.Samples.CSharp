using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace OfficeOpenXml.Compatibility;

internal class ImageCompat
{
	internal static byte[] GetImageAsByteArray(Image image)
	{
		MemoryStream memoryStream = new MemoryStream();
		if (image.RawFormat.Guid == ImageFormat.Gif.Guid)
		{
			image.Save(memoryStream, ImageFormat.Gif);
		}
		else if (image.RawFormat.Guid == ImageFormat.Bmp.Guid)
		{
			image.Save(memoryStream, ImageFormat.Bmp);
		}
		else if (image.RawFormat.Guid == ImageFormat.Png.Guid)
		{
			image.Save(memoryStream, ImageFormat.Png);
		}
		else if (image.RawFormat.Guid == ImageFormat.Tiff.Guid)
		{
			image.Save(memoryStream, ImageFormat.Tiff);
		}
		else
		{
			image.Save(memoryStream, ImageFormat.Jpeg);
		}
		return memoryStream.ToArray();
	}
}
