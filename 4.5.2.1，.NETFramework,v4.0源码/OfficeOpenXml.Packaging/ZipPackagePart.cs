using System;
using System.IO;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Packaging.Ionic.Zlib;

namespace OfficeOpenXml.Packaging;

internal class ZipPackagePart : ZipPackageRelationshipBase, IDisposable
{
	internal delegate void SaveHandlerDelegate(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName);

	internal CompressionLevel CompressionLevel;

	private MemoryStream _stream;

	private string _contentType = "";

	internal ZipPackage Package { get; set; }

	internal ZipEntry Entry { get; set; }

	internal MemoryStream Stream
	{
		get
		{
			return _stream;
		}
		set
		{
			_stream = value;
		}
	}

	public string ContentType
	{
		get
		{
			return _contentType;
		}
		internal set
		{
			if (!string.IsNullOrEmpty(_contentType) && Package._contentTypes.ContainsKey(Package.GetUriKey(Uri.OriginalString)))
			{
				Package._contentTypes.Remove(Package.GetUriKey(Uri.OriginalString));
				Package._contentTypes.Add(Package.GetUriKey(Uri.OriginalString), new ZipPackage.ContentType(value, isExtension: false, Uri.OriginalString));
			}
			_contentType = value;
		}
	}

	public Uri Uri { get; private set; }

	internal SaveHandlerDelegate SaveHandler { get; set; }

	internal ZipPackagePart(ZipPackage package, ZipEntry entry)
	{
		Package = package;
		Entry = entry;
		SaveHandler = null;
		Uri = new Uri(package.GetUriKey(entry.FileName), UriKind.Relative);
	}

	internal ZipPackagePart(ZipPackage package, Uri partUri, string contentType, CompressionLevel compressionLevel)
	{
		Package = package;
		Uri = partUri;
		ContentType = contentType;
		CompressionLevel = compressionLevel;
	}

	internal override ZipPackageRelationship CreateRelationship(Uri targetUri, TargetMode targetMode, string relationshipType)
	{
		ZipPackageRelationship zipPackageRelationship = base.CreateRelationship(targetUri, targetMode, relationshipType);
		zipPackageRelationship.SourceUri = Uri;
		return zipPackageRelationship;
	}

	internal MemoryStream GetStream()
	{
		return GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite);
	}

	internal MemoryStream GetStream(FileMode fileMode)
	{
		return GetStream(FileMode.Create, FileAccess.ReadWrite);
	}

	internal MemoryStream GetStream(FileMode fileMode, FileAccess fileAccess)
	{
		if (_stream == null || fileMode == FileMode.CreateNew || fileMode == FileMode.Create)
		{
			_stream = new MemoryStream();
		}
		else
		{
			_stream.Seek(0L, SeekOrigin.Begin);
		}
		return _stream;
	}

	public Stream GetZipStream()
	{
		return new ZipOutputStream(new MemoryStream());
	}

	internal void WriteZip(ZipOutputStream os)
	{
		byte[] array;
		if (SaveHandler == null)
		{
			array = GetStream().ToArray();
			if (array.Length == 0)
			{
				return;
			}
			os.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)CompressionLevel;
			os.PutNextEntry(Uri.OriginalString);
			os.Write(array, 0, array.Length);
		}
		else
		{
			SaveHandler(os, CompressionLevel, Uri.OriginalString);
		}
		if (_rels.Count > 0)
		{
			string originalString = Uri.OriginalString;
			string fileName = Path.GetFileName(originalString);
			_rels.WriteZip(os, $"{originalString.Substring(0, originalString.Length - fileName.Length)}_rels/{fileName}.rels");
		}
		array = null;
	}

	public void Dispose()
	{
		_stream.Close();
		_stream.Dispose();
	}
}
