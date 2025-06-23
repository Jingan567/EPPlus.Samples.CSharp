using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using Ionic.Zip;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Packaging.Ionic.Zlib;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Packaging;

public class ZipPackage : ZipPackageRelationshipBase
{
	internal class ContentType
	{
		internal string Name;

		internal bool IsExtension;

		internal string Match;

		public ContentType(string name, bool isExtension, string match)
		{
			Name = name;
			IsExtension = isExtension;
			Match = match;
		}
	}

	private Dictionary<string, ZipPackagePart> Parts = new Dictionary<string, ZipPackagePart>(StringComparer.OrdinalIgnoreCase);

	internal Dictionary<string, ContentType> _contentTypes = new Dictionary<string, ContentType>(StringComparer.OrdinalIgnoreCase);

	internal char _dirSeparator = '/';

	private CompressionLevel _compression = CompressionLevel.Level6;

	public CompressionLevel Compression
	{
		get
		{
			return _compression;
		}
		set
		{
			foreach (ZipPackagePart value2 in Parts.Values)
			{
				if (value2.CompressionLevel == _compression)
				{
					value2.CompressionLevel = value;
				}
			}
			_compression = value;
		}
	}

	internal ZipPackage()
	{
		AddNew();
	}

    /// <summary>
    /// 增加默认的ContentTypes
    /// </summary>
    private void AddNew()
	{
		_contentTypes.Add("xml", new ContentType("application/xml", isExtension: true, "xml"));
		_contentTypes.Add("rels", new ContentType("application/vnd.openxmlformats-package.relationships+xml", isExtension: true, "rels"));
	}

	internal ZipPackage(Stream stream)
	{
		bool flag = false;
		if (stream == null || stream.Length == 0L)
		{
			AddNew();
			return;
		}
		Dictionary<string, string> dictionary = new Dictionary<string, string>();
		stream.Seek(0L, SeekOrigin.Begin);
		using ZipInputStream zipInputStream = new ZipInputStream(stream);
		ZipEntry nextEntry = zipInputStream.GetNextEntry();
		if (nextEntry.FileName.Contains("\\"))
		{
			_dirSeparator = '\\';
		}
		else
		{
			_dirSeparator = '/';
		}
		while (nextEntry != null)
		{
			if (nextEntry.UncompressedSize > 0)
			{
				byte[] array = new byte[nextEntry.UncompressedSize];
				zipInputStream.Read(array, 0, (int)nextEntry.UncompressedSize);
				if (nextEntry.FileName.Equals("[content_types].xml", StringComparison.OrdinalIgnoreCase))
				{
					AddContentTypes(Encoding.UTF8.GetString(array));
					flag = true;
				}
				else if (nextEntry.FileName.Equals($"_rels{_dirSeparator}.rels", StringComparison.OrdinalIgnoreCase))
				{
					ReadRelation(Encoding.UTF8.GetString(array), "");
				}
				else if (nextEntry.FileName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
				{
					dictionary.Add(GetUriKey(nextEntry.FileName), Encoding.UTF8.GetString(array));
				}
				else
				{
					ZipPackagePart zipPackagePart = new ZipPackagePart(this, nextEntry);
					zipPackagePart.Stream = new MemoryStream();
					zipPackagePart.Stream.Write(array, 0, array.Length);
					Parts.Add(GetUriKey(nextEntry.FileName), zipPackagePart);
				}
			}
			nextEntry = zipInputStream.GetNextEntry();
		}
		foreach (KeyValuePair<string, ZipPackagePart> part in Parts)
		{
			string fileName = Path.GetFileName(part.Key);
			string extension = Path.GetExtension(part.Key);
			string key = $"{part.Key.Substring(0, part.Key.Length - fileName.Length)}_rels/{fileName}.rels";
			if (dictionary.ContainsKey(key))
			{
				part.Value.ReadRelation(dictionary[key], part.Value.Uri.OriginalString);
			}
			if (_contentTypes.ContainsKey(part.Key))
			{
				part.Value.ContentType = _contentTypes[part.Key].Name;
			}
			else if (extension.Length > 1 && _contentTypes.ContainsKey(extension.Substring(1)))
			{
				part.Value.ContentType = _contentTypes[extension.Substring(1)].Name;
			}
		}
		if (!flag)
		{
			throw new InvalidDataException("The file is not an valid Package file. If the file is encrypted, please supply the password in the constructor.");
		}
		if (!flag)
		{
			throw new InvalidDataException("The file is not an valid Package file. If the file is encrypted, please supply the password in the constructor.");
		}
		zipInputStream.Close();
		zipInputStream.Dispose();
	}

	private void AddContentTypes(string xml)
	{
		XmlDocument xmlDocument = new XmlDocument();
		XmlHelper.LoadXmlSafe(xmlDocument, xml, Encoding.UTF8);
		foreach (XmlElement childNode in xmlDocument.DocumentElement.ChildNodes)
		{
			ContentType contentType = ((!string.IsNullOrEmpty(childNode.GetAttribute("Extension"))) ? new ContentType(childNode.GetAttribute("ContentType"), isExtension: true, childNode.GetAttribute("Extension")) : new ContentType(childNode.GetAttribute("ContentType"), isExtension: false, childNode.GetAttribute("PartName")));
			_contentTypes.Add(GetUriKey(contentType.Match), contentType);
		}
	}

	internal ZipPackagePart CreatePart(Uri partUri, string contentType)
	{
		return CreatePart(partUri, contentType, CompressionLevel.Level6);
	}

	internal ZipPackagePart CreatePart(Uri partUri, string contentType, CompressionLevel compressionLevel)
	{
		if (PartExists(partUri))
		{
			throw new InvalidOperationException("Part already exist");
		}
		ZipPackagePart zipPackagePart = new ZipPackagePart(this, partUri, contentType, compressionLevel);
		_contentTypes.Add(GetUriKey(zipPackagePart.Uri.OriginalString), new ContentType(contentType, isExtension: false, zipPackagePart.Uri.OriginalString));
		Parts.Add(GetUriKey(zipPackagePart.Uri.OriginalString), zipPackagePart);
		return zipPackagePart;
	}

	internal ZipPackagePart GetPart(Uri partUri)
	{
		if (PartExists(partUri))
		{
			return Parts.Single((KeyValuePair<string, ZipPackagePart> x) => x.Key.Equals(GetUriKey(partUri.OriginalString), StringComparison.OrdinalIgnoreCase)).Value;
		}
		throw new InvalidOperationException("Part does not exist.");
	}

	internal string GetUriKey(string uri)
	{
		string text = uri.Replace('\\', '/');
		if (text[0] != '/')
		{
			text = "/" + text;
		}
		return text;
	}

	internal bool PartExists(Uri partUri)
	{
		string uriKey = GetUriKey(partUri.OriginalString.ToLowerInvariant());
		return Parts.ContainsKey(uriKey);
	}

	internal void DeletePart(Uri Uri)
	{
		List<object[]> list = new List<object[]>();
		foreach (ZipPackagePart value in Parts.Values)
		{
			foreach (ZipPackageRelationship relationship in value.GetRelationships())
			{
				if (UriHelper.ResolvePartUri(value.Uri, relationship.TargetUri).OriginalString.Equals(Uri.OriginalString, StringComparison.OrdinalIgnoreCase))
				{
					list.Add(new object[2] { relationship.Id, value });
				}
			}
		}
		foreach (object[] item in list)
		{
			((ZipPackagePart)item[1]).DeleteRelationship(item[0].ToString());
		}
		ZipPackageRelationshipCollection relationships = GetPart(Uri).GetRelationships();
		while (relationships.Count > 0)
		{
			relationships.Remove(relationships.First().Id);
		}
		relationships = null;
		_contentTypes.Remove(GetUriKey(Uri.OriginalString));
		Parts.Remove(GetUriKey(Uri.OriginalString));
	}

	internal void Save(Stream stream)
	{
		Encoding uTF = Encoding.UTF8;
		ZipOutputStream zipOutputStream = new ZipOutputStream(stream, leaveOpen: true);
		zipOutputStream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)_compression;
		zipOutputStream.PutNextEntry("[Content_Types].xml");
		byte[] bytes = uTF.GetBytes(GetContentTypeXml());
		zipOutputStream.Write(bytes, 0, bytes.Length);
		_rels.WriteZip(zipOutputStream, "_rels/.rels");
		ZipPackagePart zipPackagePart = null;
		foreach (ZipPackagePart value in Parts.Values)
		{
			if (value.ContentType != "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
			{
				value.WriteZip(zipOutputStream);
			}
			else
			{
				zipPackagePart = value;
			}
		}
		zipPackagePart?.WriteZip(zipOutputStream);
		zipOutputStream.Flush();
		zipOutputStream.Close();
		zipOutputStream.Dispose();
	}

	private string GetContentTypeXml()
	{
		StringBuilder stringBuilder = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
		foreach (ContentType value in _contentTypes.Values)
		{
			if (value.IsExtension)
			{
				stringBuilder.AppendFormat("<Default ContentType=\"{0}\" Extension=\"{1}\"/>", value.Name, value.Match);
			}
			else
			{
				stringBuilder.AppendFormat("<Override ContentType=\"{0}\" PartName=\"{1}\" />", value.Name, GetUriKey(value.Match));
			}
		}
		stringBuilder.Append("</Types>");
		return stringBuilder.ToString();
	}

	internal void Flush()
	{
	}

	internal void Close()
	{
	}
}
