using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Encryption;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.CompundDocument;

namespace OfficeOpenXml;

public sealed class ExcelPackage : IDisposable
{
	internal class ImageInfo
	{
		internal string Hash { get; set; }

		internal Uri Uri { get; set; }

		internal int RefCount { get; set; }

		internal ZipPackagePart Part { get; set; }
	}

	internal const bool preserveWhitespace = false;

	private Stream _stream;

	private bool _isExternalStream;

	internal Dictionary<string, ImageInfo> _images = new Dictionary<string, ImageInfo>();

	internal const string schemaXmlExtension = "application/xml";

	internal const string schemaRelsExtension = "application/vnd.openxmlformats-package.relationships+xml";

	internal const string schemaMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

	internal const string schemaRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

	internal const string schemaDrawings = "http://schemas.openxmlformats.org/drawingml/2006/main";

	internal const string schemaSheetDrawings = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

	internal const string schemaMarkupCompatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";

	internal const string schemaMicrosoftVml = "urn:schemas-microsoft-com:vml";

	internal const string schemaMicrosoftOffice = "urn:schemas-microsoft-com:office:office";

	internal const string schemaMicrosoftExcel = "urn:schemas-microsoft-com:office:excel";

	internal const string schemaChart = "http://schemas.openxmlformats.org/drawingml/2006/chart";

	internal const string schemaHyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

	internal const string schemaComment = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

	internal const string schemaImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

	internal const string schemaCore = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

	internal const string schemaExtended = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

	internal const string schemaCustom = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

	internal const string schemaDc = "http://purl.org/dc/elements/1.1/";

	internal const string schemaDcTerms = "http://purl.org/dc/terms/";

	internal const string schemaDcmiType = "http://purl.org/dc/dcmitype/";

	internal const string schemaXsi = "http://www.w3.org/2001/XMLSchema-instance";

	internal const string schemaVt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

	internal const string schemaMainX14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

	internal const string schemaMainXm = "http://schemas.microsoft.com/office/excel/2006/main";

	internal const string schemaXr = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";

	internal const string schemaXr2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";

	internal const string schemaPivotTable = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";

	internal const string schemaPivotCacheDefinition = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";

	internal const string schemaPivotCacheRecords = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";

	internal const string schemaVBA = "application/vnd.ms-office.vbaProject";

	internal const string schemaVBASignature = "application/vnd.ms-office.vbaProjectSignature";

	internal const string contentTypeWorkbookDefault = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

	internal const string contentTypeWorkbookMacroEnabled = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";

	internal const string contentTypeSharedString = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";

	private ZipPackage _package;

	internal ExcelWorkbook _workbook;

	public const int MaxColumns = 16384;

	public const int MaxRows = 1048576;

	internal static int _id = 1;

	private ExcelEncryption _encryption;

	private FileInfo _file;

	private CompatibilitySettings _compatibility;

	private static object _lock = new object();

	internal int _worksheetAdd = 1;

	public ZipPackage Package => _package;

	public ExcelEncryption Encryption
	{
		get
		{
			if (_encryption == null)
			{
				_encryption = new ExcelEncryption();
			}
			return _encryption;
		}
	}

	public ExcelWorkbook Workbook
	{
		get
		{
			if (_workbook == null)
			{
				XmlNamespaceManager namespaceManager = CreateDefaultNSM();
				_workbook = new ExcelWorkbook(this, namespaceManager);
				_workbook.GetExternalReferences();
				_workbook.GetDefinedNames();
			}
			return _workbook;
		}
	}

	public bool DoAdjustDrawings { get; set; }

	public FileInfo File
	{
		get
		{
			return _file;
		}
		set
		{
			_file = value;
		}
	}

	public Stream Stream => _stream;

	public CompressionLevel Compression
	{
		get
		{
			return Package.Compression;
		}
		set
		{
			Package.Compression = value;
		}
	}

	public CompatibilitySettings Compatibility
	{
		get
		{
			if (_compatibility == null)
			{
				_compatibility = new CompatibilitySettings(this);
			}
			return _compatibility;
		}
	}

	public ExcelPackage()
	{
		Init();
		ConstructNewFile(null);
	}

	public ExcelPackage(FileInfo newFile)
	{
		Init();
		File = newFile;
		ConstructNewFile(null);
	}

	public ExcelPackage(FileInfo newFile, string password)
	{
		Init();
		File = newFile;
		ConstructNewFile(password);
	}

	public ExcelPackage(FileInfo newFile, FileInfo template)
	{
		Init();
		File = newFile;
		CreateFromTemplate(template, null);
	}

	public ExcelPackage(FileInfo newFile, FileInfo template, string password)
	{
		Init();
		File = newFile;
		CreateFromTemplate(template, password);
	}

	public ExcelPackage(FileInfo template, bool useStream)
	{
		Init();
		CreateFromTemplate(template, null);
		if (!useStream)
		{
			File = new FileInfo(Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx");
		}
	}

	public ExcelPackage(FileInfo template, bool useStream, string password)
	{
		Init();
		CreateFromTemplate(template, password);
		if (!useStream)
		{
			File = new FileInfo(Path.GetTempPath() + Guid.NewGuid().ToString() + ".xlsx");
		}
	}

	public ExcelPackage(Stream newStream)
	{
		Init();
		if (newStream.Length == 0L)
		{
			_stream = newStream;
			_isExternalStream = true;
			ConstructNewFile(null);
		}
		else
		{
			Load(newStream);
		}
	}

	public ExcelPackage(Stream newStream, string Password)
	{
		if (!newStream.CanRead || !newStream.CanWrite)
		{
			throw new Exception("The stream must be read/write");
		}
		Init();
		if (newStream.Length > 0)
		{
			Load(newStream, Password);
			return;
		}
		_stream = newStream;
		_isExternalStream = true;
		_package = new ZipPackage(_stream);
		CreateBlankWb();
	}

	public ExcelPackage(Stream newStream, Stream templateStream)
	{
		if (newStream.Length > 0)
		{
			throw new Exception("The output stream must be empty. Length > 0");
		}
		if (!newStream.CanRead || !newStream.CanWrite)
		{
			throw new Exception("The stream must be read/write");
		}
		Init();
		Load(templateStream, newStream, null);
	}

	public ExcelPackage(Stream newStream, Stream templateStream, string Password)
	{
		if (newStream.Length > 0)
		{
			throw new Exception("The output stream must be empty. Length > 0");
		}
		if (!newStream.CanRead || !newStream.CanWrite)
		{
			throw new Exception("The stream must be read/write");
		}
		Init();
		Load(templateStream, newStream, Password);
	}

	internal ImageInfo AddImage(byte[] image)
	{
		return AddImage(image, null, "");
	}

	internal ImageInfo AddImage(byte[] image, Uri uri, string contentType)
	{
		string text = BitConverter.ToString(new SHA1CryptoServiceProvider().ComputeHash(image)).Replace("-", "");
		lock (_images)
		{
			if (_images.ContainsKey(text))
			{
				_images[text].RefCount++;
			}
			else
			{
				ZipPackagePart zipPackagePart;
				if (uri == null)
				{
					uri = GetNewUri(Package, "/xl/media/image{0}.jpg");
					zipPackagePart = Package.CreatePart(uri, "image/jpeg", CompressionLevel.Level0);
				}
				else
				{
					zipPackagePart = Package.CreatePart(uri, contentType, CompressionLevel.Level0);
				}
				zipPackagePart.GetStream(FileMode.Create, FileAccess.Write).Write(image, 0, image.GetLength(0));
				_images.Add(text, new ImageInfo
				{
					Uri = uri,
					RefCount = 1,
					Hash = text,
					Part = zipPackagePart
				});
			}
		}
		return _images[text];
	}

	internal ImageInfo LoadImage(byte[] image, Uri uri, ZipPackagePart imagePart)
	{
		string text = BitConverter.ToString(new SHA1CryptoServiceProvider().ComputeHash(image)).Replace("-", "");
		if (_images.ContainsKey(text))
		{
			_images[text].RefCount++;
		}
		else
		{
			_images.Add(text, new ImageInfo
			{
				Uri = uri,
				RefCount = 1,
				Hash = text,
				Part = imagePart
			});
		}
		return _images[text];
	}

	internal void RemoveImage(string hash)
	{
		lock (_images)
		{
			if (_images.ContainsKey(hash))
			{
				ImageInfo imageInfo = _images[hash];
				imageInfo.RefCount--;
				if (imageInfo.RefCount == 0)
				{
					Package.DeletePart(imageInfo.Uri);
					_images.Remove(hash);
				}
			}
		}
	}

	internal ImageInfo GetImageInfo(byte[] image)
	{
		string key = BitConverter.ToString(new SHA1CryptoServiceProvider().ComputeHash(image)).Replace("-", "");
		if (_images.ContainsKey(key))
		{
			return _images[key];
		}
		return null;
	}

	private Uri GetNewUri(ZipPackage package, string sUri)
	{
		Uri uri;
		do
		{
			uri = new Uri(string.Format(sUri, _id++), UriKind.Relative);
		}
		while (package.PartExists(uri));
		return uri;
	}

	private void Init()
	{
		DoAdjustDrawings = true;
		string text = ConfigurationManager.AppSettings["EPPlus:ExcelPackage.Compatibility.IsWorksheets1Based"];
		if (text != null && bool.TryParse(text.ToLowerInvariant(), out var result))
		{
			Compatibility.IsWorksheets1Based = result;
		}
	}

	private void CreateFromTemplate(FileInfo template, string password)
	{
		template?.Refresh();
		if (template.Exists)
		{
			if (_stream == null)
			{
				_stream = new MemoryStream();
			}
			MemoryStream memoryStream = new MemoryStream();
			if (password != null)
			{
				Encryption.IsEncrypted = true;
				Encryption.Password = password;
				memoryStream = new EncryptedPackageHandler().DecryptPackage(template, Encryption);
			}
			else
			{
				WriteFileToStream(template.FullName, memoryStream);
			}
			try
			{
				_package = new ZipPackage(memoryStream);
				return;
			}
			catch (Exception innerException)
			{
				if (password == null && CompoundDocument.IsCompoundDocument(memoryStream))
				{
					throw new Exception("Can not open the package. Package is an OLE compound document. If this is an encrypted package, please supply the password", innerException);
				}
				throw;
			}
		}
		throw new Exception("Passed invalid TemplatePath to Excel Template");
	}

	private void ConstructNewFile(string password)
	{
		MemoryStream stream = new MemoryStream();
		if (_stream == null)
		{
			_stream = new MemoryStream();
		}
		if (File != null)
		{
			File.Refresh();
		}
		if (File != null && File.Exists)
		{
			if (password != null)
			{
				EncryptedPackageHandler encryptedPackageHandler = new EncryptedPackageHandler();
				Encryption.IsEncrypted = true;
				Encryption.Password = password;
				stream = encryptedPackageHandler.DecryptPackage(File, Encryption);
			}
			else
			{
				WriteFileToStream(File.FullName, stream);
			}
			try
			{
				_package = new ZipPackage(stream);
				return;
			}
			catch (Exception innerException)
			{
				if (password == null && CompoundDocument.IsCompoundDocument(File))
				{
					throw new Exception("Can not open the package. Package is an OLE compound document. If this is an encrypted package, please supply the password", innerException);
				}
				throw;
			}
		}
		_package = new ZipPackage(stream);
		CreateBlankWb();
	}

	private static void WriteFileToStream(string path, Stream stream)
	{
		using FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
		byte[] array = new byte[4096];
		int count;
		while ((count = fileStream.Read(array, 0, array.Length)) > 0)
		{
			stream.Write(array, 0, count);
		}
	}

	private void CreateBlankWb()
	{
		_ = Workbook.WorkbookXml;
		_package.CreateRelationship(UriHelper.GetRelativeUri(new Uri("/xl", UriKind.Relative), Workbook.WorkbookUri), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
	}

	private XmlNamespaceManager CreateDefaultNSM()
	{
		XmlNamespaceManager xmlNamespaceManager = new XmlNamespaceManager(new NameTable());
		xmlNamespaceManager.AddNamespace(string.Empty, "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlNamespaceManager.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		xmlNamespaceManager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
		xmlNamespaceManager.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		xmlNamespaceManager.AddNamespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
		xmlNamespaceManager.AddNamespace("xp", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
		xmlNamespaceManager.AddNamespace("ctp", "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties");
		xmlNamespaceManager.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
		xmlNamespaceManager.AddNamespace("dc", "http://purl.org/dc/elements/1.1/");
		xmlNamespaceManager.AddNamespace("dcterms", "http://purl.org/dc/terms/");
		xmlNamespaceManager.AddNamespace("dcmitype", "http://purl.org/dc/dcmitype/");
		xmlNamespaceManager.AddNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance");
		xmlNamespaceManager.AddNamespace("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
		xmlNamespaceManager.AddNamespace("xm", "http://schemas.microsoft.com/office/excel/2006/main");
		xmlNamespaceManager.AddNamespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
		return xmlNamespaceManager;
	}

	internal void SavePart(Uri uri, XmlDocument xmlDoc)
	{
		XmlTextWriter xmlTextWriter = new XmlTextWriter(_package.GetPart(uri).GetStream(FileMode.Create, FileAccess.Write), Encoding.UTF8);
		xmlTextWriter.Formatting = Formatting.None;
		xmlDoc.Save(xmlTextWriter);
	}

	internal void SaveWorkbook(Uri uri, XmlDocument xmlDoc)
	{
		ZipPackagePart zipPackagePart = _package.GetPart(uri);
		if (Workbook.VbaProject == null)
		{
			if (zipPackagePart.ContentType != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
			{
				zipPackagePart = _package.CreatePart(uri, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", Compression);
			}
		}
		else if (zipPackagePart.ContentType != "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
		{
			ZipPackageRelationshipCollection relationships = zipPackagePart.GetRelationships();
			_package.DeletePart(uri);
			zipPackagePart = Package.CreatePart(uri, "application/vnd.ms-excel.sheet.macroEnabled.main+xml");
			foreach (ZipPackageRelationship item in relationships)
			{
				Package.DeleteRelationship(item.Id);
				zipPackagePart.CreateRelationship(item.TargetUri, item.TargetMode, item.RelationshipType);
			}
		}
		XmlTextWriter xmlTextWriter = new XmlTextWriter(zipPackagePart.GetStream(FileMode.Create, FileAccess.Write), Encoding.UTF8);
		xmlTextWriter.Formatting = Formatting.None;
		xmlDoc.Save(xmlTextWriter);
	}

	public void Dispose()
	{
		if (_package != null)
		{
			if (!_isExternalStream && _stream != null && (_stream.CanRead || _stream.CanWrite))
			{
				CloseStream();
			}
			_package.Close();
			if (_workbook != null)
			{
				_workbook.Dispose();
			}
			_package = null;
			_images = null;
			_file = null;
			_workbook = null;
			_stream = null;
			_workbook = null;
			GC.Collect();
		}
	}

	public void Save()
	{
		try
		{
			if (_stream is MemoryStream && _stream.Length > 0)
			{
				CloseStream();
			}
			Workbook.Save();
			if (File == null)
			{
				if (Encryption.IsEncrypted)
				{
					MemoryStream memoryStream = new MemoryStream();
					_package.Save(memoryStream);
					byte[] package = memoryStream.ToArray();
					CopyStream(new EncryptedPackageHandler().EncryptPackage(package, Encryption), ref _stream);
				}
				else
				{
					_package.Save(_stream);
				}
				_stream.Flush();
				_package.Close();
				return;
			}
			if (System.IO.File.Exists(File.FullName))
			{
				try
				{
					System.IO.File.Delete(File.FullName);
				}
				catch (Exception innerException)
				{
					throw new Exception($"Error overwriting file {File.FullName}", innerException);
				}
			}
			_package.Save(_stream);
			_package.Close();
			if (Stream is MemoryStream)
			{
				FileStream fileStream = new FileStream(File.FullName, FileMode.Create);
				if (Encryption.IsEncrypted)
				{
					byte[] package2 = ((MemoryStream)Stream).ToArray();
					MemoryStream memoryStream2 = new EncryptedPackageHandler().EncryptPackage(package2, Encryption);
					fileStream.Write(memoryStream2.ToArray(), 0, (int)memoryStream2.Length);
				}
				else
				{
					fileStream.Write(((MemoryStream)Stream).ToArray(), 0, (int)Stream.Length);
				}
				fileStream.Close();
				fileStream.Dispose();
			}
			else
			{
				System.IO.File.WriteAllBytes(File.FullName, GetAsByteArray(save: false));
			}
		}
		catch (Exception innerException2)
		{
			if (File == null)
			{
				throw;
			}
			throw new InvalidOperationException($"Error saving file {File.FullName}", innerException2);
		}
	}

	public void Save(string password)
	{
		Encryption.Password = password;
		Save();
	}

	public void SaveAs(FileInfo file)
	{
		File = file;
		Save();
	}

	public void SaveAs(FileInfo file, string password)
	{
		File = file;
		Encryption.Password = password;
		Save();
	}

	public void SaveAs(Stream OutputStream)
	{
		File = null;
		Save();
		if (OutputStream != _stream)
		{
			CopyStream(_stream, ref OutputStream);
		}
	}

	public void SaveAs(Stream OutputStream, string password)
	{
		Encryption.Password = password;
		SaveAs(OutputStream);
	}

	internal void CloseStream()
	{
		if (_stream != null)
		{
			_stream.Close();
			_stream.Dispose();
		}
		_stream = new MemoryStream();
	}

	internal XmlDocument GetXmlFromUri(Uri uri)
	{
		XmlDocument xmlDocument = new XmlDocument();
		ZipPackagePart part = _package.GetPart(uri);
		XmlHelper.LoadXmlSafe(xmlDocument, part.GetStream());
		return xmlDocument;
	}

	public byte[] GetAsByteArray()
	{
		return GetAsByteArray(save: true);
	}

	public byte[] GetAsByteArray(string password)
	{
		if (password != null)
		{
			Encryption.Password = password;
		}
		return GetAsByteArray(save: true);
	}

	internal byte[] GetAsByteArray(bool save)
	{
		if (save)
		{
			Workbook.Save();
			_package.Close();
			_package.Save(_stream);
		}
		byte[] array = new byte[Stream.Length];
		long position = Stream.Position;
		Stream.Seek(0L, SeekOrigin.Begin);
		Stream.Read(array, 0, (int)Stream.Length);
		if (Encryption.IsEncrypted)
		{
			array = new EncryptedPackageHandler().EncryptPackage(array, Encryption).ToArray();
		}
		Stream.Seek(position, SeekOrigin.Begin);
		Stream.Close();
		return array;
	}

	public void Load(Stream input)
	{
		Load(input, new MemoryStream(), null);
	}

	public void Load(Stream input, string Password)
	{
		Load(input, new MemoryStream(), Password);
	}

	private void Load(Stream input, Stream output, string Password)
	{
		if (_package != null)
		{
			_package.Close();
			_package = null;
		}
		if (_stream != null)
		{
			_stream.Close();
			_stream.Dispose();
			_stream = null;
		}
		_isExternalStream = true;
		if (input.Length == 0L)
		{
			_stream = output;
			ConstructNewFile(Password);
		}
		else
		{
			_stream = output;
			Stream outputStream2;
			if (Password != null)
			{
				Stream outputStream = new MemoryStream();
				CopyStream(input, ref outputStream);
				EncryptedPackageHandler encryptedPackageHandler = new EncryptedPackageHandler();
				Encryption.Password = Password;
				outputStream2 = encryptedPackageHandler.DecryptPackage((MemoryStream)outputStream, Encryption);
			}
			else
			{
				outputStream2 = new MemoryStream();
				CopyStream(input, ref outputStream2);
			}
			try
			{
				_package = new ZipPackage(outputStream2);
			}
			catch (Exception innerException)
			{
				new EncryptedPackageHandler();
				if (Password == null && CompoundDocument.IsCompoundDocument((MemoryStream)_stream))
				{
					throw new Exception("Can not open the package. Package is an OLE compound document. If this is an encrypted package, please supply the password", innerException);
				}
				throw;
			}
		}
		_workbook = null;
	}

	internal static void CopyStream(Stream inputStream, ref Stream outputStream)
	{
		if (!inputStream.CanRead)
		{
			throw new Exception("Can not read from inputstream");
		}
		if (!outputStream.CanWrite)
		{
			throw new Exception("Can not write to outputstream");
		}
		if (inputStream.CanSeek)
		{
			inputStream.Seek(0L, SeekOrigin.Begin);
		}
		byte[] buffer = new byte[8096];
		lock (_lock)
		{
			for (int num = inputStream.Read(buffer, 0, 8096); num > 0; num = inputStream.Read(buffer, 0, 8096))
			{
				outputStream.Write(buffer, 0, num);
			}
			outputStream.Flush();
		}
	}
}
