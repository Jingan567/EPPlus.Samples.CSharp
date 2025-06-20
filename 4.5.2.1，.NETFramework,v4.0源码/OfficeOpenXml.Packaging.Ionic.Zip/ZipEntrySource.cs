namespace OfficeOpenXml.Packaging.Ionic.Zip;

internal enum ZipEntrySource
{
	None,
	FileSystem,
	Stream,
	ZipFile,
	WriteDelegate,
	JitStream,
	ZipOutputStream
}
