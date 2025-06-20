namespace OfficeOpenXml.Packaging.Ionic.Zip;

internal enum ExtractExistingFileAction
{
	Throw,
	OverwriteSilently,
	DoNotOverwrite,
	InvokeExtractProgressEvent
}
