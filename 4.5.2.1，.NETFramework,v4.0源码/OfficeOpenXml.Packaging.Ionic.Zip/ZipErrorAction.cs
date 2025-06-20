namespace OfficeOpenXml.Packaging.Ionic.Zip;

internal enum ZipErrorAction
{
	Throw,
	Skip,
	Retry,
	InvokeErrorEvent
}
