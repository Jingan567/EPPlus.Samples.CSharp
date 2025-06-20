using System;

namespace OfficeOpenXml.Packaging.Ionic.Zip;

[Flags]
internal enum ZipEntryTimestamp
{
	None = 0,
	DOS = 1,
	Windows = 2,
	Unix = 4,
	InfoZip1 = 8
}
