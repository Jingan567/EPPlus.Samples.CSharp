using System;

namespace OfficeOpenXml.Encryption;

[Flags]
internal enum Flags
{
	Reserved1 = 1,
	Reserved2 = 2,
	fCryptoAPI = 4,
	fDocProps = 8,
	fExternal = 0x10,
	fAES = 0x20
}
