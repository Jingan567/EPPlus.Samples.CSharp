using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace OfficeOpenXml.Utils;

[ComImport]
[Guid("0000000d-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IEnumSTATSTG
{
	[PreserveSig]
	uint Next(uint celt, [Out][MarshalAs(UnmanagedType.LPArray)] System.Runtime.InteropServices.ComTypes.STATSTG[] rgelt, out uint pceltFetched);

	void Skip(uint celt);

	void Reset();

	[return: MarshalAs(UnmanagedType.Interface)]
	IEnumSTATSTG Clone();
}
