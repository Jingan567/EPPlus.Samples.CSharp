using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace OfficeOpenXml.Utils;

[ComImport]
[ComVisible(false)]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("0000000A-0000-0000-C000-000000000046")]
internal interface ILockBytes
{
	void ReadAt(long ulOffset, IntPtr pv, int cb, out UIntPtr pcbRead);

	void WriteAt(long ulOffset, IntPtr pv, int cb, out UIntPtr pcbWritten);

	void Flush();

	void SetSize(long cb);

	void LockRegion(long libOffset, long cb, int dwLockType);

	void UnlockRegion(long libOffset, long cb, int dwLockType);

	void Stat(out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, int grfStatFlag);
}
