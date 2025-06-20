using System.IO;

namespace OfficeOpenXml.Packaging.Ionic.Zip;

public delegate void CloseDelegate(string entryName, Stream stream);
