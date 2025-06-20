using System.IO;

namespace OfficeOpenXml.Packaging.Ionic.Zip;

public delegate void WriteDelegate(string entryName, Stream stream);
