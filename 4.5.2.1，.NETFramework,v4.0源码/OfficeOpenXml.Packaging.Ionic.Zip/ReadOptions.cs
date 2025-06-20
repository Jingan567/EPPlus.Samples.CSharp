using System;
using System.IO;
using System.Text;

namespace OfficeOpenXml.Packaging.Ionic.Zip;

internal class ReadOptions
{
	public EventHandler<ReadProgressEventArgs> ReadProgress { get; set; }

	public TextWriter StatusMessageWriter { get; set; }

	public Encoding Encoding { get; set; }
}
