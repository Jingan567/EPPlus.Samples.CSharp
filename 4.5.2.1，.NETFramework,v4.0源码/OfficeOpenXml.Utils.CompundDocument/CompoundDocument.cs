using System.Collections.Generic;
using System.IO;

namespace OfficeOpenXml.Utils.CompundDocument;

internal class CompoundDocument
{
	internal class StoragePart
	{
		internal Dictionary<string, StoragePart> SubStorage = new Dictionary<string, StoragePart>();

		internal Dictionary<string, byte[]> DataStreams = new Dictionary<string, byte[]>();
	}

	internal StoragePart Storage;

	internal CompoundDocument()
	{
		Storage = new StoragePart();
	}

	internal CompoundDocument(MemoryStream ms)
	{
		Read(ms);
	}

	internal CompoundDocument(FileInfo fi)
	{
		Read(fi);
	}

	internal static bool IsCompoundDocument(FileInfo fi)
	{
		return CompoundDocumentFile.IsCompoundDocument(fi);
	}

	internal static bool IsCompoundDocument(MemoryStream ms)
	{
		return CompoundDocumentFile.IsCompoundDocument(ms);
	}

	internal CompoundDocument(byte[] doc)
	{
		Read(doc);
	}

	internal void Read(FileInfo fi)
	{
		byte[] doc = File.ReadAllBytes(fi.FullName);
		Read(doc);
	}

	internal void Read(byte[] doc)
	{
		Read(new MemoryStream(doc));
	}

	internal void Read(MemoryStream ms)
	{
		using CompoundDocumentFile compoundDocumentFile = new CompoundDocumentFile(ms);
		Storage = new StoragePart();
		GetStorageAndStreams(Storage, compoundDocumentFile.RootItem);
	}

	private void GetStorageAndStreams(StoragePart storage, CompoundDocumentItem parent)
	{
		foreach (CompoundDocumentItem child in parent.Children)
		{
			if (child.ObjectType == 1)
			{
				StoragePart storagePart = new StoragePart();
				storage.SubStorage.Add(child.Name, storagePart);
				GetStorageAndStreams(storagePart, child);
			}
			else if (child.ObjectType == 2)
			{
				storage.DataStreams.Add(child.Name, child.Stream);
			}
		}
	}

	internal void Save(MemoryStream ms)
	{
		CompoundDocumentFile compoundDocumentFile = new CompoundDocumentFile();
		WriteStorageAndStreams(Storage, compoundDocumentFile.RootItem);
		compoundDocumentFile.Write(ms);
	}

	private void WriteStorageAndStreams(StoragePart storage, CompoundDocumentItem parent)
	{
		foreach (KeyValuePair<string, StoragePart> item2 in storage.SubStorage)
		{
			CompoundDocumentItem compoundDocumentItem = new CompoundDocumentItem
			{
				Name = item2.Key,
				ObjectType = 1,
				Stream = null,
				StreamSize = 0L,
				Parent = parent
			};
			parent.Children.Add(compoundDocumentItem);
			WriteStorageAndStreams(item2.Value, compoundDocumentItem);
		}
		foreach (KeyValuePair<string, byte[]> dataStream in storage.DataStreams)
		{
			CompoundDocumentItem item = new CompoundDocumentItem
			{
				Name = dataStream.Key,
				ObjectType = 2,
				Stream = dataStream.Value,
				StreamSize = ((dataStream.Value != null) ? dataStream.Value.Length : 0),
				Parent = parent
			};
			parent.Children.Add(item);
		}
	}
}
