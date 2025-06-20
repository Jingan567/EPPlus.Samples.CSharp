using System;
using System.Collections;
using System.Collections.Generic;
using System.Security;
using System.Text;
using OfficeOpenXml.Packaging.Ionic.Zip;

namespace OfficeOpenXml.Packaging;

public class ZipPackageRelationshipCollection : IEnumerable<ZipPackageRelationship>, IEnumerable
{
	protected internal Dictionary<string, ZipPackageRelationship> _rels = new Dictionary<string, ZipPackageRelationship>(StringComparer.OrdinalIgnoreCase);

	internal ZipPackageRelationship this[string id] => _rels[id];

	public int Count => _rels.Count;

	internal void Add(ZipPackageRelationship item)
	{
		_rels.Add(item.Id, item);
	}

	public IEnumerator<ZipPackageRelationship> GetEnumerator()
	{
		return _rels.Values.GetEnumerator();
	}

	IEnumerator IEnumerable.GetEnumerator()
	{
		return _rels.Values.GetEnumerator();
	}

	internal void Remove(string id)
	{
		_rels.Remove(id);
	}

	internal bool ContainsKey(string id)
	{
		return _rels.ContainsKey(id);
	}

	internal ZipPackageRelationshipCollection GetRelationshipsByType(string relationshipType)
	{
		ZipPackageRelationshipCollection zipPackageRelationshipCollection = new ZipPackageRelationshipCollection();
		foreach (ZipPackageRelationship value in _rels.Values)
		{
			if (value.RelationshipType == relationshipType)
			{
				zipPackageRelationshipCollection.Add(value);
			}
		}
		return zipPackageRelationshipCollection;
	}

	internal void WriteZip(ZipOutputStream os, string fileName)
	{
		StringBuilder stringBuilder = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
		foreach (ZipPackageRelationship value in _rels.Values)
		{
			stringBuilder.AppendFormat("<Relationship Id=\"{0}\" Type=\"{1}\" Target=\"{2}\"{3}/>", SecurityElement.Escape(value.Id), value.RelationshipType, SecurityElement.Escape(value.TargetUri.OriginalString), (value.TargetMode == TargetMode.External) ? " TargetMode=\"External\"" : "");
		}
		stringBuilder.Append("</Relationships>");
		os.PutNextEntry(fileName);
		byte[] bytes = Encoding.UTF8.GetBytes(stringBuilder.ToString());
		os.Write(bytes, 0, bytes.Length);
	}
}
