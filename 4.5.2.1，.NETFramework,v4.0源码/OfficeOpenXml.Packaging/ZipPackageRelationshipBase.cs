using System;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Packaging;

public abstract class ZipPackageRelationshipBase
{
	protected ZipPackageRelationshipCollection _rels = new ZipPackageRelationshipCollection();

	protected internal int maxRId = 1;

	internal void DeleteRelationship(string id)
	{
		_rels.Remove(id);
		UpdateMaxRId(id, ref maxRId);
	}

	protected void UpdateMaxRId(string id, ref int maxRId)
	{
		if (id.StartsWith("rId") && int.TryParse(id.Substring(3), out var result) && result == maxRId - 1)
		{
			maxRId--;
		}
	}

	internal virtual ZipPackageRelationship CreateRelationship(Uri targetUri, TargetMode targetMode, string relationshipType)
	{
		ZipPackageRelationship zipPackageRelationship = new ZipPackageRelationship();
		zipPackageRelationship.TargetUri = targetUri;
		zipPackageRelationship.TargetMode = targetMode;
		zipPackageRelationship.RelationshipType = relationshipType;
		zipPackageRelationship.Id = "rId" + maxRId++;
		_rels.Add(zipPackageRelationship);
		return zipPackageRelationship;
	}

	internal bool RelationshipExists(string id)
	{
		return _rels.ContainsKey(id);
	}

	internal ZipPackageRelationshipCollection GetRelationshipsByType(string schema)
	{
		return _rels.GetRelationshipsByType(schema);
	}

	internal ZipPackageRelationshipCollection GetRelationships()
	{
		return _rels;
	}

	internal ZipPackageRelationship GetRelationship(string id)
	{
		return _rels[id];
	}

	internal void ReadRelation(string xml, string source)
	{
		XmlDocument xmlDocument = new XmlDocument();
		XmlHelper.LoadXmlSafe(xmlDocument, xml, Encoding.UTF8);
		foreach (XmlElement childNode in xmlDocument.DocumentElement.ChildNodes)
		{
			ZipPackageRelationship zipPackageRelationship = new ZipPackageRelationship();
			zipPackageRelationship.Id = childNode.GetAttribute("Id");
			zipPackageRelationship.RelationshipType = childNode.GetAttribute("Type");
			zipPackageRelationship.TargetMode = (childNode.GetAttribute("TargetMode").Equals("external", StringComparison.OrdinalIgnoreCase) ? TargetMode.External : TargetMode.Internal);
			try
			{
				zipPackageRelationship.TargetUri = new Uri(childNode.GetAttribute("Target"), UriKind.RelativeOrAbsolute);
			}
			catch
			{
				zipPackageRelationship.TargetUri = new Uri(Uri.EscapeUriString("Invalid:URI " + childNode.GetAttribute("Target")), UriKind.RelativeOrAbsolute);
			}
			if (!string.IsNullOrEmpty(source))
			{
				zipPackageRelationship.SourceUri = new Uri(source, UriKind.Relative);
			}
			if (zipPackageRelationship.Id.StartsWith("rid", StringComparison.OrdinalIgnoreCase) && int.TryParse(zipPackageRelationship.Id.Substring(3), out var result) && result >= maxRId && result < 2147473647)
			{
				maxRId = result + 1;
			}
			_rels.Add(zipPackageRelationship);
		}
	}
}
