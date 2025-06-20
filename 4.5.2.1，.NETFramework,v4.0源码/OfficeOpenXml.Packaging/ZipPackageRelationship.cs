using System;

namespace OfficeOpenXml.Packaging;

public class ZipPackageRelationship
{
	public Uri TargetUri { get; internal set; }

	public Uri SourceUri { get; internal set; }

	public string RelationshipType { get; internal set; }

	public TargetMode TargetMode { get; internal set; }

	public string Id { get; internal set; }
}
