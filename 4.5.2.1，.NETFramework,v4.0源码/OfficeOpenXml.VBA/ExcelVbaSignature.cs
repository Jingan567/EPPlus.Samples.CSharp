using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.CompundDocument;

namespace OfficeOpenXml.VBA;

public class ExcelVbaSignature
{
	private const string schemaRelVbaSignature = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";

	private ZipPackagePart _vbaPart;

	public X509Certificate2 Certificate { get; set; }

	public SignedCms Verifier { get; internal set; }

	internal CompoundDocument Signature { get; set; }

	internal ZipPackagePart Part { get; set; }

	internal Uri Uri { get; private set; }

	internal ExcelVbaSignature(ZipPackagePart vbaPart)
	{
		_vbaPart = vbaPart;
		GetSignature();
	}

	private void GetSignature()
	{
		if (_vbaPart == null)
		{
			return;
		}
		ZipPackageRelationship zipPackageRelationship = _vbaPart.GetRelationshipsByType("http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature").FirstOrDefault();
		if (zipPackageRelationship != null)
		{
			Uri = UriHelper.ResolvePartUri(zipPackageRelationship.SourceUri, zipPackageRelationship.TargetUri);
			Part = _vbaPart.Package.GetPart(Uri);
			BinaryReader binaryReader = new BinaryReader(Part.GetStream());
			uint count = binaryReader.ReadUInt32();
			binaryReader.ReadUInt32();
			binaryReader.ReadUInt32();
			binaryReader.ReadUInt32();
			binaryReader.ReadUInt32();
			binaryReader.ReadUInt32();
			binaryReader.ReadUInt32();
			binaryReader.ReadUInt32();
			binaryReader.ReadUInt32();
			byte[] encodedMessage = binaryReader.ReadBytes((int)count);
			binaryReader.ReadUInt32();
			binaryReader.ReadUInt32();
			for (uint num = binaryReader.ReadUInt32(); num != 0; num = binaryReader.ReadUInt32())
			{
				binaryReader.ReadUInt32();
				uint num2 = binaryReader.ReadUInt32();
				if (num2 != 0)
				{
					byte[] rawData = binaryReader.ReadBytes((int)num2);
					if (num == 32)
					{
						Certificate = new X509Certificate2(rawData);
					}
				}
			}
			binaryReader.ReadUInt32();
			binaryReader.ReadUInt32();
			binaryReader.ReadUInt16();
			binaryReader.ReadUInt16();
			Verifier = new SignedCms();
			Verifier.Decode(encodedMessage);
		}
		else
		{
			Certificate = null;
			Verifier = null;
		}
	}

	internal void Save(ExcelVbaProject proj)
	{
		if (Certificate == null)
		{
			return;
		}
		if (!Certificate.HasPrivateKey)
		{
			X509Certificate2 certFromStore = GetCertFromStore(StoreLocation.CurrentUser);
			if (certFromStore == null)
			{
				certFromStore = GetCertFromStore(StoreLocation.LocalMachine);
			}
			if (certFromStore == null || !certFromStore.HasPrivateKey)
			{
				foreach (ZipPackageRelationship relationship in Part.GetRelationships())
				{
					Part.DeleteRelationship(relationship.Id);
				}
				Part.Package.DeletePart(Part.Uri);
				return;
			}
			Certificate = certFromStore;
		}
		MemoryStream memoryStream = new MemoryStream();
		BinaryWriter binaryWriter = new BinaryWriter(memoryStream);
		byte[] certStore = GetCertStore();
		byte[] array = SignProject(proj);
		binaryWriter.Write((uint)array.Length);
		binaryWriter.Write(44u);
		binaryWriter.Write((uint)certStore.Length);
		binaryWriter.Write((uint)(array.Length + 44));
		binaryWriter.Write(0u);
		binaryWriter.Write((uint)(array.Length + certStore.Length + 44));
		binaryWriter.Write(0u);
		binaryWriter.Write(0u);
		binaryWriter.Write((uint)(array.Length + certStore.Length + 44 + 2));
		binaryWriter.Write(array);
		binaryWriter.Write(certStore);
		binaryWriter.Write((ushort)0);
		binaryWriter.Write((ushort)0);
		binaryWriter.Write((ushort)0);
		binaryWriter.Flush();
		ZipPackageRelationship zipPackageRelationship = proj.Part.GetRelationshipsByType("http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature").FirstOrDefault();
		if (Part == null)
		{
			if (zipPackageRelationship != null)
			{
				Uri = zipPackageRelationship.TargetUri;
				Part = proj._pck.GetPart(zipPackageRelationship.TargetUri);
			}
			else
			{
				Uri = new Uri("/xl/vbaProjectSignature.bin", UriKind.Relative);
				Part = proj._pck.CreatePart(Uri, "application/vnd.ms-office.vbaProjectSignature");
			}
		}
		if (zipPackageRelationship == null)
		{
			proj.Part.CreateRelationship(UriHelper.ResolvePartUri(proj.Uri, Uri), TargetMode.Internal, "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature");
		}
		byte[] array2 = memoryStream.ToArray();
		Part.GetStream(FileMode.Create).Write(array2, 0, array2.Length);
	}

	private X509Certificate2 GetCertFromStore(StoreLocation loc)
	{
		try
		{
			X509Store x509Store = new X509Store(StoreName.My, loc);
			x509Store.Open(OpenFlags.ReadOnly);
			try
			{
				return x509Store.Certificates.Find(X509FindType.FindByThumbprint, Certificate.Thumbprint, validOnly: true).OfType<X509Certificate2>().FirstOrDefault();
			}
			finally
			{
				x509Store.Close();
			}
		}
		catch
		{
			return null;
		}
	}

	private byte[] GetCertStore()
	{
		MemoryStream memoryStream = new MemoryStream();
		BinaryWriter binaryWriter = new BinaryWriter(memoryStream);
		binaryWriter.Write(0u);
		binaryWriter.Write(1414677827u);
		byte[] rawData = Certificate.RawData;
		binaryWriter.Write(32u);
		binaryWriter.Write(1u);
		binaryWriter.Write((uint)rawData.Length);
		binaryWriter.Write(rawData);
		binaryWriter.Write(0u);
		binaryWriter.Write(0uL);
		binaryWriter.Flush();
		return memoryStream.ToArray();
	}

	private void WriteProp(BinaryWriter bw, int id, byte[] data)
	{
		bw.Write((uint)id);
		bw.Write(1u);
		bw.Write((uint)data.Length);
		bw.Write(data);
	}

	internal byte[] SignProject(ExcelVbaProject proj)
	{
		if (!Certificate.HasPrivateKey)
		{
			Certificate = null;
			return null;
		}
		byte[] contentHash = GetContentHash(proj);
		BinaryWriter binaryWriter = new BinaryWriter(new MemoryStream());
		binaryWriter.Write((byte)48);
		binaryWriter.Write((byte)50);
		binaryWriter.Write((byte)48);
		binaryWriter.Write((byte)14);
		binaryWriter.Write((byte)6);
		binaryWriter.Write((byte)10);
		binaryWriter.Write(new byte[10] { 43, 6, 1, 4, 1, 130, 55, 2, 1, 29 });
		binaryWriter.Write((byte)4);
		binaryWriter.Write((byte)0);
		binaryWriter.Write((byte)48);
		binaryWriter.Write((byte)32);
		binaryWriter.Write((byte)48);
		binaryWriter.Write((byte)12);
		binaryWriter.Write((byte)6);
		binaryWriter.Write((byte)8);
		binaryWriter.Write(new byte[8] { 42, 134, 72, 134, 247, 13, 2, 5 });
		binaryWriter.Write((byte)5);
		binaryWriter.Write((byte)0);
		binaryWriter.Write((byte)4);
		binaryWriter.Write((byte)contentHash.Length);
		binaryWriter.Write(contentHash);
		ContentInfo contentInfo = new ContentInfo(((MemoryStream)binaryWriter.BaseStream).ToArray());
		contentInfo.ContentType.Value = "1.3.6.1.4.1.311.2.1.4";
		Verifier = new SignedCms(contentInfo);
		CmsSigner signer = new CmsSigner(Certificate);
		Verifier.ComputeSignature(signer, silent: false);
		return Verifier.Encode();
	}

	private byte[] GetContentHash(ExcelVbaProject proj)
	{
		Encoding encoding = Encoding.GetEncoding(proj.CodePage);
		BinaryWriter binaryWriter = new BinaryWriter(new MemoryStream());
		binaryWriter.Write(encoding.GetBytes(proj.Name));
		binaryWriter.Write(encoding.GetBytes(proj.Constants));
		foreach (ExcelVbaReference reference in proj.References)
		{
			if (reference.ReferenceRecordID == 13)
			{
				binaryWriter.Write((byte)123);
			}
			if (reference.ReferenceRecordID != 14)
			{
				continue;
			}
			byte[] bytes = BitConverter.GetBytes((uint)reference.Libid.Length);
			foreach (byte b in bytes)
			{
				if (b == 0)
				{
					break;
				}
				binaryWriter.Write(b);
			}
		}
		foreach (ExcelVBAModule module in proj.Modules)
		{
			string[] array = module.Code.Split(new char[2] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
			foreach (string text in array)
			{
				if (!text.StartsWith("attribute", StringComparison.OrdinalIgnoreCase))
				{
					binaryWriter.Write(encoding.GetBytes(text));
				}
			}
		}
		byte[] buffer = (binaryWriter.BaseStream as MemoryStream).ToArray();
		return MD5.Create().ComputeHash(buffer);
	}
}
