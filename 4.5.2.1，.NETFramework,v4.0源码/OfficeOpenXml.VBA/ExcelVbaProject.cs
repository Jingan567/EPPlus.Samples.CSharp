using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.CompundDocument;

namespace OfficeOpenXml.VBA;

public class ExcelVbaProject
{
	public enum eSyskind
	{
		Win16,
		Win32,
		Macintosh,
		Win64
	}

	private const string schemaRelVba = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

	internal const string PartUri = "/xl/vbaProject.bin";

	internal ExcelWorkbook _wb;

	internal ZipPackage _pck;

	private ExcelVbaSignature _signature;

	private ExcelVbaProtection _protection;

	public eSyskind SystemKind { get; set; }

	public string Name { get; set; }

	public string Description { get; set; }

	public string HelpFile1 { get; set; }

	public string HelpFile2 { get; set; }

	public int HelpContextID { get; set; }

	public string Constants { get; set; }

	public int CodePage { get; internal set; }

	internal int LibFlags { get; set; }

	internal int MajorVersion { get; set; }

	internal int MinorVersion { get; set; }

	internal int Lcid { get; set; }

	internal int LcidInvoke { get; set; }

	internal string ProjectID { get; set; }

	internal string ProjectStreamText { get; set; }

	public ExcelVbaReferenceCollection References { get; set; }

	public ExcelVbaModuleCollection Modules { get; set; }

	public ExcelVbaSignature Signature
	{
		get
		{
			if (_signature == null)
			{
				_signature = new ExcelVbaSignature(Part);
			}
			return _signature;
		}
	}

	public ExcelVbaProtection Protection
	{
		get
		{
			if (_protection == null)
			{
				_protection = new ExcelVbaProtection(this);
			}
			return _protection;
		}
	}

	internal CompoundDocument Document { get; set; }

	internal ZipPackagePart Part { get; set; }

	internal Uri Uri { get; private set; }

	internal ExcelVbaProject(ExcelWorkbook wb)
	{
		_wb = wb;
		_pck = _wb._package.Package;
		References = new ExcelVbaReferenceCollection();
		Modules = new ExcelVbaModuleCollection(this);
		ZipPackageRelationship zipPackageRelationship = _wb.Part.GetRelationshipsByType("http://schemas.microsoft.com/office/2006/relationships/vbaProject").FirstOrDefault();
		if (zipPackageRelationship != null)
		{
			Uri = UriHelper.ResolvePartUri(zipPackageRelationship.SourceUri, zipPackageRelationship.TargetUri);
			Part = _pck.GetPart(Uri);
			GetProject();
		}
		else
		{
			Lcid = 0;
			Part = null;
		}
	}

	private void GetProject()
	{
		MemoryStream stream = Part.GetStream();
		byte[] array = new byte[stream.Length];
		stream.Read(array, 0, (int)stream.Length);
		Document = new CompoundDocument(array);
		ReadDirStream();
		ProjectStreamText = Encoding.GetEncoding(CodePage).GetString(Document.Storage.DataStreams["PROJECT"]);
		ReadModules();
		ReadProjectProperties();
	}

	private void ReadModules()
	{
		foreach (ExcelVBAModule module in Modules)
		{
			byte[] bytes = VBACompression.DecompressPart(Document.Storage.SubStorage["VBA"].DataStreams[module.streamName], (int)module.ModuleOffset);
			string @string = Encoding.GetEncoding(CodePage).GetString(bytes);
			int num = 0;
			while (num + 9 < @string.Length && @string.Substring(num, 9) == "Attribute")
			{
				int num2 = @string.IndexOf("\r\n", num);
				string[] array = ((num2 <= 0) ? @string.Substring(num + 9).Split(new char[1] { '=' }, 1) : @string.Substring(num + 9, num2 - num - 9).Split('='));
				if (array.Length > 1)
				{
					array[1] = array[1].Trim();
					ExcelVbaModuleAttribute item = new ExcelVbaModuleAttribute
					{
						Name = array[0].Trim(),
						DataType = ((!array[1].StartsWith("\"")) ? eAttributeDataType.NonString : eAttributeDataType.String),
						Value = (array[1].StartsWith("\"") ? array[1].Substring(1, array[1].Length - 2) : array[1])
					};
					module.Attributes._list.Add(item);
				}
				num = num2 + 2;
			}
			module.Code = @string.Substring(num);
		}
	}

	private void ReadProjectProperties()
	{
		_protection = new ExcelVbaProtection(this);
		string classID = "";
		string[] array = Regex.Split(ProjectStreamText, "\r\n");
		foreach (string text in array)
		{
			if (text.StartsWith("["))
			{
				continue;
			}
			string[] array2 = text.Split('=');
			if (array2.Length > 1 && array2[1].Length > 1 && array2[1].StartsWith("\""))
			{
				array2[1] = array2[1].Substring(1, array2[1].Length - 2);
			}
			switch (array2[0])
			{
			case "ID":
				ProjectID = array2[1];
				break;
			case "Document":
			{
				string name = array2[1].Substring(0, array2[1].IndexOf("/&H"));
				Modules[name].Type = eModuleType.Document;
				break;
			}
			case "Package":
				classID = array2[1];
				break;
			case "BaseClass":
				Modules[array2[1]].Type = eModuleType.Designer;
				Modules[array2[1]].ClassID = classID;
				break;
			case "Module":
				Modules[array2[1]].Type = eModuleType.Module;
				break;
			case "Class":
				Modules[array2[1]].Type = eModuleType.Class;
				break;
			case "CMG":
			{
				byte[] array7 = Decrypt(array2[1]);
				_protection.UserProtected = (array7[0] & 1) != 0;
				_protection.HostProtected = (array7[0] & 2) != 0;
				_protection.VbeProtected = (array7[0] & 4) != 0;
				break;
			}
			case "DPB":
			{
				byte[] array3 = Decrypt(array2[1]);
				if (array3.Length < 28)
				{
					break;
				}
				_ = array3[0];
				byte[] array4 = new byte[3];
				Array.Copy(array3, 1, array4, 0, 3);
				byte[] array5 = new byte[4];
				_protection.PasswordKey = new byte[4];
				Array.Copy(array3, 4, array5, 0, 4);
				byte[] array6 = new byte[20];
				_protection.PasswordHash = new byte[20];
				Array.Copy(array3, 8, array6, 0, 20);
				for (int j = 0; j < 24; j++)
				{
					int num = 128 >> j % 8;
					if (j < 4)
					{
						if ((array4[0] & num) == 0)
						{
							_protection.PasswordKey[j] = 0;
						}
						else
						{
							_protection.PasswordKey[j] = array5[j];
						}
						continue;
					}
					int num2 = (j - j % 8) / 8;
					if ((array4[num2] & num) == 0)
					{
						_protection.PasswordHash[j - 4] = 0;
					}
					else
					{
						_protection.PasswordHash[j - 4] = array6[j - 4];
					}
				}
				break;
			}
			case "GC":
				_protection.VisibilityState = Decrypt(array2[1])[0] == byte.MaxValue;
				break;
			}
		}
	}

	private byte[] Decrypt(string value)
	{
		byte[] @byte = GetByte(value);
		byte[] array = new byte[value.Length - 1];
		byte b = @byte[0];
		array[0] = (byte)(@byte[1] ^ b);
		array[1] = (byte)(@byte[2] ^ b);
		for (int i = 2; i < @byte.Length - 1; i++)
		{
			array[i] = (byte)(@byte[i + 1] ^ (@byte[i - 1] + array[i - 1]));
		}
		_ = array[0];
		_ = array[1];
		byte b2 = (byte)((b & 6) / 2);
		int num = BitConverter.ToInt32(array, b2 + 2);
		byte[] array2 = new byte[num];
		Array.Copy(array, 6 + b2, array2, 0, num);
		return array2;
	}

	private string Encrypt(byte[] value)
	{
		byte[] array = new byte[1];
		RandomNumberGenerator.Create().GetBytes(array);
		BinaryWriter binaryWriter = new BinaryWriter(new MemoryStream());
		byte[] array2 = new byte[value.Length + 10];
		array2[0] = array[0];
		array2[1] = (byte)(2u ^ array[0]);
		byte b = 0;
		string projectID = ProjectID;
		foreach (char c in projectID)
		{
			b += (byte)c;
		}
		array2[2] = (byte)(b ^ array[0]);
		int num = (array[0] & 6) / 2;
		for (int j = 0; j < num; j++)
		{
			binaryWriter.Write(array[0]);
		}
		binaryWriter.Write(value.Length);
		binaryWriter.Write(value);
		int num2 = 3;
		byte b2 = b;
		byte[] array3 = ((MemoryStream)binaryWriter.BaseStream).ToArray();
		foreach (byte b3 in array3)
		{
			array2[num2] = (byte)(b3 ^ (array2[num2 - 2] + b2));
			num2++;
			b2 = b3;
		}
		return GetString(array2, num2 - 1);
	}

	private string GetString(byte[] value, int max)
	{
		string text = "";
		for (int i = 0; i <= max; i++)
		{
			text = ((value[i] >= 16) ? (text + value[i].ToString("x")) : (text + "0" + value[i].ToString("x")));
		}
		return text.ToUpperInvariant();
	}

	private byte[] GetByte(string value)
	{
		byte[] array = new byte[value.Length / 2];
		for (int i = 0; i < array.Length; i++)
		{
			array[i] = byte.Parse(value.Substring(i * 2, 2), NumberStyles.AllowHexSpecifier);
		}
		return array;
	}

	private void ReadDirStream()
	{
		BinaryReader binaryReader = new BinaryReader(new MemoryStream(VBACompression.DecompressPart(Document.Storage.SubStorage["VBA"].DataStreams["dir"])));
		ExcelVbaReference excelVbaReference = null;
		string name = "";
		ExcelVBAModule excelVBAModule = null;
		bool flag = false;
		while (binaryReader.BaseStream.Position < binaryReader.BaseStream.Length && !flag)
		{
			ushort num = binaryReader.ReadUInt16();
			uint size = binaryReader.ReadUInt32();
			switch (num)
			{
			case 1:
				SystemKind = (eSyskind)binaryReader.ReadUInt32();
				break;
			case 2:
				Lcid = (int)binaryReader.ReadUInt32();
				break;
			case 3:
				CodePage = binaryReader.ReadUInt16();
				break;
			case 4:
				Name = GetString(binaryReader, size);
				break;
			case 5:
				Description = GetUnicodeString(binaryReader, size);
				break;
			case 6:
				HelpFile1 = GetString(binaryReader, size);
				break;
			case 61:
				HelpFile2 = GetString(binaryReader, size);
				break;
			case 7:
				HelpContextID = (int)binaryReader.ReadUInt32();
				break;
			case 8:
				LibFlags = (int)binaryReader.ReadUInt32();
				break;
			case 9:
				MajorVersion = (int)binaryReader.ReadUInt32();
				MinorVersion = binaryReader.ReadUInt16();
				break;
			case 12:
				Constants = GetUnicodeString(binaryReader, size);
				break;
			case 13:
			{
				uint size4 = binaryReader.ReadUInt32();
				ExcelVbaReference excelVbaReference2 = new ExcelVbaReference();
				excelVbaReference2.Name = name;
				excelVbaReference2.ReferenceRecordID = num;
				excelVbaReference2.Libid = GetString(binaryReader, size4);
				binaryReader.ReadUInt32();
				binaryReader.ReadUInt16();
				References.Add(excelVbaReference2);
				break;
			}
			case 14:
			{
				ExcelVbaReferenceProject excelVbaReferenceProject = new ExcelVbaReferenceProject();
				excelVbaReferenceProject.ReferenceRecordID = num;
				excelVbaReferenceProject.Name = name;
				uint size4 = binaryReader.ReadUInt32();
				excelVbaReferenceProject.Libid = GetString(binaryReader, size4);
				size4 = binaryReader.ReadUInt32();
				excelVbaReferenceProject.LibIdRelative = GetString(binaryReader, size4);
				excelVbaReferenceProject.MajorVersion = binaryReader.ReadUInt32();
				excelVbaReferenceProject.MinorVersion = binaryReader.ReadUInt16();
				References.Add(excelVbaReferenceProject);
				break;
			}
			case 15:
				binaryReader.ReadUInt16();
				break;
			case 19:
				binaryReader.ReadUInt16();
				break;
			case 20:
				LcidInvoke = (int)binaryReader.ReadUInt32();
				break;
			case 22:
				name = GetUnicodeString(binaryReader, size);
				break;
			case 25:
				excelVBAModule = new ExcelVBAModule();
				excelVBAModule.Name = GetUnicodeString(binaryReader, size);
				Modules.Add(excelVBAModule);
				break;
			case 26:
				excelVBAModule.streamName = GetUnicodeString(binaryReader, size);
				break;
			case 28:
				excelVBAModule.Description = GetUnicodeString(binaryReader, size);
				break;
			case 30:
				excelVBAModule.HelpContext = (int)binaryReader.ReadUInt32();
				break;
			case 44:
				excelVBAModule.Cookie = binaryReader.ReadUInt16();
				break;
			case 49:
				excelVBAModule.ModuleOffset = binaryReader.ReadUInt32();
				break;
			case 16:
				flag = true;
				break;
			case 48:
			{
				ExcelVbaReferenceControl obj2 = (ExcelVbaReferenceControl)excelVbaReference;
				uint size3 = binaryReader.ReadUInt32();
				obj2.LibIdExternal = GetString(binaryReader, size3);
				binaryReader.ReadUInt32();
				binaryReader.ReadUInt16();
				obj2.OriginalTypeLib = new Guid(binaryReader.ReadBytes(16));
				obj2.Cookie = binaryReader.ReadUInt32();
				break;
			}
			case 51:
				excelVbaReference = new ExcelVbaReferenceControl();
				excelVbaReference.ReferenceRecordID = num;
				excelVbaReference.Name = name;
				excelVbaReference.Libid = GetString(binaryReader, size);
				References.Add(excelVbaReference);
				break;
			case 47:
			{
				ExcelVbaReferenceControl obj = (ExcelVbaReferenceControl)excelVbaReference;
				obj.ReferenceRecordID = num;
				uint size2 = binaryReader.ReadUInt32();
				obj.LibIdTwiddled = GetString(binaryReader, size2);
				binaryReader.ReadUInt32();
				binaryReader.ReadUInt16();
				break;
			}
			case 37:
				excelVBAModule.ReadOnly = true;
				break;
			case 40:
				excelVBAModule.Private = true;
				break;
			}
		}
	}

	internal void Save()
	{
		if (!Validate())
		{
			return;
		}
		CompoundDocument compoundDocument = new CompoundDocument();
		compoundDocument.Storage = new CompoundDocument.StoragePart();
		CompoundDocument.StoragePart storagePart = new CompoundDocument.StoragePart();
		compoundDocument.Storage.SubStorage.Add("VBA", storagePart);
		storagePart.DataStreams.Add("_VBA_PROJECT", CreateVBAProjectStream());
		storagePart.DataStreams.Add("dir", CreateDirStream());
		foreach (ExcelVBAModule module in Modules)
		{
			storagePart.DataStreams.Add(module.Name, VBACompression.CompressPart(Encoding.GetEncoding(CodePage).GetBytes(module.Attributes.GetAttributeText() + module.Code)));
		}
		if (Document != null)
		{
			foreach (KeyValuePair<string, CompoundDocument.StoragePart> item in Document.Storage.SubStorage)
			{
				if (item.Key != "VBA")
				{
					compoundDocument.Storage.SubStorage.Add(item.Key, item.Value);
				}
			}
			foreach (KeyValuePair<string, byte[]> dataStream in Document.Storage.DataStreams)
			{
				if (dataStream.Key != "dir" && dataStream.Key != "PROJECT" && dataStream.Key != "PROJECTwm")
				{
					compoundDocument.Storage.DataStreams.Add(dataStream.Key, dataStream.Value);
				}
			}
		}
		compoundDocument.Storage.DataStreams.Add("PROJECT", CreateProjectStream());
		compoundDocument.Storage.DataStreams.Add("PROJECTwm", CreateProjectwmStream());
		if (Part == null)
		{
			Uri = new Uri("/xl/vbaProject.bin", UriKind.Relative);
			Part = _pck.CreatePart(Uri, "application/vnd.ms-office.vbaProject");
			_wb.Part.CreateRelationship(Uri, TargetMode.Internal, "http://schemas.microsoft.com/office/2006/relationships/vbaProject");
		}
		MemoryStream stream = Part.GetStream(FileMode.Create);
		compoundDocument.Save(stream);
		stream.Flush();
		Signature.Save(this);
	}

	private bool Validate()
	{
		Description = Description ?? "";
		HelpFile1 = HelpFile1 ?? "";
		HelpFile2 = HelpFile2 ?? "";
		Constants = Constants ?? "";
		return true;
	}

	private byte[] CreateVBAProjectStream()
	{
		BinaryWriter binaryWriter = new BinaryWriter(new MemoryStream());
		binaryWriter.Write((ushort)25036);
		binaryWriter.Write(ushort.MaxValue);
		binaryWriter.Write((byte)0);
		binaryWriter.Write((ushort)0);
		return ((MemoryStream)binaryWriter.BaseStream).ToArray();
	}

	private byte[] CreateDirStream()
	{
		BinaryWriter binaryWriter = new BinaryWriter(new MemoryStream());
		binaryWriter.Write((ushort)1);
		binaryWriter.Write(4u);
		binaryWriter.Write((uint)SystemKind);
		binaryWriter.Write((ushort)2);
		binaryWriter.Write(4u);
		binaryWriter.Write((uint)Lcid);
		binaryWriter.Write((ushort)20);
		binaryWriter.Write(4u);
		binaryWriter.Write((uint)LcidInvoke);
		binaryWriter.Write((ushort)3);
		binaryWriter.Write(2u);
		binaryWriter.Write((ushort)CodePage);
		binaryWriter.Write((ushort)4);
		binaryWriter.Write((uint)Name.Length);
		binaryWriter.Write(Encoding.GetEncoding(CodePage).GetBytes(Name));
		binaryWriter.Write((ushort)5);
		binaryWriter.Write((uint)Description.Length);
		binaryWriter.Write(Encoding.GetEncoding(CodePage).GetBytes(Description));
		binaryWriter.Write((ushort)64);
		binaryWriter.Write((uint)(Description.Length * 2));
		binaryWriter.Write(Encoding.Unicode.GetBytes(Description));
		binaryWriter.Write((ushort)6);
		binaryWriter.Write((uint)HelpFile1.Length);
		binaryWriter.Write(Encoding.GetEncoding(CodePage).GetBytes(HelpFile1));
		binaryWriter.Write((ushort)61);
		binaryWriter.Write((uint)HelpFile2.Length);
		binaryWriter.Write(Encoding.GetEncoding(CodePage).GetBytes(HelpFile2));
		binaryWriter.Write((ushort)7);
		binaryWriter.Write(4u);
		binaryWriter.Write((uint)HelpContextID);
		binaryWriter.Write((ushort)8);
		binaryWriter.Write(4u);
		binaryWriter.Write(0u);
		binaryWriter.Write((ushort)9);
		binaryWriter.Write(4u);
		binaryWriter.Write((uint)MajorVersion);
		binaryWriter.Write((ushort)MinorVersion);
		binaryWriter.Write((ushort)12);
		binaryWriter.Write((uint)Constants.Length);
		binaryWriter.Write(Encoding.GetEncoding(CodePage).GetBytes(Constants));
		binaryWriter.Write((ushort)60);
		binaryWriter.Write((uint)Constants.Length / 2u);
		binaryWriter.Write(Encoding.Unicode.GetBytes(Constants));
		foreach (ExcelVbaReference reference in References)
		{
			WriteNameReference(binaryWriter, reference);
			if (reference.ReferenceRecordID == 47)
			{
				WriteControlReference(binaryWriter, reference);
			}
			else if (reference.ReferenceRecordID == 51)
			{
				WriteOrginalReference(binaryWriter, reference);
			}
			else if (reference.ReferenceRecordID == 13)
			{
				WriteRegisteredReference(binaryWriter, reference);
			}
			else if (reference.ReferenceRecordID == 14)
			{
				WriteProjectReference(binaryWriter, reference);
			}
		}
		binaryWriter.Write((ushort)15);
		binaryWriter.Write(2u);
		binaryWriter.Write((ushort)Modules.Count);
		binaryWriter.Write((ushort)19);
		binaryWriter.Write(2u);
		binaryWriter.Write(ushort.MaxValue);
		foreach (ExcelVBAModule module in Modules)
		{
			WriteModuleRecord(binaryWriter, module);
		}
		binaryWriter.Write((ushort)16);
		binaryWriter.Write(0u);
		return VBACompression.CompressPart(((MemoryStream)binaryWriter.BaseStream).ToArray());
	}

	private void WriteModuleRecord(BinaryWriter bw, ExcelVBAModule module)
	{
		bw.Write((ushort)25);
		bw.Write((uint)module.Name.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Name));
		bw.Write((ushort)71);
		bw.Write((uint)(module.Name.Length * 2));
		bw.Write(Encoding.Unicode.GetBytes(module.Name));
		bw.Write((ushort)26);
		bw.Write((uint)module.Name.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Name));
		bw.Write((ushort)50);
		bw.Write((uint)(module.Name.Length * 2));
		bw.Write(Encoding.Unicode.GetBytes(module.Name));
		module.Description = module.Description ?? "";
		bw.Write((ushort)28);
		bw.Write((uint)module.Description.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Description));
		bw.Write((ushort)72);
		bw.Write((uint)(module.Description.Length * 2));
		bw.Write(Encoding.Unicode.GetBytes(module.Description));
		bw.Write((ushort)49);
		bw.Write(4u);
		bw.Write(0u);
		bw.Write((ushort)30);
		bw.Write(4u);
		bw.Write((uint)module.HelpContext);
		bw.Write((ushort)44);
		bw.Write(2u);
		bw.Write(ushort.MaxValue);
		bw.Write((ushort)((module.Type == eModuleType.Module) ? 33u : 34u));
		bw.Write(0u);
		if (module.ReadOnly)
		{
			bw.Write((ushort)37);
			bw.Write(0u);
		}
		if (module.Private)
		{
			bw.Write((ushort)40);
			bw.Write(0u);
		}
		bw.Write((ushort)43);
		bw.Write(0u);
	}

	private void WriteNameReference(BinaryWriter bw, ExcelVbaReference reference)
	{
		bw.Write((ushort)22);
		bw.Write((uint)reference.Name.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(reference.Name));
		bw.Write((ushort)62);
		bw.Write((uint)(reference.Name.Length * 2));
		bw.Write(Encoding.Unicode.GetBytes(reference.Name));
	}

	private void WriteControlReference(BinaryWriter bw, ExcelVbaReference reference)
	{
		WriteOrginalReference(bw, reference);
		bw.Write((ushort)47);
		ExcelVbaReferenceControl excelVbaReferenceControl = (ExcelVbaReferenceControl)reference;
		bw.Write((uint)(4 + excelVbaReferenceControl.LibIdTwiddled.Length + 4 + 2));
		bw.Write((uint)excelVbaReferenceControl.LibIdTwiddled.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(excelVbaReferenceControl.LibIdTwiddled));
		bw.Write(0u);
		bw.Write((ushort)0);
		WriteNameReference(bw, reference);
		bw.Write((ushort)48);
		bw.Write((uint)(4 + excelVbaReferenceControl.LibIdExternal.Length + 4 + 2 + 16 + 4));
		bw.Write((uint)excelVbaReferenceControl.LibIdExternal.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(excelVbaReferenceControl.LibIdExternal));
		bw.Write(0u);
		bw.Write((ushort)0);
		bw.Write(excelVbaReferenceControl.OriginalTypeLib.ToByteArray());
		bw.Write(excelVbaReferenceControl.Cookie);
	}

	private void WriteOrginalReference(BinaryWriter bw, ExcelVbaReference reference)
	{
		bw.Write((ushort)51);
		bw.Write((uint)reference.Libid.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(reference.Libid));
	}

	private void WriteProjectReference(BinaryWriter bw, ExcelVbaReference reference)
	{
		bw.Write((ushort)14);
		ExcelVbaReferenceProject excelVbaReferenceProject = (ExcelVbaReferenceProject)reference;
		bw.Write((uint)(4 + excelVbaReferenceProject.Libid.Length + 4 + excelVbaReferenceProject.LibIdRelative.Length + 4 + 2));
		bw.Write((uint)excelVbaReferenceProject.Libid.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(excelVbaReferenceProject.Libid));
		bw.Write((uint)excelVbaReferenceProject.LibIdRelative.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(excelVbaReferenceProject.LibIdRelative));
		bw.Write(excelVbaReferenceProject.MajorVersion);
		bw.Write(excelVbaReferenceProject.MinorVersion);
	}

	private void WriteRegisteredReference(BinaryWriter bw, ExcelVbaReference reference)
	{
		bw.Write((ushort)13);
		bw.Write((uint)(4 + reference.Libid.Length + 4 + 2));
		bw.Write((uint)reference.Libid.Length);
		bw.Write(Encoding.GetEncoding(CodePage).GetBytes(reference.Libid));
		bw.Write(0u);
		bw.Write((ushort)0);
	}

	private byte[] CreateProjectwmStream()
	{
		BinaryWriter binaryWriter = new BinaryWriter(new MemoryStream());
		foreach (ExcelVBAModule module in Modules)
		{
			binaryWriter.Write(Encoding.GetEncoding(CodePage).GetBytes(module.Name));
			binaryWriter.Write((byte)0);
			binaryWriter.Write(Encoding.Unicode.GetBytes(module.Name));
			binaryWriter.Write((ushort)0);
		}
		binaryWriter.Write((ushort)0);
		return ((MemoryStream)binaryWriter.BaseStream).ToArray();
	}

	private byte[] CreateProjectStream()
	{
		StringBuilder stringBuilder = new StringBuilder();
		stringBuilder.AppendFormat("ID=\"{0}\"\r\n", ProjectID);
		foreach (ExcelVBAModule module in Modules)
		{
			if (module.Type == eModuleType.Document)
			{
				stringBuilder.AppendFormat("Document={0}/&H00000000\r\n", module.Name);
				continue;
			}
			if (module.Type == eModuleType.Module)
			{
				stringBuilder.AppendFormat("Module={0}\r\n", module.Name);
				continue;
			}
			if (module.Type == eModuleType.Class)
			{
				stringBuilder.AppendFormat("Class={0}\r\n", module.Name);
				continue;
			}
			stringBuilder.AppendFormat("Package={0}\r\n", module.ClassID);
			stringBuilder.AppendFormat("BaseClass={0}\r\n", module.Name);
		}
		if (HelpFile1 != "")
		{
			stringBuilder.AppendFormat("HelpFile={0}\r\n", HelpFile1);
		}
		stringBuilder.AppendFormat("Name=\"{0}\"\r\n", Name);
		stringBuilder.AppendFormat("HelpContextID={0}\r\n", HelpContextID);
		if (!string.IsNullOrEmpty(Description))
		{
			stringBuilder.AppendFormat("Description=\"{0}\"\r\n", Description);
		}
		stringBuilder.AppendFormat("VersionCompatible32=\"393222000\"\r\n");
		stringBuilder.AppendFormat("CMG=\"{0}\"\r\n", WriteProtectionStat());
		stringBuilder.AppendFormat("DPB=\"{0}\"\r\n", WritePassword());
		stringBuilder.AppendFormat("GC=\"{0}\"\r\n\r\n", WriteVisibilityState());
		stringBuilder.Append("[Host Extender Info]\r\n");
		stringBuilder.Append("&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n");
		stringBuilder.Append("\r\n");
		stringBuilder.Append("[Workspace]\r\n");
		foreach (ExcelVBAModule module2 in Modules)
		{
			stringBuilder.AppendFormat("{0}=0, 0, 0, 0, C \r\n", module2.Name);
		}
		string s = stringBuilder.ToString();
		return Encoding.GetEncoding(CodePage).GetBytes(s);
	}

	private string WriteProtectionStat()
	{
		int value = (_protection.UserProtected ? 1 : 0) | (_protection.HostProtected ? 2 : 0) | (_protection.VbeProtected ? 4 : 0);
		return Encrypt(BitConverter.GetBytes(value));
	}

	private string WritePassword()
	{
		byte[] array = new byte[3];
		byte[] array2 = new byte[4];
		byte[] array3 = new byte[20];
		if (Protection.PasswordKey == null)
		{
			return Encrypt(new byte[1]);
		}
		Array.Copy(Protection.PasswordKey, array2, 4);
		Array.Copy(Protection.PasswordHash, array3, 20);
		for (int i = 0; i < 24; i++)
		{
			byte b = (byte)(128 >> i % 8);
			if (i < 4)
			{
				if (array2[i] == 0)
				{
					array2[i] = 1;
				}
				else
				{
					array[0] |= b;
				}
			}
			else if (array3[i - 4] == 0)
			{
				array3[i - 4] = 1;
			}
			else
			{
				int num = (i - i % 8) / 8;
				array[num] |= b;
			}
		}
		BinaryWriter binaryWriter = new BinaryWriter(new MemoryStream());
		binaryWriter.Write(byte.MaxValue);
		binaryWriter.Write(array);
		binaryWriter.Write(array2);
		binaryWriter.Write(array3);
		binaryWriter.Write((byte)0);
		return Encrypt(((MemoryStream)binaryWriter.BaseStream).ToArray());
	}

	private string WriteVisibilityState()
	{
		return Encrypt(new byte[1] { (byte)(Protection.VisibilityState ? 255u : 0u) });
	}

	private string GetString(BinaryReader br, uint size)
	{
		return GetString(br, size, Encoding.GetEncoding(CodePage));
	}

	private string GetString(BinaryReader br, uint size, Encoding enc)
	{
		if (size != 0)
		{
			byte[] array = new byte[size];
			array = br.ReadBytes((int)size);
			return enc.GetString(array);
		}
		return "";
	}

	private string GetUnicodeString(BinaryReader br, uint size)
	{
		string @string = GetString(br, size);
		br.ReadUInt16();
		uint size2 = br.ReadUInt32();
		string string2 = GetString(br, size2, Encoding.Unicode);
		if (string2.Length != 0)
		{
			return string2;
		}
		return @string;
	}

	internal void Create()
	{
		if (Lcid > 0)
		{
			throw new InvalidOperationException("Package already contains a VBAProject");
		}
		ProjectID = "{5DD90D76-4904-47A2-AF0D-D69B4673604E}";
		Name = "VBAProject";
		SystemKind = eSyskind.Win32;
		Lcid = 1033;
		LcidInvoke = 1033;
		CodePage = Encoding.GetEncoding(0).CodePage;
		MajorVersion = 1361024421;
		MinorVersion = 6;
		HelpContextID = 0;
		Modules.Add(new ExcelVBAModule(_wb.CodeNameChange)
		{
			Name = "ThisWorkbook",
			Code = "",
			Attributes = GetDocumentAttributes("ThisWorkbook", "0{00020819-0000-0000-C000-000000000046}"),
			Type = eModuleType.Document,
			HelpContext = 0
		});
		foreach (ExcelWorksheet worksheet in _wb.Worksheets)
		{
			string moduleNameFromWorksheet = GetModuleNameFromWorksheet(worksheet);
			if (!Modules.Exists(moduleNameFromWorksheet))
			{
				Modules.Add(new ExcelVBAModule(worksheet.CodeNameChange)
				{
					Name = moduleNameFromWorksheet,
					Code = "",
					Attributes = GetDocumentAttributes(worksheet.Name, "0{00020820-0000-0000-C000-000000000046}"),
					Type = eModuleType.Document,
					HelpContext = 0
				});
			}
		}
		_protection = new ExcelVbaProtection(this)
		{
			UserProtected = false,
			HostProtected = false,
			VbeProtected = false,
			VisibilityState = true
		};
	}

	internal string GetModuleNameFromWorksheet(ExcelWorksheet sheet)
	{
		string name = sheet.Name;
		name = name.Substring(0, (name.Length < 31) ? name.Length : 31);
		if (Modules[name] != null || !Regex.IsMatch(name, "^[a-zA-Z][a-zA-Z0-9_ ]*$"))
		{
			int num = sheet.PositionID;
			name = "Sheet" + num;
			while (Modules[name] != null)
			{
				int num2 = num + 1;
				num = num2;
				name = "Sheet" + num2;
			}
		}
		return name;
	}

	internal ExcelVbaModuleAttributesCollection GetDocumentAttributes(string name, string clsid)
	{
		return new ExcelVbaModuleAttributesCollection
		{
			_list = 
			{
				new ExcelVbaModuleAttribute
				{
					Name = "VB_Name",
					Value = name,
					DataType = eAttributeDataType.String
				},
				new ExcelVbaModuleAttribute
				{
					Name = "VB_Base",
					Value = clsid,
					DataType = eAttributeDataType.String
				},
				new ExcelVbaModuleAttribute
				{
					Name = "VB_GlobalNameSpace",
					Value = "False",
					DataType = eAttributeDataType.NonString
				},
				new ExcelVbaModuleAttribute
				{
					Name = "VB_Creatable",
					Value = "False",
					DataType = eAttributeDataType.NonString
				},
				new ExcelVbaModuleAttribute
				{
					Name = "VB_PredeclaredId",
					Value = "True",
					DataType = eAttributeDataType.NonString
				},
				new ExcelVbaModuleAttribute
				{
					Name = "VB_Exposed",
					Value = "False",
					DataType = eAttributeDataType.NonString
				},
				new ExcelVbaModuleAttribute
				{
					Name = "VB_TemplateDerived",
					Value = "False",
					DataType = eAttributeDataType.NonString
				},
				new ExcelVbaModuleAttribute
				{
					Name = "VB_Customizable",
					Value = "True",
					DataType = eAttributeDataType.NonString
				}
			}
		};
	}

	public void Remove()
	{
		if (Part == null)
		{
			return;
		}
		foreach (ZipPackageRelationship relationship in Part.GetRelationships())
		{
			_pck.DeleteRelationship(relationship.Id);
		}
		if (_pck.PartExists(Uri))
		{
			_pck.DeletePart(Uri);
		}
		Part = null;
		Modules.Clear();
		References.Clear();
		Lcid = 0;
		LcidInvoke = 0;
		CodePage = 0;
		MajorVersion = 0;
		MinorVersion = 0;
		HelpContextID = 0;
	}

	public override string ToString()
	{
		return Name;
	}
}
