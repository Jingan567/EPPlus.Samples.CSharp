using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeOpenXml.Utils.CompundDocument;

internal class CompoundDocumentItem : IComparable<CompoundDocumentItem>
{
	internal bool _handled;

	public CompoundDocumentItem Parent { get; set; }

	public List<CompoundDocumentItem> Children { get; set; }

	public string Name { get; set; }

	public byte ColorFlag { get; set; }

	public byte ObjectType { get; set; }

	public int ChildID { get; set; }

	public Guid ClsID { get; set; }

	public int LeftSibling { get; set; }

	public int RightSibling { get; set; }

	public int StatBits { get; set; }

	public long CreationTime { get; set; }

	public long ModifiedTime { get; set; }

	public int StartingSectorLocation { get; set; }

	public long StreamSize { get; set; }

	public byte[] Stream { get; set; }

	public CompoundDocumentItem()
	{
		Children = new List<CompoundDocumentItem>();
	}

	internal void Read(BinaryReader br)
	{
		byte[] bytes = br.ReadBytes(64);
		short num = br.ReadInt16();
		if (num > 0)
		{
			Name = Encoding.Unicode.GetString(bytes, 0, num - 2);
		}
		ObjectType = br.ReadByte();
		ColorFlag = br.ReadByte();
		LeftSibling = br.ReadInt32();
		RightSibling = br.ReadInt32();
		ChildID = br.ReadInt32();
		ClsID = new Guid(br.ReadBytes(16));
		StatBits = br.ReadInt32();
		CreationTime = br.ReadInt64();
		ModifiedTime = br.ReadInt64();
		StartingSectorLocation = br.ReadInt32();
		StreamSize = br.ReadInt64();
	}

	internal void Write(BinaryWriter bw)
	{
		byte[] bytes = Encoding.Unicode.GetBytes(Name);
		bw.Write(bytes);
		bw.Write(new byte[64 - bytes.Length]);
		bw.Write((short)(bytes.Length + 2));
		bw.Write(ObjectType);
		bw.Write(ColorFlag);
		bw.Write(LeftSibling);
		bw.Write(RightSibling);
		bw.Write(ChildID);
		bw.Write(ClsID.ToByteArray());
		bw.Write(StatBits);
		bw.Write(CreationTime);
		bw.Write(ModifiedTime);
		bw.Write(StartingSectorLocation);
		bw.Write(StreamSize);
	}

	public override string ToString()
	{
		return Name;
	}

	public int CompareTo(CompoundDocumentItem other)
	{
		if (Name.Length < other.Name.Length)
		{
			return -1;
		}
		if (Name.Length > other.Name.Length)
		{
			return 1;
		}
		string text = Name.ToUpperInvariant();
		string text2 = other.Name.ToUpperInvariant();
		for (int i = 0; i < text.Length; i++)
		{
			if (text[i] < text2[i])
			{
				return -1;
			}
			if (text[i] > text2[i])
			{
				return 1;
			}
		}
		return 0;
	}
}
