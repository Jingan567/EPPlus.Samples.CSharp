using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeOpenXml.Utils.CompundDocument;

internal class CompoundDocumentFile : IDisposable
{
	private struct DocWriteInfo
	{
		internal List<int> DIFAT;

		internal List<int> FAT;

		internal List<int> miniFAT;
	}

	private const int miniFATSectorSize = 64;

	private const int FATSectorSizeV3 = 512;

	private const int FATSectorSizeV4 = 4096;

	private const int DIFAT_SECTOR = -4;

	private const int FAT_SECTOR = -3;

	private const int END_OF_CHAIN = -2;

	private const int FREE_SECTOR = -1;

	private static readonly byte[] header = new byte[8] { 208, 207, 17, 224, 161, 177, 26, 225 };

	private short minorVersion;

	private short majorVersion;

	private int numberOfDirectorySector;

	private short sectorShif;

	private short minSectorShift;

	private int _numberOfFATSectors;

	private int _firstDirectorySectorLocation;

	private int _transactionSignatureNumber;

	private int _miniStreamCutoffSize;

	private int _firstMiniFATSectorLocation;

	private int _numberofMiniFATSectors;

	private int _firstDIFATSectorLocation;

	private int _numberofDIFATSectors;

	private List<byte[]> _sectors;

	private List<byte[]> _miniSectors;

	private int _sectorSize;

	private int _miniSectorSize;

	private int _sectorSizeInt;

	private int _currentDIFATSectorPos;

	private int _currentFATSectorPos;

	private int _currentDirSectorPos;

	private int _prevDirFATSectorPos;

	public CompoundDocumentItem RootItem { get; set; }

	public CompoundDocumentFile()
	{
		RootItem = new CompoundDocumentItem
		{
			Name = "<Root>",
			Children = new List<CompoundDocumentItem>(),
			ObjectType = 5
		};
		minorVersion = 62;
		majorVersion = 3;
		sectorShif = 9;
		minSectorShift = 6;
		_sectorSize = 1 << (int)sectorShif;
		_miniSectorSize = 1 << (int)minSectorShift;
		_sectorSizeInt = _sectorSize / 4;
	}

	internal CompoundDocumentFile(FileInfo fi)
		: this(File.ReadAllBytes(fi.FullName))
	{
	}

	public CompoundDocumentFile(byte[] file)
		: this(new MemoryStream(file))
	{
	}

	public CompoundDocumentFile(MemoryStream ms)
	{
		ms.Seek(0L, SeekOrigin.Begin);
		Read(new BinaryReader(ms));
	}

	public static bool IsCompoundDocument(FileInfo fi)
	{
		try
		{
			FileStream fileStream = fi.OpenRead();
			byte[] array = new byte[8];
			fileStream.Read(array, 0, 8);
			return IsCompoundDocument(array);
		}
		catch
		{
			return false;
		}
	}

	public static bool IsCompoundDocument(MemoryStream ms)
	{
		long position = ms.Position;
		ms.Position = 0L;
		byte[] array = new byte[8];
		ms.Read(array, 0, 8);
		ms.Position = position;
		return IsCompoundDocument(array);
	}

	public static bool IsCompoundDocument(byte[] b)
	{
		if (b == null || b.Length < 8)
		{
			return false;
		}
		for (int i = 0; i < 8; i++)
		{
			if (b[i] != header[i])
			{
				return false;
			}
		}
		return true;
	}

	internal void Read(BinaryReader br)
	{
		br.ReadBytes(8);
		br.ReadBytes(16);
		minorVersion = br.ReadInt16();
		majorVersion = br.ReadInt16();
		br.ReadInt16();
		sectorShif = br.ReadInt16();
		minSectorShift = br.ReadInt16();
		_sectorSize = 1 << (int)sectorShif;
		_miniSectorSize = 1 << (int)minSectorShift;
		_sectorSizeInt = _sectorSize / 4;
		br.ReadBytes(6);
		numberOfDirectorySector = br.ReadInt32();
		_numberOfFATSectors = br.ReadInt32();
		_firstDirectorySectorLocation = br.ReadInt32();
		_transactionSignatureNumber = br.ReadInt32();
		_miniStreamCutoffSize = br.ReadInt32();
		_firstMiniFATSectorLocation = br.ReadInt32();
		_numberofMiniFATSectors = br.ReadInt32();
		_firstDIFATSectorLocation = br.ReadInt32();
		_numberofDIFATSectors = br.ReadInt32();
		DocWriteInfo docWriteInfo = default(DocWriteInfo);
		docWriteInfo.DIFAT = new List<int>();
		docWriteInfo.FAT = new List<int>();
		docWriteInfo.miniFAT = new List<int>();
		DocWriteInfo dwi = docWriteInfo;
		for (int i = 0; i < 109; i++)
		{
			int num = br.ReadInt32();
			if (num >= 0)
			{
				dwi.DIFAT.Add(num);
			}
		}
		LoadSectors(br);
		if (_firstDIFATSectorLocation > 0)
		{
			LoadDIFATSectors(dwi);
		}
		dwi.FAT = ReadFAT(_sectors, dwi);
		List<CompoundDocumentItem> list = ReadDirectories(_sectors, dwi);
		LoadMinSectors(ref dwi, list);
		foreach (CompoundDocumentItem item in list)
		{
			if (item.Stream == null && item.StreamSize > 0)
			{
				if (item.StreamSize < _miniStreamCutoffSize)
				{
					item.Stream = GetStream(item.StartingSectorLocation, item.StreamSize, dwi.miniFAT, _miniSectors);
				}
				else
				{
					item.Stream = GetStream(item.StartingSectorLocation, item.StreamSize, dwi.FAT, _sectors);
				}
			}
		}
		AddChildTree(list[0], list);
	}

	private void LoadDIFATSectors(DocWriteInfo dwi)
	{
		int num = _firstDIFATSectorLocation;
		while (num > 0)
		{
			BinaryReader binaryReader = new BinaryReader(new MemoryStream(_sectors[num]));
			int num2 = -1;
			while (binaryReader.BaseStream.Position < _sectorSize)
			{
				if (num2 > 0)
				{
					dwi.DIFAT.Add(num2);
				}
				num2 = binaryReader.ReadInt32();
			}
			num = num2;
		}
	}

	private void LoadSectors(BinaryReader br)
	{
		_sectors = new List<byte[]>();
		while (br.BaseStream.Position < br.BaseStream.Length)
		{
			_sectors.Add(br.ReadBytes(_sectorSize));
		}
	}

	private void LoadMinSectors(ref DocWriteInfo dwi, List<CompoundDocumentItem> dir)
	{
		dwi.miniFAT = ReadMiniFAT(_sectors, dwi);
		dir[0].Stream = GetStream(dir[0].StartingSectorLocation, dir[0].StreamSize, dwi.FAT, _sectors);
		GetMiniSectors(dir[0].Stream);
	}

	private void GetMiniSectors(byte[] miniFATStream)
	{
		BinaryReader binaryReader = new BinaryReader(new MemoryStream(miniFATStream));
		_miniSectors = new List<byte[]>();
		while (binaryReader.BaseStream.Position < binaryReader.BaseStream.Length)
		{
			_miniSectors.Add(binaryReader.ReadBytes(_miniSectorSize));
		}
	}

	private byte[] GetStream(int startingSectorLocation, long streamSize, List<int> FAT, List<byte[]> sectors)
	{
		MemoryStream memoryStream = new MemoryStream();
		BinaryWriter binaryWriter = new BinaryWriter(memoryStream);
		int num = 0;
		int index = startingSectorLocation;
		while (num < streamSize)
		{
			if (streamSize > num + sectors[index].Length)
			{
				binaryWriter.Write(sectors[index]);
				num += sectors[index].Length;
			}
			else
			{
				byte[] array = new byte[streamSize - num];
				Array.Copy(sectors[index], array, (int)streamSize - num);
				binaryWriter.Write(array);
				num += array.Length;
			}
			index = FAT[index];
		}
		binaryWriter.Flush();
		return memoryStream.ToArray();
	}

	private List<int> ReadMiniFAT(List<byte[]> sectors, DocWriteInfo dwi)
	{
		List<int> list = new List<int>();
		for (int num = _firstMiniFATSectorLocation; num != -2; num = dwi.FAT[num])
		{
			BinaryReader binaryReader = new BinaryReader(new MemoryStream(sectors[num]));
			while (binaryReader.BaseStream.Position < _sectorSize)
			{
				int item = binaryReader.ReadInt32();
				list.Add(item);
			}
		}
		return list;
	}

	private List<CompoundDocumentItem> ReadDirectories(List<byte[]> sectors, DocWriteInfo dwi)
	{
		List<CompoundDocumentItem> list = new List<CompoundDocumentItem>();
		for (int num = _firstDirectorySectorLocation; num != -2; num = dwi.FAT[num])
		{
			ReadDirectory(sectors, num, list);
		}
		return list;
	}

	private List<int> ReadFAT(List<byte[]> sectors, DocWriteInfo dwi)
	{
		List<int> list = new List<int>();
		foreach (int item2 in dwi.DIFAT)
		{
			BinaryReader binaryReader = new BinaryReader(new MemoryStream(sectors[item2]));
			while (binaryReader.BaseStream.Position < _sectorSize)
			{
				int item = binaryReader.ReadInt32();
				list.Add(item);
			}
		}
		return list;
	}

	private void ReadDirectory(List<byte[]> sectors, int index, List<CompoundDocumentItem> l)
	{
		BinaryReader binaryReader = new BinaryReader(new MemoryStream(sectors[index]));
		while (binaryReader.BaseStream.Position < binaryReader.BaseStream.Length)
		{
			CompoundDocumentItem compoundDocumentItem = new CompoundDocumentItem();
			compoundDocumentItem.Read(binaryReader);
			if (compoundDocumentItem.ObjectType != 0)
			{
				l.Add(compoundDocumentItem);
			}
		}
	}

	internal void AddChildTree(CompoundDocumentItem e, List<CompoundDocumentItem> dirs)
	{
		if (!e._handled)
		{
			e._handled = true;
			if (e.ChildID > 0)
			{
				CompoundDocumentItem compoundDocumentItem = dirs[e.ChildID];
				compoundDocumentItem.Parent = e;
				e.Children.Add(compoundDocumentItem);
				AddChildTree(compoundDocumentItem, dirs);
			}
			if (e.LeftSibling > 0)
			{
				CompoundDocumentItem compoundDocumentItem2 = dirs[e.LeftSibling];
				compoundDocumentItem2.Parent = e.Parent;
				compoundDocumentItem2.Parent.Children.Insert(e.Parent.Children.IndexOf(e), compoundDocumentItem2);
				AddChildTree(compoundDocumentItem2, dirs);
			}
			if (e.RightSibling > 0)
			{
				CompoundDocumentItem compoundDocumentItem3 = dirs[e.RightSibling];
				compoundDocumentItem3.Parent = e.Parent;
				e.Parent.Children.Insert(e.Parent.Children.IndexOf(e) + 1, compoundDocumentItem3);
				AddChildTree(compoundDocumentItem3, dirs);
			}
			if (e.ObjectType == 5)
			{
				RootItem = e;
			}
		}
	}

	internal void AddLeftSiblingTree(CompoundDocumentItem e, List<CompoundDocumentItem> dirs)
	{
		if (e.LeftSibling > 0)
		{
			CompoundDocumentItem compoundDocumentItem = dirs[e.LeftSibling];
			if (compoundDocumentItem.Parent != null)
			{
				compoundDocumentItem.Parent = e.Parent;
				compoundDocumentItem.Parent.Children.Insert(e.Parent.Children.IndexOf(e), compoundDocumentItem);
				e._handled = true;
				AddLeftSiblingTree(compoundDocumentItem, dirs);
			}
		}
	}

	internal void AddRightSiblingTree(CompoundDocumentItem e, List<CompoundDocumentItem> dirs)
	{
		if (e.RightSibling > 0)
		{
			CompoundDocumentItem compoundDocumentItem = dirs[e.RightSibling];
			compoundDocumentItem.Parent = e.Parent;
			e.Parent.Children.Insert(e.Parent.Children.IndexOf(e) + 1, compoundDocumentItem);
			e._handled = true;
			AddRightSiblingTree(compoundDocumentItem, dirs);
		}
	}

	public void Write(MemoryStream ms)
	{
		BinaryWriter binaryWriter = new BinaryWriter(ms);
		minorVersion = 62;
		majorVersion = 3;
		sectorShif = 9;
		minSectorShift = 6;
		_miniStreamCutoffSize = 4096;
		_transactionSignatureNumber = 0;
		_firstDIFATSectorLocation = -2;
		_firstDirectorySectorLocation = 1;
		_firstMiniFATSectorLocation = 2;
		_numberOfFATSectors = 1;
		_currentDIFATSectorPos = 76;
		_currentFATSectorPos = _sectorSize;
		_currentDirSectorPos = _sectorSize * 2;
		_prevDirFATSectorPos = _sectorSize + 4;
		binaryWriter.Write(new byte[2048]);
		WritePosition(binaryWriter, 0, ref _currentDIFATSectorPos, isFATEntry: false);
		WritePosition(binaryWriter, new int[3] { -3, -2, -2 }, ref _currentFATSectorPos);
		List<CompoundDocumentItem> dirs = FlattenDirs();
		WriteDirs(binaryWriter, dirs);
		FillDIFAT(binaryWriter);
		WriteHeader(binaryWriter);
	}

	private List<CompoundDocumentItem> FlattenDirs()
	{
		List<CompoundDocumentItem> list = new List<CompoundDocumentItem>();
		InitItem(RootItem);
		list.Add(RootItem);
		RootItem.ChildID = AddChildren(RootItem, list);
		return list;
	}

	private void InitItem(CompoundDocumentItem item)
	{
		item.LeftSibling = -1;
		item.RightSibling = -1;
		item._handled = false;
	}

	private int AddChildren(CompoundDocumentItem item, List<CompoundDocumentItem> l)
	{
		int result = -1;
		item.ColorFlag = 1;
		if (item.Children.Count > 0)
		{
			foreach (CompoundDocumentItem child in item.Children)
			{
				InitItem(child);
			}
			item.Children.Sort();
			result = SetSiblings(l.Count, item.Children, 0, item.Children.Count - 1, -1);
			l.AddRange(item.Children);
			foreach (CompoundDocumentItem child2 in item.Children)
			{
				child2.ChildID = AddChildren(child2, l);
			}
		}
		return result;
	}

	private void SetUnhandled(int listAdd, List<CompoundDocumentItem> children)
	{
		for (int i = 0; i < children.Count; i++)
		{
			if (children[i]._handled)
			{
				continue;
			}
			if (i > 0 && children[i - 1].RightSibling == -1 && children[i].LeftSibling != i + listAdd - 1)
			{
				children[i - 1].RightSibling = i + listAdd;
				continue;
			}
			if (i < children.Count - 1 && children[i + 1].LeftSibling == -1 && children[i].RightSibling != i + listAdd + 1)
			{
				children[i + 1].LeftSibling = i + listAdd;
				continue;
			}
			throw new InvalidOperationException("Invalid sibling handling in Document");
		}
	}

	private int SetSiblings(int listAdd, List<CompoundDocumentItem> children, int fromPos, int toPos, int currSibl)
	{
		int pos = GetPos(fromPos, toPos);
		CompoundDocumentItem compoundDocumentItem = children[pos];
		if (compoundDocumentItem._handled)
		{
			return currSibl;
		}
		compoundDocumentItem._handled = true;
		if (fromPos == toPos)
		{
			return fromPos + listAdd;
		}
		int num = pos / 2;
		if (num <= 0)
		{
			num = 1;
		}
		int pos2 = GetPos(fromPos, pos - 1);
		int pos3 = GetPos(pos + 1, toPos);
		if (num == 1 && children[pos2]._handled && children[pos3]._handled)
		{
			return pos + listAdd;
		}
		if (pos2 > -1 && pos2 >= fromPos)
		{
			compoundDocumentItem.LeftSibling = SetSiblings(listAdd, children, fromPos, pos - 1, compoundDocumentItem.LeftSibling);
		}
		if (pos3 < children.Count && pos3 <= toPos)
		{
			compoundDocumentItem.RightSibling = SetSiblings(listAdd, children, pos + 1, toPos, compoundDocumentItem.RightSibling);
		}
		return pos + listAdd;
	}

	private int GetPos(int fromPos, int toPos)
	{
		int num = (toPos - fromPos) / 2;
		return fromPos + num;
	}

	private bool NoGreater(List<CompoundDocumentItem> children, int pos, int lPos, int listAdd)
	{
		if (pos - lPos <= 1)
		{
			return true;
		}
		for (int i = lPos + 1; i <= pos; i++)
		{
			if (children[i].RightSibling != -1 && children[i].RightSibling > lPos + listAdd)
			{
				return false;
			}
		}
		return true;
	}

	private bool NoLess(List<CompoundDocumentItem> children, int pos, int rPos, int listAdd)
	{
		if (rPos - pos <= 1)
		{
			return true;
		}
		for (int i = pos + 1; i <= rPos; i++)
		{
			if (children[i].LeftSibling != -1 && children[i].LeftSibling < rPos + listAdd)
			{
				return false;
			}
		}
		return true;
	}

	private int GetLevels(int c)
	{
		c--;
		int num = 0;
		while (c > 0)
		{
			c >>= 1;
			num++;
		}
		return num;
	}

	private void FillDIFAT(BinaryWriter bw)
	{
		if (_currentDIFATSectorPos >= _sectorSize)
		{
			return;
		}
		bw.Seek(_currentDIFATSectorPos, SeekOrigin.Begin);
		while (_currentDIFATSectorPos < _sectorSize)
		{
			if (_currentDIFATSectorPos < 512)
			{
				bw.Write(uint.MaxValue);
			}
			else
			{
				bw.Write(0);
			}
			_currentDIFATSectorPos += 4;
		}
	}

	private void WritePosition(BinaryWriter bw, int sector, ref int writePos, bool isFATEntry)
	{
		int offset = (int)bw.BaseStream.Position;
		bw.Seek(writePos, SeekOrigin.Begin);
		bw.Write(sector);
		writePos += 4;
		if (isFATEntry)
		{
			CheckUpdateDIFAT(bw);
		}
		bw.Seek(offset, SeekOrigin.Begin);
	}

	private void WritePosition(BinaryWriter bw, int[] sectors, ref int writePos)
	{
		int offset = (int)bw.BaseStream.Position;
		bw.Seek(writePos, SeekOrigin.Begin);
		foreach (int value in sectors)
		{
			bw.Write(value);
			writePos += 4;
		}
		bw.Seek(offset, SeekOrigin.Begin);
	}

	private void WriteDirs(BinaryWriter bw, List<CompoundDocumentItem> dirs)
	{
		byte[] array = SetMiniStream(dirs);
		AllocateFAT(bw, array.Length, dirs);
		WriteMiniFAT(bw, array);
		foreach (CompoundDocumentItem dir in dirs)
		{
			if (dir.ObjectType == 5 || dir.StreamSize > _miniStreamCutoffSize)
			{
				dir.StartingSectorLocation = WriteStream(bw, dir.Stream);
			}
		}
		WriteDirStream(bw, dirs);
	}

	private int WriteDirStream(BinaryWriter bw, List<CompoundDocumentItem> dirs)
	{
		if (dirs.Count > 0)
		{
			bw.Seek((_firstDirectorySectorLocation + 1) * _sectorSize, SeekOrigin.Begin);
			for (int i = 0; i < Math.Min(_sectorSize / 128, dirs.Count); i++)
			{
				dirs[i].Write(bw);
			}
			bw.Seek(0, SeekOrigin.End);
			int num = (int)bw.BaseStream.Position / _sectorSize - 1;
			int writePos = _sectorSize + 4;
			WritePosition(bw, num, ref writePos, isFATEntry: false);
			int num2 = 0;
			for (int j = 4; j < dirs.Count; j++)
			{
				dirs[j].Write(bw);
				num2 += 128;
			}
			WriteStreamFullSector(bw, _sectorSize);
			WriteFAT(bw, num, num2);
			return num;
		}
		return -1;
	}

	private void WriteMiniFAT(BinaryWriter bw, byte[] miniFAT)
	{
		if (miniFAT.Length >= _sectorSize)
		{
			bw.Seek((_firstMiniFATSectorLocation + 1) * _sectorSize, SeekOrigin.Begin);
			bw.Write(miniFAT, 0, _sectorSize);
			bw.Seek(0, SeekOrigin.End);
			if (miniFAT.Length > _sectorSize)
			{
				int sector = (int)bw.BaseStream.Position / _sectorSize - 1;
				int writePos = _sectorSize + 8;
				WritePosition(bw, sector, ref writePos, isFATEntry: false);
				byte[] array = new byte[miniFAT.Length - _sectorSize];
				Array.Copy(miniFAT, _sectorSize, array, 0, array.Length);
				WriteStream(bw, array);
			}
			_numberofMiniFATSectors = (miniFAT.Length + 1) / _sectorSize;
		}
	}

	private int WriteStream(BinaryWriter bw, byte[] stream)
	{
		bw.Seek(0, SeekOrigin.End);
		int num = (int)bw.BaseStream.Position / _sectorSize - 1;
		bw.Write(stream);
		WriteStreamFullSector(bw, _sectorSize);
		WriteFAT(bw, num, stream.Length);
		return num;
	}

	private void WriteFAT(BinaryWriter bw, int sector, long size)
	{
		bw.Seek(_currentFATSectorPos, SeekOrigin.Begin);
		int num = _sectorSize;
		while (size > num)
		{
			bw.Write(++sector);
			num += _sectorSize;
			CheckUpdateDIFAT(bw);
		}
		bw.Write(-2);
		CheckUpdateDIFAT(bw);
		_currentFATSectorPos = (int)bw.BaseStream.Position;
		bw.Seek(0, SeekOrigin.End);
	}

	private void CheckUpdateDIFAT(BinaryWriter bw)
	{
		if (bw.BaseStream.Position % _sectorSize != 0L)
		{
			return;
		}
		if (_currentDIFATSectorPos % _sectorSize == 0)
		{
			bw.Seek(512, SeekOrigin.Current);
		}
		else if (bw.BaseStream.Position == _sectorSize * 2)
		{
			bw.Seek(4 * _sectorSize, SeekOrigin.Begin);
		}
		int num = (int)(bw.BaseStream.Position / _sectorSize - 1);
		WritePosition(bw, num, ref _currentDIFATSectorPos, isFATEntry: false);
		_numberOfFATSectors++;
		if (_currentDIFATSectorPos == _sectorSize || ((_currentDIFATSectorPos + 4) % _sectorSize == 0 && _currentDIFATSectorPos > _sectorSize))
		{
			bw.Write(new byte[_sectorSize]);
			if (_currentDIFATSectorPos > _sectorSize)
			{
				WritePosition(bw, num + 1, ref _currentDIFATSectorPos, isFATEntry: false);
			}
			else
			{
				_firstDIFATSectorLocation = num + 1;
			}
			_currentDIFATSectorPos = (int)bw.BaseStream.Position;
			for (int i = 0; i < _sectorSize; i++)
			{
				bw.Write(byte.MaxValue);
			}
			bw.Seek(-(_sectorSize * 2), SeekOrigin.Current);
		}
	}

	private void AllocateFAT(BinaryWriter bw, int miniFatLength, List<CompoundDocumentItem> dirs)
	{
		long num = (long)miniFatLength - (long)_sectorSize;
		foreach (CompoundDocumentItem dir in dirs)
		{
			if (dir.ObjectType == 5 || dir.StreamSize > _miniStreamCutoffSize)
			{
				long num2 = _sectorSize - dir.StreamSize % _sectorSize;
				num += dir.StreamSize;
				if (num2 > 0 && num2 < _sectorSize)
				{
					num += num2;
				}
			}
		}
		long num3 = num / _sectorSize;
		int num4 = _sectorSize / 128;
		int num5 = 0;
		_ = _currentFATSectorPos;
		if (dirs.Count > num4)
		{
			num5 = GetSectors(dirs.Count, num4);
			num3 += num5 - 1;
		}
		int sectors = GetSectors((int)num3, _sectorSizeInt);
		_numberofDIFATSectors = GetDIFatSectors(sectors);
		num3 += sectors + _numberofDIFATSectors;
		sectors = GetSectors((int)num3, _sectorSizeInt) + _numberofDIFATSectors;
		_numberofDIFATSectors = GetDIFatSectors(sectors);
		bw.Write(new byte[(sectors + ((_numberofDIFATSectors > 0) ? (_numberofDIFATSectors - 1) : 0)) * _sectorSize]);
		bw.Seek(_currentFATSectorPos, SeekOrigin.Begin);
		int num6 = 1;
		for (int i = 1; i < 109; i++)
		{
			if (i < sectors + _numberofDIFATSectors)
			{
				WriteFATItem(bw, -3);
				num6++;
				continue;
			}
			WriteFATItem(bw, -2);
			break;
		}
		if (_numberofDIFATSectors > 0)
		{
			_firstDIFATSectorLocation = num6 + 1;
		}
		for (int j = 0; j < _numberofDIFATSectors; j++)
		{
			WriteFATItem(bw, -4);
			for (int k = 0; k < _sectorSizeInt - 1; k++)
			{
				WriteFATItem(bw, -3);
				num6++;
				if (num6 >= sectors)
				{
					break;
				}
			}
			if (num6 > sectors)
			{
				break;
			}
		}
		bw.Seek(0, SeekOrigin.End);
	}

	private int GetDIFatSectors(int FATSectors)
	{
		if (FATSectors > 109)
		{
			return GetSectors(FATSectors - 109, _sectorSizeInt - 1);
		}
		return 0;
	}

	private void WriteFATItem(BinaryWriter bw, int value)
	{
		bw.Write(value);
		CheckUpdateDIFAT(bw);
		_currentFATSectorPos = (int)bw.BaseStream.Position;
	}

	private int GetSectors(int v, int size)
	{
		if (v % size == 0)
		{
			return v / size;
		}
		return v / size + 1;
	}

	private byte[] SetMiniStream(List<CompoundDocumentItem> dirs)
	{
		MemoryStream memoryStream = new MemoryStream();
		BinaryWriter binaryWriter = new BinaryWriter(memoryStream);
		BinaryWriter binaryWriter2 = new BinaryWriter(new MemoryStream());
		int num = 0;
		foreach (CompoundDocumentItem dir in dirs)
		{
			if (dir.ObjectType != 5 && dir.StreamSize > 0 && dir.StreamSize <= _miniStreamCutoffSize)
			{
				binaryWriter.Write(dir.Stream);
				WriteStreamFullSector(binaryWriter, 64);
				int i = _miniSectorSize;
				dir.StartingSectorLocation = num;
				for (; dir.StreamSize > i; i += _miniSectorSize)
				{
					binaryWriter2.Write(++num);
				}
				binaryWriter2.Write(-2);
				num++;
			}
		}
		dirs[0].StreamSize = memoryStream.Length;
		dirs[0].Stream = memoryStream.ToArray();
		WriteStreamFullSector(binaryWriter2, _sectorSize);
		return ((MemoryStream)binaryWriter2.BaseStream).ToArray();
	}

	private static void WriteStreamFullSector(BinaryWriter bw, int sectorSize)
	{
		long num = sectorSize - bw.BaseStream.Length % sectorSize;
		if (num > 0 && num < sectorSize)
		{
			bw.Write(new byte[num]);
		}
	}

	private void WriteHeader(BinaryWriter bw)
	{
		bw.Seek(0, SeekOrigin.Begin);
		bw.Write(header);
		bw.Write(new byte[16]);
		bw.Write((short)62);
		bw.Write((short)3);
		bw.Write((ushort)65534);
		bw.Write((short)9);
		bw.Write((short)6);
		bw.Write(new byte[6]);
		bw.Write(0);
		bw.Write(_numberOfFATSectors);
		bw.Write(1);
		bw.Write(0);
		bw.Write(_miniStreamCutoffSize);
		bw.Write(2);
		bw.Write(_numberofMiniFATSectors);
		bw.Write(_firstDIFATSectorLocation);
		bw.Write(_numberofDIFATSectors);
	}

	private void CreateFATStreams(CompoundDocumentItem item, BinaryWriter bw, BinaryWriter bwMini, DocWriteInfo dwi)
	{
		if (item.ObjectType != 5 && item.StreamSize > 0)
		{
			item.StreamSize = item.Stream.Length;
			if (item.StreamSize < _miniStreamCutoffSize)
			{
				item.StartingSectorLocation = WriteStream(bwMini, dwi.miniFAT, item.Stream, 64);
			}
			else
			{
				item.StartingSectorLocation = WriteStream(bw, dwi.FAT, item.Stream, 512);
			}
		}
		foreach (CompoundDocumentItem child in item.Children)
		{
			CreateFATStreams(child, bw, bwMini, dwi);
		}
	}

	private int WriteStream(BinaryWriter bw, List<int> fat, byte[] stream, int FATSectorSize)
	{
		int num = FATSectorSize - stream.Length % FATSectorSize;
		bw.Write(stream);
		if (num > 0 && num < FATSectorSize)
		{
			bw.Write(new byte[num]);
		}
		int count = fat.Count;
		AddFAT(fat, stream.Length, FATSectorSize, 0);
		return count;
	}

	private void AddFAT(List<int> fat, long streamSize, int sectorSize, int addPos)
	{
		for (int i = 0; i < streamSize; i += sectorSize)
		{
			if (i + sectorSize < streamSize)
			{
				fat.Add(fat.Count + 1);
			}
			else
			{
				fat.Add(-2);
			}
		}
	}

	public void Dispose()
	{
		_miniSectors = null;
		_sectors = null;
	}
}
