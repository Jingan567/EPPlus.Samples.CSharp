using System;
using System.Collections.Generic;

namespace OfficeOpenXml;

public static class FontSize
{
	public static readonly Dictionary<string, Dictionary<float, FontSizeInfo>> FontHeights = new Dictionary<string, Dictionary<float, FontSizeInfo>>(StringComparer.OrdinalIgnoreCase)
	{
		{
			"Times New Roman",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(20f, 7f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 9f)
				},
				{
					16f,
					new FontSizeInfo(27f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 12f)
				},
				{
					20f,
					new FontSizeInfo(35f, 14f)
				},
				{
					22f,
					new FontSizeInfo(37f, 14f)
				},
				{
					24f,
					new FontSizeInfo(41f, 16f)
				},
				{
					26f,
					new FontSizeInfo(44f, 18f)
				},
				{
					28f,
					new FontSizeInfo(47f, 19f)
				},
				{
					36f,
					new FontSizeInfo(61f, 24f)
				},
				{
					48f,
					new FontSizeInfo(82f, 32f)
				},
				{
					72f,
					new FontSizeInfo(122f, 48f)
				},
				{
					96f,
					new FontSizeInfo(164f, 64f)
				},
				{
					128f,
					new FontSizeInfo(216f, 86f)
				},
				{
					256f,
					new FontSizeInfo(428f, 171f)
				}
			}
		},
		{
			"Arial",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 13f)
				},
				{
					20f,
					new FontSizeInfo(34f, 15f)
				},
				{
					22f,
					new FontSizeInfo(36f, 16f)
				},
				{
					24f,
					new FontSizeInfo(40f, 18f)
				},
				{
					26f,
					new FontSizeInfo(44f, 19f)
				},
				{
					28f,
					new FontSizeInfo(46f, 21f)
				},
				{
					36f,
					new FontSizeInfo(59f, 27f)
				},
				{
					48f,
					new FontSizeInfo(79f, 36f)
				},
				{
					72f,
					new FontSizeInfo(120f, 53f)
				},
				{
					96f,
					new FontSizeInfo(159f, 71f)
				},
				{
					128f,
					new FontSizeInfo(213f, 95f)
				},
				{
					256f,
					new FontSizeInfo(424f, 190f)
				}
			}
		},
		{
			"Courier New",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(25f, 11f)
				},
				{
					16f,
					new FontSizeInfo(28f, 13f)
				},
				{
					18f,
					new FontSizeInfo(32f, 14f)
				},
				{
					20f,
					new FontSizeInfo(35f, 16f)
				},
				{
					22f,
					new FontSizeInfo(38f, 17f)
				},
				{
					24f,
					new FontSizeInfo(42f, 19f)
				},
				{
					26f,
					new FontSizeInfo(46f, 21f)
				},
				{
					28f,
					new FontSizeInfo(48f, 22f)
				},
				{
					36f,
					new FontSizeInfo(62f, 29f)
				},
				{
					48f,
					new FontSizeInfo(83f, 38f)
				},
				{
					72f,
					new FontSizeInfo(123f, 58f)
				},
				{
					96f,
					new FontSizeInfo(164f, 77f)
				},
				{
					128f,
					new FontSizeInfo(219f, 103f)
				},
				{
					256f,
					new FontSizeInfo(444f, 205f)
				}
			}
		},
		{
			"Symbol",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 4f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(24f, 10f)
				},
				{
					16f,
					new FontSizeInfo(29f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 12f)
				},
				{
					20f,
					new FontSizeInfo(35f, 14f)
				},
				{
					22f,
					new FontSizeInfo(38f, 15f)
				},
				{
					24f,
					new FontSizeInfo(40f, 16f)
				},
				{
					26f,
					new FontSizeInfo(44f, 18f)
				},
				{
					28f,
					new FontSizeInfo(48f, 19f)
				},
				{
					36f,
					new FontSizeInfo(60f, 24f)
				},
				{
					48f,
					new FontSizeInfo(79f, 32f)
				},
				{
					72f,
					new FontSizeInfo(120f, 48f)
				},
				{
					96f,
					new FontSizeInfo(159f, 64f)
				},
				{
					128f,
					new FontSizeInfo(212f, 86f)
				},
				{
					256f,
					new FontSizeInfo(421f, 171f)
				}
			}
		},
		{
			"Wingdings",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 11f)
				},
				{
					8f,
					new FontSizeInfo(20f, 15f)
				},
				{
					10f,
					new FontSizeInfo(20f, 17f)
				},
				{
					11f,
					new FontSizeInfo(20f, 20f)
				},
				{
					12f,
					new FontSizeInfo(21f, 22f)
				},
				{
					14f,
					new FontSizeInfo(24f, 26f)
				},
				{
					16f,
					new FontSizeInfo(26f, 28f)
				},
				{
					18f,
					new FontSizeInfo(30f, 32f)
				},
				{
					20f,
					new FontSizeInfo(34f, 36f)
				},
				{
					22f,
					new FontSizeInfo(36f, 39f)
				},
				{
					24f,
					new FontSizeInfo(40f, 43f)
				},
				{
					26f,
					new FontSizeInfo(43f, 47f)
				},
				{
					28f,
					new FontSizeInfo(46f, 50f)
				},
				{
					36f,
					new FontSizeInfo(59f, 64f)
				},
				{
					48f,
					new FontSizeInfo(79f, 86f)
				},
				{
					72f,
					new FontSizeInfo(117f, 129f)
				},
				{
					96f,
					new FontSizeInfo(156f, 172f)
				},
				{
					128f,
					new FontSizeInfo(208f, 230f)
				},
				{
					256f,
					new FontSizeInfo(414f, 457f)
				}
			}
		},
		{
			"SimSun",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(27f, 11f)
				},
				{
					18f,
					new FontSizeInfo(30f, 12f)
				},
				{
					20f,
					new FontSizeInfo(34f, 14f)
				},
				{
					22f,
					new FontSizeInfo(36f, 15f)
				},
				{
					24f,
					new FontSizeInfo(42f, 16f)
				},
				{
					26f,
					new FontSizeInfo(45f, 18f)
				},
				{
					28f,
					new FontSizeInfo(47f, 19f)
				},
				{
					36f,
					new FontSizeInfo(62f, 24f)
				},
				{
					48f,
					new FontSizeInfo(82f, 32f)
				},
				{
					72f,
					new FontSizeInfo(123f, 48f)
				},
				{
					96f,
					new FontSizeInfo(163f, 64f)
				},
				{
					128f,
					new FontSizeInfo(218f, 86f)
				},
				{
					256f,
					new FontSizeInfo(436f, 171f)
				}
			}
		},
		{
			"Century",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(30f, 13f)
				},
				{
					20f,
					new FontSizeInfo(34f, 15f)
				},
				{
					22f,
					new FontSizeInfo(36f, 16f)
				},
				{
					24f,
					new FontSizeInfo(40f, 18f)
				},
				{
					26f,
					new FontSizeInfo(44f, 19f)
				},
				{
					28f,
					new FontSizeInfo(46f, 21f)
				},
				{
					36f,
					new FontSizeInfo(59f, 27f)
				},
				{
					48f,
					new FontSizeInfo(79f, 36f)
				},
				{
					72f,
					new FontSizeInfo(118f, 53f)
				},
				{
					96f,
					new FontSizeInfo(157f, 71f)
				},
				{
					128f,
					new FontSizeInfo(209f, 95f)
				},
				{
					256f,
					new FontSizeInfo(416f, 190f)
				}
			}
		},
		{
			"Sylfaen",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(21f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(24f, 8f)
				},
				{
					14f,
					new FontSizeInfo(26f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(32f, 12f)
				},
				{
					20f,
					new FontSizeInfo(36f, 14f)
				},
				{
					22f,
					new FontSizeInfo(41f, 15f)
				},
				{
					24f,
					new FontSizeInfo(44f, 16f)
				},
				{
					26f,
					new FontSizeInfo(48f, 18f)
				},
				{
					28f,
					new FontSizeInfo(49f, 19f)
				},
				{
					36f,
					new FontSizeInfo(63f, 24f)
				},
				{
					48f,
					new FontSizeInfo(86f, 32f)
				},
				{
					72f,
					new FontSizeInfo(129f, 48f)
				},
				{
					96f,
					new FontSizeInfo(167f, 64f)
				},
				{
					128f,
					new FontSizeInfo(224f, 86f)
				},
				{
					256f,
					new FontSizeInfo(452f, 171f)
				}
			}
		},
		{
			"Cambria Math",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(85f, 6f)
				},
				{
					10f,
					new FontSizeInfo(102f, 7f)
				},
				{
					11f,
					new FontSizeInfo(117f, 8f)
				},
				{
					12f,
					new FontSizeInfo(124f, 9f)
				},
				{
					14f,
					new FontSizeInfo(147f, 11f)
				},
				{
					16f,
					new FontSizeInfo(162f, 12f)
				},
				{
					18f,
					new FontSizeInfo(186f, 13f)
				},
				{
					20f,
					new FontSizeInfo(209f, 15f)
				},
				{
					22f,
					new FontSizeInfo(223f, 16f)
				},
				{
					24f,
					new FontSizeInfo(248f, 18f)
				},
				{
					26f,
					new FontSizeInfo(270f, 19f)
				},
				{
					28f,
					new FontSizeInfo(285f, 20f)
				},
				{
					36f,
					new FontSizeInfo(371f, 27f)
				},
				{
					48f,
					new FontSizeInfo(493f, 35f)
				},
				{
					72f,
					new FontSizeInfo(739f, 53f)
				},
				{
					96f,
					new FontSizeInfo(986f, 71f)
				},
				{
					128f,
					new FontSizeInfo(1317f, 95f)
				},
				{
					256f,
					new FontSizeInfo(2047f, 189f)
				}
			}
		},
		{
			"Yu Gothic",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(22f, 7f)
				},
				{
					11f,
					new FontSizeInfo(25f, 8f)
				},
				{
					12f,
					new FontSizeInfo(26f, 9f)
				},
				{
					14f,
					new FontSizeInfo(32f, 11f)
				},
				{
					16f,
					new FontSizeInfo(34f, 12f)
				},
				{
					18f,
					new FontSizeInfo(40f, 13f)
				},
				{
					20f,
					new FontSizeInfo(44f, 15f)
				},
				{
					22f,
					new FontSizeInfo(47f, 16f)
				},
				{
					24f,
					new FontSizeInfo(53f, 18f)
				},
				{
					26f,
					new FontSizeInfo(57f, 19f)
				},
				{
					28f,
					new FontSizeInfo(59f, 21f)
				},
				{
					36f,
					new FontSizeInfo(78f, 27f)
				},
				{
					48f,
					new FontSizeInfo(102f, 36f)
				},
				{
					72f,
					new FontSizeInfo(154f, 53f)
				},
				{
					96f,
					new FontSizeInfo(206f, 71f)
				},
				{
					128f,
					new FontSizeInfo(276f, 95f)
				},
				{
					256f,
					new FontSizeInfo(551f, 190f)
				}
			}
		},
		{
			"DengXian",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 11f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 14f)
				},
				{
					20f,
					new FontSizeInfo(35f, 16f)
				},
				{
					22f,
					new FontSizeInfo(37f, 17f)
				},
				{
					24f,
					new FontSizeInfo(41f, 19f)
				},
				{
					26f,
					new FontSizeInfo(45f, 20f)
				},
				{
					28f,
					new FontSizeInfo(47f, 21f)
				},
				{
					36f,
					new FontSizeInfo(61f, 28f)
				},
				{
					48f,
					new FontSizeInfo(81f, 37f)
				},
				{
					72f,
					new FontSizeInfo(121f, 56f)
				},
				{
					96f,
					new FontSizeInfo(161f, 74f)
				},
				{
					128f,
					new FontSizeInfo(215f, 99f)
				},
				{
					256f,
					new FontSizeInfo(428f, 197f)
				}
			}
		},
		{
			"Arial Unicode MS",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(21f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 11f)
				},
				{
					16f,
					new FontSizeInfo(30f, 12f)
				},
				{
					18f,
					new FontSizeInfo(36f, 13f)
				},
				{
					20f,
					new FontSizeInfo(39f, 15f)
				},
				{
					22f,
					new FontSizeInfo(42f, 16f)
				},
				{
					24f,
					new FontSizeInfo(46f, 18f)
				},
				{
					26f,
					new FontSizeInfo(49f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 21f)
				},
				{
					36f,
					new FontSizeInfo(68f, 27f)
				},
				{
					48f,
					new FontSizeInfo(90f, 36f)
				},
				{
					72f,
					new FontSizeInfo(137f, 53f)
				},
				{
					96f,
					new FontSizeInfo(182f, 71f)
				},
				{
					128f,
					new FontSizeInfo(242f, 95f)
				},
				{
					256f,
					new FontSizeInfo(480f, 190f)
				}
			}
		},
		{
			"Calibri Light",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 12f)
				},
				{
					20f,
					new FontSizeInfo(35f, 14f)
				},
				{
					22f,
					new FontSizeInfo(38f, 15f)
				},
				{
					24f,
					new FontSizeInfo(42f, 16f)
				},
				{
					26f,
					new FontSizeInfo(45f, 18f)
				},
				{
					28f,
					new FontSizeInfo(48f, 19f)
				},
				{
					36f,
					new FontSizeInfo(62f, 24f)
				},
				{
					48f,
					new FontSizeInfo(82f, 32f)
				},
				{
					72f,
					new FontSizeInfo(123f, 49f)
				},
				{
					96f,
					new FontSizeInfo(163f, 65f)
				},
				{
					128f,
					new FontSizeInfo(218f, 87f)
				},
				{
					256f,
					new FontSizeInfo(434f, 173f)
				}
			}
		},
		{
			"Calibri",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 7f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 12f)
				},
				{
					20f,
					new FontSizeInfo(35f, 14f)
				},
				{
					22f,
					new FontSizeInfo(38f, 15f)
				},
				{
					24f,
					new FontSizeInfo(42f, 16f)
				},
				{
					26f,
					new FontSizeInfo(45f, 18f)
				},
				{
					28f,
					new FontSizeInfo(48f, 19f)
				},
				{
					36f,
					new FontSizeInfo(62f, 24f)
				},
				{
					48f,
					new FontSizeInfo(82f, 32f)
				},
				{
					72f,
					new FontSizeInfo(123f, 49f)
				},
				{
					96f,
					new FontSizeInfo(163f, 65f)
				},
				{
					128f,
					new FontSizeInfo(218f, 87f)
				},
				{
					256f,
					new FontSizeInfo(434f, 173f)
				}
			}
		},
		{
			"Georgia",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 8f)
				},
				{
					10f,
					new FontSizeInfo(20f, 9f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 12f)
				},
				{
					16f,
					new FontSizeInfo(27f, 13f)
				},
				{
					18f,
					new FontSizeInfo(31f, 15f)
				},
				{
					20f,
					new FontSizeInfo(34f, 17f)
				},
				{
					22f,
					new FontSizeInfo(36f, 18f)
				},
				{
					24f,
					new FontSizeInfo(40f, 20f)
				},
				{
					26f,
					new FontSizeInfo(44f, 21f)
				},
				{
					28f,
					new FontSizeInfo(46f, 23f)
				},
				{
					36f,
					new FontSizeInfo(60f, 29f)
				},
				{
					48f,
					new FontSizeInfo(79f, 39f)
				},
				{
					72f,
					new FontSizeInfo(118f, 59f)
				},
				{
					96f,
					new FontSizeInfo(157f, 79f)
				},
				{
					128f,
					new FontSizeInfo(209f, 105f)
				},
				{
					256f,
					new FontSizeInfo(417f, 209f)
				}
			}
		},
		{
			"Cambria",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(30f, 13f)
				},
				{
					20f,
					new FontSizeInfo(34f, 15f)
				},
				{
					22f,
					new FontSizeInfo(36f, 16f)
				},
				{
					24f,
					new FontSizeInfo(40f, 18f)
				},
				{
					26f,
					new FontSizeInfo(44f, 19f)
				},
				{
					28f,
					new FontSizeInfo(46f, 20f)
				},
				{
					36f,
					new FontSizeInfo(60f, 27f)
				},
				{
					48f,
					new FontSizeInfo(79f, 35f)
				},
				{
					72f,
					new FontSizeInfo(118f, 53f)
				},
				{
					96f,
					new FontSizeInfo(157f, 71f)
				},
				{
					128f,
					new FontSizeInfo(210f, 95f)
				},
				{
					256f,
					new FontSizeInfo(418f, 189f)
				}
			}
		},
		{
			"Marlett",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 8f)
				},
				{
					8f,
					new FontSizeInfo(20f, 11f)
				},
				{
					10f,
					new FontSizeInfo(20f, 13f)
				},
				{
					11f,
					new FontSizeInfo(22f, 15f)
				},
				{
					12f,
					new FontSizeInfo(23f, 16f)
				},
				{
					14f,
					new FontSizeInfo(26f, 19f)
				},
				{
					16f,
					new FontSizeInfo(28f, 21f)
				},
				{
					18f,
					new FontSizeInfo(31f, 24f)
				},
				{
					20f,
					new FontSizeInfo(34f, 27f)
				},
				{
					22f,
					new FontSizeInfo(36f, 29f)
				},
				{
					24f,
					new FontSizeInfo(39f, 32f)
				},
				{
					26f,
					new FontSizeInfo(42f, 35f)
				},
				{
					28f,
					new FontSizeInfo(44f, 37f)
				},
				{
					36f,
					new FontSizeInfo(55f, 48f)
				},
				{
					48f,
					new FontSizeInfo(71f, 64f)
				},
				{
					72f,
					new FontSizeInfo(103f, 96f)
				},
				{
					96f,
					new FontSizeInfo(135f, 128f)
				},
				{
					128f,
					new FontSizeInfo(178f, 171f)
				},
				{
					256f,
					new FontSizeInfo(348f, 341f)
				}
			}
		},
		{
			"Arial Black",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(21f, 9f)
				},
				{
					11f,
					new FontSizeInfo(25f, 10f)
				},
				{
					12f,
					new FontSizeInfo(26f, 11f)
				},
				{
					14f,
					new FontSizeInfo(30f, 13f)
				},
				{
					16f,
					new FontSizeInfo(33f, 14f)
				},
				{
					18f,
					new FontSizeInfo(36f, 16f)
				},
				{
					20f,
					new FontSizeInfo(42f, 18f)
				},
				{
					22f,
					new FontSizeInfo(45f, 19f)
				},
				{
					24f,
					new FontSizeInfo(49f, 21f)
				},
				{
					26f,
					new FontSizeInfo(55f, 23f)
				},
				{
					28f,
					new FontSizeInfo(57f, 25f)
				},
				{
					36f,
					new FontSizeInfo(74f, 32f)
				},
				{
					48f,
					new FontSizeInfo(97f, 43f)
				},
				{
					72f,
					new FontSizeInfo(147f, 64f)
				},
				{
					96f,
					new FontSizeInfo(195f, 85f)
				},
				{
					128f,
					new FontSizeInfo(259f, 114f)
				},
				{
					256f,
					new FontSizeInfo(516f, 227f)
				}
			}
		},
		{
			"Candara",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 13f)
				},
				{
					20f,
					new FontSizeInfo(35f, 15f)
				},
				{
					22f,
					new FontSizeInfo(38f, 16f)
				},
				{
					24f,
					new FontSizeInfo(42f, 18f)
				},
				{
					26f,
					new FontSizeInfo(45f, 19f)
				},
				{
					28f,
					new FontSizeInfo(48f, 20f)
				},
				{
					36f,
					new FontSizeInfo(62f, 26f)
				},
				{
					48f,
					new FontSizeInfo(82f, 35f)
				},
				{
					72f,
					new FontSizeInfo(123f, 53f)
				},
				{
					96f,
					new FontSizeInfo(163f, 71f)
				},
				{
					128f,
					new FontSizeInfo(218f, 94f)
				},
				{
					256f,
					new FontSizeInfo(434f, 188f)
				}
			}
		},
		{
			"Comic Sans MS",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(21f, 8f)
				},
				{
					11f,
					new FontSizeInfo(22f, 9f)
				},
				{
					12f,
					new FontSizeInfo(26f, 10f)
				},
				{
					14f,
					new FontSizeInfo(28f, 12f)
				},
				{
					16f,
					new FontSizeInfo(32f, 13f)
				},
				{
					18f,
					new FontSizeInfo(36f, 15f)
				},
				{
					20f,
					new FontSizeInfo(42f, 16f)
				},
				{
					22f,
					new FontSizeInfo(44f, 18f)
				},
				{
					24f,
					new FontSizeInfo(50f, 20f)
				},
				{
					26f,
					new FontSizeInfo(54f, 21f)
				},
				{
					28f,
					new FontSizeInfo(57f, 23f)
				},
				{
					36f,
					new FontSizeInfo(73f, 29f)
				},
				{
					48f,
					new FontSizeInfo(98f, 39f)
				},
				{
					72f,
					new FontSizeInfo(147f, 59f)
				},
				{
					96f,
					new FontSizeInfo(190f, 78f)
				},
				{
					128f,
					new FontSizeInfo(258f, 104f)
				},
				{
					256f,
					new FontSizeInfo(511f, 208f)
				}
			}
		},
		{
			"Consolas",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 13f)
				},
				{
					20f,
					new FontSizeInfo(35f, 15f)
				},
				{
					22f,
					new FontSizeInfo(37f, 16f)
				},
				{
					24f,
					new FontSizeInfo(41f, 18f)
				},
				{
					26f,
					new FontSizeInfo(45f, 19f)
				},
				{
					28f,
					new FontSizeInfo(47f, 20f)
				},
				{
					36f,
					new FontSizeInfo(61f, 26f)
				},
				{
					48f,
					new FontSizeInfo(81f, 35f)
				},
				{
					72f,
					new FontSizeInfo(121f, 53f)
				},
				{
					96f,
					new FontSizeInfo(161f, 70f)
				},
				{
					128f,
					new FontSizeInfo(215f, 94f)
				},
				{
					256f,
					new FontSizeInfo(428f, 187f)
				}
			}
		},
		{
			"Constantia",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 13f)
				},
				{
					20f,
					new FontSizeInfo(35f, 15f)
				},
				{
					22f,
					new FontSizeInfo(38f, 16f)
				},
				{
					24f,
					new FontSizeInfo(42f, 17f)
				},
				{
					26f,
					new FontSizeInfo(45f, 19f)
				},
				{
					28f,
					new FontSizeInfo(48f, 20f)
				},
				{
					36f,
					new FontSizeInfo(62f, 26f)
				},
				{
					48f,
					new FontSizeInfo(82f, 35f)
				},
				{
					72f,
					new FontSizeInfo(123f, 52f)
				},
				{
					96f,
					new FontSizeInfo(163f, 70f)
				},
				{
					128f,
					new FontSizeInfo(218f, 93f)
				},
				{
					256f,
					new FontSizeInfo(434f, 186f)
				}
			}
		},
		{
			"Corbel",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 13f)
				},
				{
					20f,
					new FontSizeInfo(35f, 14f)
				},
				{
					22f,
					new FontSizeInfo(38f, 15f)
				},
				{
					24f,
					new FontSizeInfo(42f, 17f)
				},
				{
					26f,
					new FontSizeInfo(45f, 18f)
				},
				{
					28f,
					new FontSizeInfo(48f, 19f)
				},
				{
					36f,
					new FontSizeInfo(62f, 25f)
				},
				{
					48f,
					new FontSizeInfo(82f, 34f)
				},
				{
					72f,
					new FontSizeInfo(123f, 50f)
				},
				{
					96f,
					new FontSizeInfo(163f, 67f)
				},
				{
					128f,
					new FontSizeInfo(218f, 90f)
				},
				{
					256f,
					new FontSizeInfo(434f, 179f)
				}
			}
		},
		{
			"Ebrima",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 11f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 15f)
				},
				{
					22f,
					new FontSizeInfo(44f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 17f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 26f)
				},
				{
					48f,
					new FontSizeInfo(94f, 35f)
				},
				{
					72f,
					new FontSizeInfo(138f, 52f)
				},
				{
					96f,
					new FontSizeInfo(184f, 69f)
				},
				{
					128f,
					new FontSizeInfo(245f, 92f)
				},
				{
					256f,
					new FontSizeInfo(490f, 184f)
				}
			}
		},
		{
			"Franklin Gothic Medium",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(21f, 9f)
				},
				{
					12f,
					new FontSizeInfo(22f, 9f)
				},
				{
					14f,
					new FontSizeInfo(26f, 11f)
				},
				{
					16f,
					new FontSizeInfo(28f, 12f)
				},
				{
					18f,
					new FontSizeInfo(32f, 14f)
				},
				{
					20f,
					new FontSizeInfo(36f, 16f)
				},
				{
					22f,
					new FontSizeInfo(39f, 17f)
				},
				{
					24f,
					new FontSizeInfo(40f, 19f)
				},
				{
					26f,
					new FontSizeInfo(44f, 21f)
				},
				{
					28f,
					new FontSizeInfo(46f, 22f)
				},
				{
					36f,
					new FontSizeInfo(64f, 28f)
				},
				{
					48f,
					new FontSizeInfo(85f, 38f)
				},
				{
					72f,
					new FontSizeInfo(126f, 56f)
				},
				{
					96f,
					new FontSizeInfo(168f, 75f)
				},
				{
					128f,
					new FontSizeInfo(224f, 100f)
				},
				{
					256f,
					new FontSizeInfo(416f, 200f)
				}
			}
		},
		{
			"Gabriola",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(25f, 4f)
				},
				{
					10f,
					new FontSizeInfo(26f, 5f)
				},
				{
					11f,
					new FontSizeInfo(32f, 6f)
				},
				{
					12f,
					new FontSizeInfo(33f, 6f)
				},
				{
					14f,
					new FontSizeInfo(40f, 8f)
				},
				{
					16f,
					new FontSizeInfo(44f, 8f)
				},
				{
					18f,
					new FontSizeInfo(51f, 10f)
				},
				{
					20f,
					new FontSizeInfo(56f, 11f)
				},
				{
					22f,
					new FontSizeInfo(61f, 12f)
				},
				{
					24f,
					new FontSizeInfo(66f, 13f)
				},
				{
					26f,
					new FontSizeInfo(73f, 14f)
				},
				{
					28f,
					new FontSizeInfo(76f, 15f)
				},
				{
					36f,
					new FontSizeInfo(98f, 19f)
				},
				{
					48f,
					new FontSizeInfo(131f, 26f)
				},
				{
					72f,
					new FontSizeInfo(195f, 38f)
				},
				{
					96f,
					new FontSizeInfo(262f, 51f)
				},
				{
					128f,
					new FontSizeInfo(349f, 69f)
				},
				{
					256f,
					new FontSizeInfo(693f, 137f)
				}
			}
		},
		{
			"Gadugi",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(24f, 10f)
				},
				{
					16f,
					new FontSizeInfo(27f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 13f)
				},
				{
					20f,
					new FontSizeInfo(34f, 15f)
				},
				{
					22f,
					new FontSizeInfo(37f, 16f)
				},
				{
					24f,
					new FontSizeInfo(40f, 17f)
				},
				{
					26f,
					new FontSizeInfo(44f, 19f)
				},
				{
					28f,
					new FontSizeInfo(47f, 20f)
				},
				{
					36f,
					new FontSizeInfo(60f, 26f)
				},
				{
					48f,
					new FontSizeInfo(81f, 35f)
				},
				{
					72f,
					new FontSizeInfo(120f, 52f)
				},
				{
					96f,
					new FontSizeInfo(159f, 69f)
				},
				{
					128f,
					new FontSizeInfo(213f, 92f)
				},
				{
					256f,
					new FontSizeInfo(482f, 184f)
				}
			}
		},
		{
			"Impact",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(21f, 8f)
				},
				{
					12f,
					new FontSizeInfo(22f, 9f)
				},
				{
					14f,
					new FontSizeInfo(24f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(30f, 13f)
				},
				{
					20f,
					new FontSizeInfo(36f, 15f)
				},
				{
					22f,
					new FontSizeInfo(38f, 16f)
				},
				{
					24f,
					new FontSizeInfo(40f, 17f)
				},
				{
					26f,
					new FontSizeInfo(45f, 19f)
				},
				{
					28f,
					new FontSizeInfo(46f, 20f)
				},
				{
					36f,
					new FontSizeInfo(63f, 26f)
				},
				{
					48f,
					new FontSizeInfo(83f, 35f)
				},
				{
					72f,
					new FontSizeInfo(119f, 52f)
				},
				{
					96f,
					new FontSizeInfo(158f, 69f)
				},
				{
					128f,
					new FontSizeInfo(212f, 93f)
				},
				{
					256f,
					new FontSizeInfo(420f, 185f)
				}
			}
		},
		{
			"Javanese Text",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(30f, 6f)
				},
				{
					10f,
					new FontSizeInfo(33f, 8f)
				},
				{
					11f,
					new FontSizeInfo(39f, 9f)
				},
				{
					12f,
					new FontSizeInfo(41f, 9f)
				},
				{
					14f,
					new FontSizeInfo(49f, 11f)
				},
				{
					16f,
					new FontSizeInfo(53f, 12f)
				},
				{
					18f,
					new FontSizeInfo(61f, 14f)
				},
				{
					20f,
					new FontSizeInfo(70f, 16f)
				},
				{
					22f,
					new FontSizeInfo(74f, 17f)
				},
				{
					24f,
					new FontSizeInfo(82f, 19f)
				},
				{
					26f,
					new FontSizeInfo(90f, 21f)
				},
				{
					28f,
					new FontSizeInfo(94f, 22f)
				},
				{
					36f,
					new FontSizeInfo(122f, 28f)
				},
				{
					48f,
					new FontSizeInfo(162f, 38f)
				},
				{
					72f,
					new FontSizeInfo(243f, 57f)
				},
				{
					96f,
					new FontSizeInfo(324f, 76f)
				},
				{
					128f,
					new FontSizeInfo(433f, 101f)
				},
				{
					256f,
					new FontSizeInfo(860f, 200f)
				}
			}
		},
		{
			"Leelawadee UI",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 11f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 15f)
				},
				{
					22f,
					new FontSizeInfo(44f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 17f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 26f)
				},
				{
					48f,
					new FontSizeInfo(92f, 35f)
				},
				{
					72f,
					new FontSizeInfo(136f, 52f)
				},
				{
					96f,
					new FontSizeInfo(180f, 69f)
				},
				{
					128f,
					new FontSizeInfo(241f, 92f)
				},
				{
					256f,
					new FontSizeInfo(482f, 184f)
				}
			}
		},
		{
			"Leelawadee UI Semilight",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 12f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 15f)
				},
				{
					22f,
					new FontSizeInfo(44f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 18f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 26f)
				},
				{
					48f,
					new FontSizeInfo(92f, 35f)
				},
				{
					72f,
					new FontSizeInfo(136f, 53f)
				},
				{
					96f,
					new FontSizeInfo(180f, 70f)
				},
				{
					128f,
					new FontSizeInfo(241f, 94f)
				},
				{
					256f,
					new FontSizeInfo(482f, 187f)
				}
			}
		},
		{
			"Lucida Console",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(26f, 13f)
				},
				{
					18f,
					new FontSizeInfo(30f, 14f)
				},
				{
					20f,
					new FontSizeInfo(34f, 16f)
				},
				{
					22f,
					new FontSizeInfo(36f, 17f)
				},
				{
					24f,
					new FontSizeInfo(40f, 19f)
				},
				{
					26f,
					new FontSizeInfo(43f, 21f)
				},
				{
					28f,
					new FontSizeInfo(46f, 22f)
				},
				{
					36f,
					new FontSizeInfo(59f, 29f)
				},
				{
					48f,
					new FontSizeInfo(79f, 39f)
				},
				{
					72f,
					new FontSizeInfo(117f, 58f)
				},
				{
					96f,
					new FontSizeInfo(156f, 77f)
				},
				{
					128f,
					new FontSizeInfo(208f, 103f)
				},
				{
					256f,
					new FontSizeInfo(414f, 205f)
				}
			}
		},
		{
			"Lucida Sans Unicode",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(22f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 12f)
				},
				{
					16f,
					new FontSizeInfo(26f, 13f)
				},
				{
					18f,
					new FontSizeInfo(30f, 15f)
				},
				{
					20f,
					new FontSizeInfo(36f, 17f)
				},
				{
					22f,
					new FontSizeInfo(36f, 18f)
				},
				{
					24f,
					new FontSizeInfo(40f, 20f)
				},
				{
					26f,
					new FontSizeInfo(43f, 22f)
				},
				{
					28f,
					new FontSizeInfo(46f, 23f)
				},
				{
					36f,
					new FontSizeInfo(61f, 30f)
				},
				{
					48f,
					new FontSizeInfo(79f, 40f)
				},
				{
					72f,
					new FontSizeInfo(117f, 61f)
				},
				{
					96f,
					new FontSizeInfo(158f, 81f)
				},
				{
					128f,
					new FontSizeInfo(210f, 108f)
				},
				{
					256f,
					new FontSizeInfo(558f, 216f)
				}
			}
		},
		{
			"Malgun Gothic",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(35f, 12f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(42f, 15f)
				},
				{
					22f,
					new FontSizeInfo(45f, 16f)
				},
				{
					24f,
					new FontSizeInfo(51f, 18f)
				},
				{
					26f,
					new FontSizeInfo(52f, 19f)
				},
				{
					28f,
					new FontSizeInfo(55f, 20f)
				},
				{
					36f,
					new FontSizeInfo(72f, 26f)
				},
				{
					48f,
					new FontSizeInfo(93f, 35f)
				},
				{
					72f,
					new FontSizeInfo(137f, 53f)
				},
				{
					96f,
					new FontSizeInfo(181f, 71f)
				},
				{
					128f,
					new FontSizeInfo(242f, 94f)
				},
				{
					256f,
					new FontSizeInfo(484f, 188f)
				}
			}
		},
		{
			"Malgun Gothic Semilight",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 11f)
				},
				{
					16f,
					new FontSizeInfo(35f, 12f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(42f, 15f)
				},
				{
					22f,
					new FontSizeInfo(45f, 16f)
				},
				{
					24f,
					new FontSizeInfo(51f, 18f)
				},
				{
					26f,
					new FontSizeInfo(52f, 20f)
				},
				{
					28f,
					new FontSizeInfo(55f, 21f)
				},
				{
					36f,
					new FontSizeInfo(72f, 27f)
				},
				{
					48f,
					new FontSizeInfo(93f, 36f)
				},
				{
					72f,
					new FontSizeInfo(137f, 54f)
				},
				{
					96f,
					new FontSizeInfo(181f, 72f)
				},
				{
					128f,
					new FontSizeInfo(242f, 96f)
				},
				{
					256f,
					new FontSizeInfo(484f, 190f)
				}
			}
		},
		{
			"Microsoft Himalaya",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(21f, 4f)
				},
				{
					11f,
					new FontSizeInfo(22f, 5f)
				},
				{
					12f,
					new FontSizeInfo(24f, 5f)
				},
				{
					14f,
					new FontSizeInfo(28f, 6f)
				},
				{
					16f,
					new FontSizeInfo(31f, 7f)
				},
				{
					18f,
					new FontSizeInfo(35f, 8f)
				},
				{
					20f,
					new FontSizeInfo(39f, 9f)
				},
				{
					22f,
					new FontSizeInfo(42f, 10f)
				},
				{
					24f,
					new FontSizeInfo(46f, 11f)
				},
				{
					26f,
					new FontSizeInfo(50f, 12f)
				},
				{
					28f,
					new FontSizeInfo(53f, 12f)
				},
				{
					36f,
					new FontSizeInfo(69f, 16f)
				},
				{
					48f,
					new FontSizeInfo(91f, 21f)
				},
				{
					72f,
					new FontSizeInfo(136f, 32f)
				},
				{
					96f,
					new FontSizeInfo(181f, 43f)
				},
				{
					128f,
					new FontSizeInfo(242f, 57f)
				},
				{
					256f,
					new FontSizeInfo(481f, 114f)
				}
			}
		},
		{
			"Microsoft JhengHei",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 11f)
				},
				{
					16f,
					new FontSizeInfo(28f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 14f)
				},
				{
					20f,
					new FontSizeInfo(35f, 16f)
				},
				{
					22f,
					new FontSizeInfo(37f, 17f)
				},
				{
					24f,
					new FontSizeInfo(41f, 19f)
				},
				{
					26f,
					new FontSizeInfo(45f, 20f)
				},
				{
					28f,
					new FontSizeInfo(48f, 21f)
				},
				{
					36f,
					new FontSizeInfo(62f, 28f)
				},
				{
					48f,
					new FontSizeInfo(82f, 37f)
				},
				{
					72f,
					new FontSizeInfo(122f, 56f)
				},
				{
					96f,
					new FontSizeInfo(162f, 74f)
				},
				{
					128f,
					new FontSizeInfo(217f, 99f)
				},
				{
					256f,
					new FontSizeInfo(481f, 198f)
				}
			}
		},
		{
			"Microsoft JhengHei UI",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 11f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 14f)
				},
				{
					20f,
					new FontSizeInfo(37f, 16f)
				},
				{
					22f,
					new FontSizeInfo(38f, 17f)
				},
				{
					24f,
					new FontSizeInfo(42f, 19f)
				},
				{
					26f,
					new FontSizeInfo(45f, 20f)
				},
				{
					28f,
					new FontSizeInfo(48f, 21f)
				},
				{
					36f,
					new FontSizeInfo(62f, 28f)
				},
				{
					48f,
					new FontSizeInfo(82f, 37f)
				},
				{
					72f,
					new FontSizeInfo(123f, 56f)
				},
				{
					96f,
					new FontSizeInfo(164f, 74f)
				},
				{
					128f,
					new FontSizeInfo(218f, 99f)
				},
				{
					256f,
					new FontSizeInfo(439f, 198f)
				}
			}
		},
		{
			"Microsoft JhengHei Light",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 11f)
				},
				{
					16f,
					new FontSizeInfo(28f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 14f)
				},
				{
					20f,
					new FontSizeInfo(35f, 16f)
				},
				{
					22f,
					new FontSizeInfo(37f, 17f)
				},
				{
					24f,
					new FontSizeInfo(41f, 18f)
				},
				{
					26f,
					new FontSizeInfo(45f, 20f)
				},
				{
					28f,
					new FontSizeInfo(48f, 21f)
				},
				{
					36f,
					new FontSizeInfo(62f, 28f)
				},
				{
					48f,
					new FontSizeInfo(82f, 37f)
				},
				{
					72f,
					new FontSizeInfo(122f, 55f)
				},
				{
					96f,
					new FontSizeInfo(162f, 74f)
				},
				{
					128f,
					new FontSizeInfo(217f, 98f)
				},
				{
					256f,
					new FontSizeInfo(481f, 196f)
				}
			}
		},
		{
			"Microsoft JhengHei UI Light",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 11f)
				},
				{
					16f,
					new FontSizeInfo(28f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 14f)
				},
				{
					20f,
					new FontSizeInfo(35f, 16f)
				},
				{
					22f,
					new FontSizeInfo(37f, 17f)
				},
				{
					24f,
					new FontSizeInfo(41f, 18f)
				},
				{
					26f,
					new FontSizeInfo(45f, 20f)
				},
				{
					28f,
					new FontSizeInfo(48f, 21f)
				},
				{
					36f,
					new FontSizeInfo(62f, 28f)
				},
				{
					48f,
					new FontSizeInfo(82f, 37f)
				},
				{
					72f,
					new FontSizeInfo(122f, 55f)
				},
				{
					96f,
					new FontSizeInfo(162f, 74f)
				},
				{
					128f,
					new FontSizeInfo(217f, 98f)
				},
				{
					256f,
					new FontSizeInfo(439f, 196f)
				}
			}
		},
		{
			"Microsoft New Tai Lue",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(21f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(30f, 11f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(38f, 15f)
				},
				{
					22f,
					new FontSizeInfo(41f, 16f)
				},
				{
					24f,
					new FontSizeInfo(45f, 17f)
				},
				{
					26f,
					new FontSizeInfo(49f, 19f)
				},
				{
					28f,
					new FontSizeInfo(52f, 20f)
				},
				{
					36f,
					new FontSizeInfo(67f, 26f)
				},
				{
					48f,
					new FontSizeInfo(89f, 35f)
				},
				{
					72f,
					new FontSizeInfo(134f, 52f)
				},
				{
					96f,
					new FontSizeInfo(178f, 69f)
				},
				{
					128f,
					new FontSizeInfo(237f, 92f)
				},
				{
					256f,
					new FontSizeInfo(472f, 184f)
				}
			}
		},
		{
			"Microsoft PhagsPa",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(26f, 10f)
				},
				{
					16f,
					new FontSizeInfo(29f, 11f)
				},
				{
					18f,
					new FontSizeInfo(33f, 13f)
				},
				{
					20f,
					new FontSizeInfo(36f, 15f)
				},
				{
					22f,
					new FontSizeInfo(39f, 16f)
				},
				{
					24f,
					new FontSizeInfo(43f, 17f)
				},
				{
					26f,
					new FontSizeInfo(48f, 19f)
				},
				{
					28f,
					new FontSizeInfo(51f, 20f)
				},
				{
					36f,
					new FontSizeInfo(64f, 26f)
				},
				{
					48f,
					new FontSizeInfo(86f, 35f)
				},
				{
					72f,
					new FontSizeInfo(128f, 52f)
				},
				{
					96f,
					new FontSizeInfo(171f, 69f)
				},
				{
					128f,
					new FontSizeInfo(228f, 92f)
				},
				{
					256f,
					new FontSizeInfo(452f, 184f)
				}
			}
		},
		{
			"Microsoft Sans Serif",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 13f)
				},
				{
					20f,
					new FontSizeInfo(33f, 15f)
				},
				{
					22f,
					new FontSizeInfo(36f, 16f)
				},
				{
					24f,
					new FontSizeInfo(41f, 18f)
				},
				{
					26f,
					new FontSizeInfo(43f, 19f)
				},
				{
					28f,
					new FontSizeInfo(45f, 21f)
				},
				{
					36f,
					new FontSizeInfo(60f, 27f)
				},
				{
					48f,
					new FontSizeInfo(79f, 36f)
				},
				{
					72f,
					new FontSizeInfo(117f, 53f)
				},
				{
					96f,
					new FontSizeInfo(157f, 71f)
				},
				{
					128f,
					new FontSizeInfo(208f, 95f)
				},
				{
					256f,
					new FontSizeInfo(414f, 190f)
				}
			}
		},
		{
			"Microsoft Tai Le",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(21f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(26f, 10f)
				},
				{
					16f,
					new FontSizeInfo(29f, 11f)
				},
				{
					18f,
					new FontSizeInfo(33f, 13f)
				},
				{
					20f,
					new FontSizeInfo(37f, 15f)
				},
				{
					22f,
					new FontSizeInfo(40f, 16f)
				},
				{
					24f,
					new FontSizeInfo(44f, 17f)
				},
				{
					26f,
					new FontSizeInfo(48f, 19f)
				},
				{
					28f,
					new FontSizeInfo(51f, 20f)
				},
				{
					36f,
					new FontSizeInfo(66f, 26f)
				},
				{
					48f,
					new FontSizeInfo(87f, 35f)
				},
				{
					72f,
					new FontSizeInfo(130f, 52f)
				},
				{
					96f,
					new FontSizeInfo(173f, 69f)
				},
				{
					128f,
					new FontSizeInfo(231f, 92f)
				},
				{
					256f,
					new FontSizeInfo(459f, 184f)
				}
			}
		},
		{
			"Microsoft YaHei",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 6f)
				},
				{
					10f,
					new FontSizeInfo(22f, 8f)
				},
				{
					11f,
					new FontSizeInfo(22f, 9f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 11f)
				},
				{
					16f,
					new FontSizeInfo(30f, 12f)
				},
				{
					18f,
					new FontSizeInfo(33f, 14f)
				},
				{
					20f,
					new FontSizeInfo(37f, 16f)
				},
				{
					22f,
					new FontSizeInfo(40f, 17f)
				},
				{
					24f,
					new FontSizeInfo(43f, 19f)
				},
				{
					26f,
					new FontSizeInfo(49f, 21f)
				},
				{
					28f,
					new FontSizeInfo(51f, 22f)
				},
				{
					36f,
					new FontSizeInfo(65f, 28f)
				},
				{
					48f,
					new FontSizeInfo(86f, 38f)
				},
				{
					72f,
					new FontSizeInfo(128f, 56f)
				},
				{
					96f,
					new FontSizeInfo(170f, 75f)
				},
				{
					128f,
					new FontSizeInfo(228f, 100f)
				},
				{
					256f,
					new FontSizeInfo(471f, 200f)
				}
			}
		},
		{
			"Microsoft YaHei UI",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 6f)
				},
				{
					10f,
					new FontSizeInfo(22f, 8f)
				},
				{
					11f,
					new FontSizeInfo(22f, 9f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 11f)
				},
				{
					16f,
					new FontSizeInfo(30f, 12f)
				},
				{
					18f,
					new FontSizeInfo(33f, 14f)
				},
				{
					20f,
					new FontSizeInfo(37f, 16f)
				},
				{
					22f,
					new FontSizeInfo(40f, 17f)
				},
				{
					24f,
					new FontSizeInfo(43f, 19f)
				},
				{
					26f,
					new FontSizeInfo(49f, 21f)
				},
				{
					28f,
					new FontSizeInfo(51f, 22f)
				},
				{
					36f,
					new FontSizeInfo(65f, 28f)
				},
				{
					48f,
					new FontSizeInfo(86f, 38f)
				},
				{
					72f,
					new FontSizeInfo(128f, 56f)
				},
				{
					96f,
					new FontSizeInfo(170f, 75f)
				},
				{
					128f,
					new FontSizeInfo(228f, 100f)
				},
				{
					256f,
					new FontSizeInfo(439f, 200f)
				}
			}
		},
		{
			"Microsoft YaHei Light",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 6f)
				},
				{
					10f,
					new FontSizeInfo(22f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 9f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 11f)
				},
				{
					16f,
					new FontSizeInfo(30f, 12f)
				},
				{
					18f,
					new FontSizeInfo(33f, 14f)
				},
				{
					20f,
					new FontSizeInfo(37f, 15f)
				},
				{
					22f,
					new FontSizeInfo(42f, 17f)
				},
				{
					24f,
					new FontSizeInfo(45f, 18f)
				},
				{
					26f,
					new FontSizeInfo(51f, 20f)
				},
				{
					28f,
					new FontSizeInfo(53f, 21f)
				},
				{
					36f,
					new FontSizeInfo(67f, 28f)
				},
				{
					48f,
					new FontSizeInfo(88f, 37f)
				},
				{
					72f,
					new FontSizeInfo(132f, 55f)
				},
				{
					96f,
					new FontSizeInfo(176f, 73f)
				},
				{
					128f,
					new FontSizeInfo(236f, 98f)
				},
				{
					256f,
					new FontSizeInfo(459f, 195f)
				}
			}
		},
		{
			"Microsoft YaHei UI Light",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 6f)
				},
				{
					10f,
					new FontSizeInfo(22f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 9f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 11f)
				},
				{
					16f,
					new FontSizeInfo(30f, 12f)
				},
				{
					18f,
					new FontSizeInfo(33f, 14f)
				},
				{
					20f,
					new FontSizeInfo(37f, 15f)
				},
				{
					22f,
					new FontSizeInfo(40f, 17f)
				},
				{
					24f,
					new FontSizeInfo(43f, 18f)
				},
				{
					26f,
					new FontSizeInfo(49f, 20f)
				},
				{
					28f,
					new FontSizeInfo(51f, 21f)
				},
				{
					36f,
					new FontSizeInfo(65f, 28f)
				},
				{
					48f,
					new FontSizeInfo(86f, 37f)
				},
				{
					72f,
					new FontSizeInfo(128f, 55f)
				},
				{
					96f,
					new FontSizeInfo(170f, 73f)
				},
				{
					128f,
					new FontSizeInfo(228f, 98f)
				},
				{
					256f,
					new FontSizeInfo(469f, 195f)
				}
			}
		},
		{
			"Microsoft Yi Baiti",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(24f, 10f)
				},
				{
					16f,
					new FontSizeInfo(26f, 11f)
				},
				{
					18f,
					new FontSizeInfo(29f, 12f)
				},
				{
					20f,
					new FontSizeInfo(32f, 14f)
				},
				{
					22f,
					new FontSizeInfo(34f, 15f)
				},
				{
					24f,
					new FontSizeInfo(38f, 16f)
				},
				{
					26f,
					new FontSizeInfo(41f, 18f)
				},
				{
					28f,
					new FontSizeInfo(43f, 19f)
				},
				{
					36f,
					new FontSizeInfo(56f, 25f)
				},
				{
					48f,
					new FontSizeInfo(74f, 33f)
				},
				{
					72f,
					new FontSizeInfo(0f, 49f)
				},
				{
					96f,
					new FontSizeInfo(149f, 66f)
				},
				{
					128f,
					new FontSizeInfo(198f, 88f)
				},
				{
					256f,
					new FontSizeInfo(398f, 175f)
				}
			}
		},
		{
			"MingLiU-ExtB",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(21f, 8f)
				},
				{
					12f,
					new FontSizeInfo(22f, 8f)
				},
				{
					14f,
					new FontSizeInfo(26f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(34f, 12f)
				},
				{
					20f,
					new FontSizeInfo(37f, 14f)
				},
				{
					22f,
					new FontSizeInfo(40f, 15f)
				},
				{
					24f,
					new FontSizeInfo(43f, 16f)
				},
				{
					26f,
					new FontSizeInfo(49f, 18f)
				},
				{
					28f,
					new FontSizeInfo(51f, 19f)
				},
				{
					36f,
					new FontSizeInfo(67f, 24f)
				},
				{
					48f,
					new FontSizeInfo(90f, 32f)
				},
				{
					72f,
					new FontSizeInfo(0f, 48f)
				},
				{
					96f,
					new FontSizeInfo(179f, 64f)
				},
				{
					128f,
					new FontSizeInfo(238f, 86f)
				},
				{
					256f,
					new FontSizeInfo(476f, 171f)
				}
			}
		},
		{
			"PMingLiU-ExtB",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(21f, 7f)
				},
				{
					12f,
					new FontSizeInfo(22f, 8f)
				},
				{
					14f,
					new FontSizeInfo(26f, 9f)
				},
				{
					16f,
					new FontSizeInfo(28f, 10f)
				},
				{
					18f,
					new FontSizeInfo(34f, 11f)
				},
				{
					20f,
					new FontSizeInfo(37f, 13f)
				},
				{
					22f,
					new FontSizeInfo(40f, 14f)
				},
				{
					24f,
					new FontSizeInfo(43f, 15f)
				},
				{
					26f,
					new FontSizeInfo(49f, 16f)
				},
				{
					28f,
					new FontSizeInfo(51f, 17f)
				},
				{
					36f,
					new FontSizeInfo(67f, 23f)
				},
				{
					48f,
					new FontSizeInfo(90f, 30f)
				},
				{
					72f,
					new FontSizeInfo(0f, 45f)
				},
				{
					96f,
					new FontSizeInfo(179f, 60f)
				},
				{
					128f,
					new FontSizeInfo(238f, 80f)
				},
				{
					256f,
					new FontSizeInfo(476f, 160f)
				}
			}
		},
		{
			"MingLiU_HKSCS-ExtB",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(21f, 8f)
				},
				{
					12f,
					new FontSizeInfo(22f, 8f)
				},
				{
					14f,
					new FontSizeInfo(26f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(34f, 12f)
				},
				{
					20f,
					new FontSizeInfo(37f, 14f)
				},
				{
					22f,
					new FontSizeInfo(40f, 15f)
				},
				{
					24f,
					new FontSizeInfo(43f, 16f)
				},
				{
					26f,
					new FontSizeInfo(49f, 18f)
				},
				{
					28f,
					new FontSizeInfo(51f, 19f)
				},
				{
					36f,
					new FontSizeInfo(67f, 24f)
				},
				{
					48f,
					new FontSizeInfo(90f, 32f)
				},
				{
					72f,
					new FontSizeInfo(0f, 48f)
				},
				{
					96f,
					new FontSizeInfo(179f, 64f)
				},
				{
					128f,
					new FontSizeInfo(238f, 86f)
				},
				{
					256f,
					new FontSizeInfo(476f, 171f)
				}
			}
		},
		{
			"Mongolian Baiti",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(24f, 10f)
				},
				{
					16f,
					new FontSizeInfo(27f, 11f)
				},
				{
					18f,
					new FontSizeInfo(30f, 12f)
				},
				{
					20f,
					new FontSizeInfo(34f, 14f)
				},
				{
					22f,
					new FontSizeInfo(38f, 15f)
				},
				{
					24f,
					new FontSizeInfo(42f, 16f)
				},
				{
					26f,
					new FontSizeInfo(46f, 18f)
				},
				{
					28f,
					new FontSizeInfo(48f, 19f)
				},
				{
					36f,
					new FontSizeInfo(61f, 24f)
				},
				{
					48f,
					new FontSizeInfo(83f, 32f)
				},
				{
					72f,
					new FontSizeInfo(0f, 48f)
				},
				{
					96f,
					new FontSizeInfo(167f, 64f)
				},
				{
					128f,
					new FontSizeInfo(223f, 86f)
				},
				{
					256f,
					new FontSizeInfo(447f, 171f)
				}
			}
		},
		{
			"MV Boli",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 8f)
				},
				{
					10f,
					new FontSizeInfo(20f, 9f)
				},
				{
					11f,
					new FontSizeInfo(22f, 11f)
				},
				{
					12f,
					new FontSizeInfo(23f, 11f)
				},
				{
					14f,
					new FontSizeInfo(27f, 13f)
				},
				{
					16f,
					new FontSizeInfo(30f, 15f)
				},
				{
					18f,
					new FontSizeInfo(33f, 17f)
				},
				{
					20f,
					new FontSizeInfo(35f, 19f)
				},
				{
					22f,
					new FontSizeInfo(42f, 20f)
				},
				{
					24f,
					new FontSizeInfo(43f, 22f)
				},
				{
					26f,
					new FontSizeInfo(49f, 25f)
				},
				{
					28f,
					new FontSizeInfo(51f, 26f)
				},
				{
					36f,
					new FontSizeInfo(65f, 34f)
				},
				{
					48f,
					new FontSizeInfo(89f, 45f)
				},
				{
					72f,
					new FontSizeInfo(0f, 67f)
				},
				{
					96f,
					new FontSizeInfo(171f, 90f)
				},
				{
					128f,
					new FontSizeInfo(231f, 120f)
				},
				{
					256f,
					new FontSizeInfo(597f, 239f)
				}
			}
		},
		{
			"Myanmar Text",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(25f, 6f)
				},
				{
					10f,
					new FontSizeInfo(26f, 7f)
				},
				{
					11f,
					new FontSizeInfo(29f, 8f)
				},
				{
					12f,
					new FontSizeInfo(31f, 9f)
				},
				{
					14f,
					new FontSizeInfo(36f, 10f)
				},
				{
					16f,
					new FontSizeInfo(39f, 11f)
				},
				{
					18f,
					new FontSizeInfo(45f, 13f)
				},
				{
					20f,
					new FontSizeInfo(50f, 15f)
				},
				{
					22f,
					new FontSizeInfo(53f, 16f)
				},
				{
					24f,
					new FontSizeInfo(58f, 17f)
				},
				{
					26f,
					new FontSizeInfo(64f, 19f)
				},
				{
					28f,
					new FontSizeInfo(67f, 20f)
				},
				{
					36f,
					new FontSizeInfo(88f, 26f)
				},
				{
					48f,
					new FontSizeInfo(116f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 52f)
				},
				{
					96f,
					new FontSizeInfo(233f, 69f)
				},
				{
					128f,
					new FontSizeInfo(309f, 92f)
				},
				{
					256f,
					new FontSizeInfo(647f, 184f)
				}
			}
		},
		{
			"Nirmala UI",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 11f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 15f)
				},
				{
					22f,
					new FontSizeInfo(44f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 17f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 26f)
				},
				{
					48f,
					new FontSizeInfo(92f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 52f)
				},
				{
					96f,
					new FontSizeInfo(180f, 69f)
				},
				{
					128f,
					new FontSizeInfo(241f, 92f)
				},
				{
					256f,
					new FontSizeInfo(482f, 184f)
				}
			}
		},
		{
			"Nirmala UI Semilight",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 12f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 15f)
				},
				{
					22f,
					new FontSizeInfo(44f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 18f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 26f)
				},
				{
					48f,
					new FontSizeInfo(92f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 53f)
				},
				{
					96f,
					new FontSizeInfo(180f, 70f)
				},
				{
					128f,
					new FontSizeInfo(241f, 94f)
				},
				{
					256f,
					new FontSizeInfo(482f, 187f)
				}
			}
		},
		{
			"Palatino Linotype",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 6f)
				},
				{
					10f,
					new FontSizeInfo(21f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(24f, 8f)
				},
				{
					14f,
					new FontSizeInfo(28f, 10f)
				},
				{
					16f,
					new FontSizeInfo(30f, 11f)
				},
				{
					18f,
					new FontSizeInfo(34f, 12f)
				},
				{
					20f,
					new FontSizeInfo(38f, 14f)
				},
				{
					22f,
					new FontSizeInfo(40f, 15f)
				},
				{
					24f,
					new FontSizeInfo(47f, 16f)
				},
				{
					26f,
					new FontSizeInfo(50f, 18f)
				},
				{
					28f,
					new FontSizeInfo(53f, 19f)
				},
				{
					36f,
					new FontSizeInfo(67f, 24f)
				},
				{
					48f,
					new FontSizeInfo(90f, 32f)
				},
				{
					72f,
					new FontSizeInfo(0f, 48f)
				},
				{
					96f,
					new FontSizeInfo(178f, 64f)
				},
				{
					128f,
					new FontSizeInfo(240f, 86f)
				},
				{
					256f,
					new FontSizeInfo(478f, 171f)
				}
			}
		},
		{
			"Segoe MDL2 Assets",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(22f, 10f)
				},
				{
					12f,
					new FontSizeInfo(23f, 10f)
				},
				{
					14f,
					new FontSizeInfo(26f, 12f)
				},
				{
					16f,
					new FontSizeInfo(28f, 14f)
				},
				{
					18f,
					new FontSizeInfo(31f, 15f)
				},
				{
					20f,
					new FontSizeInfo(34f, 17f)
				},
				{
					22f,
					new FontSizeInfo(36f, 19f)
				},
				{
					24f,
					new FontSizeInfo(39f, 21f)
				},
				{
					26f,
					new FontSizeInfo(42f, 23f)
				},
				{
					28f,
					new FontSizeInfo(44f, 24f)
				},
				{
					36f,
					new FontSizeInfo(55f, 31f)
				},
				{
					48f,
					new FontSizeInfo(71f, 41f)
				},
				{
					72f,
					new FontSizeInfo(0f, 62f)
				},
				{
					96f,
					new FontSizeInfo(135f, 83f)
				},
				{
					128f,
					new FontSizeInfo(178f, 110f)
				},
				{
					256f,
					new FontSizeInfo(348f, 219f)
				}
			}
		},
		{
			"Segoe Print",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(24f, 8f)
				},
				{
					10f,
					new FontSizeInfo(28f, 10f)
				},
				{
					11f,
					new FontSizeInfo(31f, 11f)
				},
				{
					12f,
					new FontSizeInfo(33f, 12f)
				},
				{
					14f,
					new FontSizeInfo(39f, 14f)
				},
				{
					16f,
					new FontSizeInfo(42f, 15f)
				},
				{
					18f,
					new FontSizeInfo(49f, 18f)
				},
				{
					20f,
					new FontSizeInfo(55f, 20f)
				},
				{
					22f,
					new FontSizeInfo(58f, 21f)
				},
				{
					24f,
					new FontSizeInfo(65f, 23f)
				},
				{
					26f,
					new FontSizeInfo(71f, 26f)
				},
				{
					28f,
					new FontSizeInfo(74f, 27f)
				},
				{
					36f,
					new FontSizeInfo(97f, 35f)
				},
				{
					48f,
					new FontSizeInfo(129f, 47f)
				},
				{
					72f,
					new FontSizeInfo(0f, 70f)
				},
				{
					96f,
					new FontSizeInfo(259f, 94f)
				},
				{
					128f,
					new FontSizeInfo(347f, 125f)
				},
				{
					256f,
					new FontSizeInfo(687f, 248f)
				}
			}
		},
		{
			"Segoe Script",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(22f, 8f)
				},
				{
					10f,
					new FontSizeInfo(23f, 10f)
				},
				{
					11f,
					new FontSizeInfo(25f, 11f)
				},
				{
					12f,
					new FontSizeInfo(27f, 12f)
				},
				{
					14f,
					new FontSizeInfo(33f, 14f)
				},
				{
					16f,
					new FontSizeInfo(36f, 15f)
				},
				{
					18f,
					new FontSizeInfo(41f, 18f)
				},
				{
					20f,
					new FontSizeInfo(45f, 20f)
				},
				{
					22f,
					new FontSizeInfo(50f, 21f)
				},
				{
					24f,
					new FontSizeInfo(55f, 23f)
				},
				{
					26f,
					new FontSizeInfo(59f, 26f)
				},
				{
					28f,
					new FontSizeInfo(62f, 27f)
				},
				{
					36f,
					new FontSizeInfo(81f, 35f)
				},
				{
					48f,
					new FontSizeInfo(109f, 47f)
				},
				{
					72f,
					new FontSizeInfo(0f, 70f)
				},
				{
					96f,
					new FontSizeInfo(214f, 94f)
				},
				{
					128f,
					new FontSizeInfo(287f, 125f)
				},
				{
					256f,
					new FontSizeInfo(571f, 248f)
				}
			}
		},
		{
			"Segoe UI",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 11f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 15f)
				},
				{
					22f,
					new FontSizeInfo(44f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 17f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 26f)
				},
				{
					48f,
					new FontSizeInfo(92f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 52f)
				},
				{
					96f,
					new FontSizeInfo(180f, 69f)
				},
				{
					128f,
					new FontSizeInfo(241f, 92f)
				},
				{
					256f,
					new FontSizeInfo(482f, 184f)
				}
			}
		},
		{
			"Segoe UI Black",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(22f, 9f)
				},
				{
					12f,
					new FontSizeInfo(23f, 10f)
				},
				{
					14f,
					new FontSizeInfo(27f, 12f)
				},
				{
					16f,
					new FontSizeInfo(34f, 13f)
				},
				{
					18f,
					new FontSizeInfo(35f, 15f)
				},
				{
					20f,
					new FontSizeInfo(41f, 17f)
				},
				{
					22f,
					new FontSizeInfo(44f, 18f)
				},
				{
					24f,
					new FontSizeInfo(50f, 20f)
				},
				{
					26f,
					new FontSizeInfo(51f, 22f)
				},
				{
					28f,
					new FontSizeInfo(54f, 23f)
				},
				{
					36f,
					new FontSizeInfo(70f, 30f)
				},
				{
					48f,
					new FontSizeInfo(92f, 40f)
				},
				{
					72f,
					new FontSizeInfo(0f, 60f)
				},
				{
					96f,
					new FontSizeInfo(182f, 79f)
				},
				{
					128f,
					new FontSizeInfo(245f, 106f)
				},
				{
					256f,
					new FontSizeInfo(488f, 212f)
				}
			}
		},
		{
			"Segoe UI Emoji",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 11f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 15f)
				},
				{
					22f,
					new FontSizeInfo(44f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 17f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 26f)
				},
				{
					48f,
					new FontSizeInfo(92f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 52f)
				},
				{
					96f,
					new FontSizeInfo(180f, 69f)
				},
				{
					128f,
					new FontSizeInfo(241f, 92f)
				},
				{
					256f,
					new FontSizeInfo(482f, 184f)
				}
			}
		},
		{
			"Segoe UI Historic",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 11f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 15f)
				},
				{
					22f,
					new FontSizeInfo(44f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 17f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 26f)
				},
				{
					48f,
					new FontSizeInfo(92f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 52f)
				},
				{
					96f,
					new FontSizeInfo(180f, 69f)
				},
				{
					128f,
					new FontSizeInfo(241f, 92f)
				},
				{
					256f,
					new FontSizeInfo(482f, 184f)
				}
			}
		},
		{
			"Segoe UI Light",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 8f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 11f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 14f)
				},
				{
					22f,
					new FontSizeInfo(44f, 15f)
				},
				{
					24f,
					new FontSizeInfo(50f, 17f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 25f)
				},
				{
					48f,
					new FontSizeInfo(92f, 34f)
				},
				{
					72f,
					new FontSizeInfo(0f, 51f)
				},
				{
					96f,
					new FontSizeInfo(180f, 68f)
				},
				{
					128f,
					new FontSizeInfo(241f, 91f)
				},
				{
					256f,
					new FontSizeInfo(482f, 181f)
				}
			}
		},
		{
			"Segoe UI Semibold",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 9f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 11f)
				},
				{
					16f,
					new FontSizeInfo(34f, 12f)
				},
				{
					18f,
					new FontSizeInfo(35f, 14f)
				},
				{
					20f,
					new FontSizeInfo(41f, 16f)
				},
				{
					22f,
					new FontSizeInfo(44f, 17f)
				},
				{
					24f,
					new FontSizeInfo(50f, 18f)
				},
				{
					26f,
					new FontSizeInfo(51f, 20f)
				},
				{
					28f,
					new FontSizeInfo(54f, 21f)
				},
				{
					36f,
					new FontSizeInfo(70f, 28f)
				},
				{
					48f,
					new FontSizeInfo(92f, 37f)
				},
				{
					72f,
					new FontSizeInfo(0f, 55f)
				},
				{
					96f,
					new FontSizeInfo(180f, 74f)
				},
				{
					128f,
					new FontSizeInfo(241f, 99f)
				},
				{
					256f,
					new FontSizeInfo(482f, 196f)
				}
			}
		},
		{
			"Segoe UI Semilight",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 12f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 15f)
				},
				{
					22f,
					new FontSizeInfo(44f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 18f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 26f)
				},
				{
					48f,
					new FontSizeInfo(92f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 53f)
				},
				{
					96f,
					new FontSizeInfo(180f, 70f)
				},
				{
					128f,
					new FontSizeInfo(241f, 94f)
				},
				{
					256f,
					new FontSizeInfo(482f, 187f)
				}
			}
		},
		{
			"Segoe UI Symbol",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 11f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 15f)
				},
				{
					22f,
					new FontSizeInfo(44f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 17f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 26f)
				},
				{
					48f,
					new FontSizeInfo(92f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 52f)
				},
				{
					96f,
					new FontSizeInfo(180f, 69f)
				},
				{
					128f,
					new FontSizeInfo(241f, 92f)
				},
				{
					256f,
					new FontSizeInfo(482f, 184f)
				}
			}
		},
		{
			"NSimSun",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(27f, 11f)
				},
				{
					18f,
					new FontSizeInfo(30f, 12f)
				},
				{
					20f,
					new FontSizeInfo(34f, 14f)
				},
				{
					22f,
					new FontSizeInfo(36f, 15f)
				},
				{
					24f,
					new FontSizeInfo(42f, 16f)
				},
				{
					26f,
					new FontSizeInfo(45f, 18f)
				},
				{
					28f,
					new FontSizeInfo(47f, 19f)
				},
				{
					36f,
					new FontSizeInfo(62f, 24f)
				},
				{
					48f,
					new FontSizeInfo(82f, 32f)
				},
				{
					72f,
					new FontSizeInfo(0f, 48f)
				},
				{
					96f,
					new FontSizeInfo(163f, 64f)
				},
				{
					128f,
					new FontSizeInfo(218f, 86f)
				},
				{
					256f,
					new FontSizeInfo(436f, 171f)
				}
			}
		},
		{
			"SimSun-ExtB",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(24f, 10f)
				},
				{
					16f,
					new FontSizeInfo(26f, 11f)
				},
				{
					18f,
					new FontSizeInfo(29f, 12f)
				},
				{
					20f,
					new FontSizeInfo(32f, 14f)
				},
				{
					22f,
					new FontSizeInfo(34f, 15f)
				},
				{
					24f,
					new FontSizeInfo(38f, 16f)
				},
				{
					26f,
					new FontSizeInfo(41f, 18f)
				},
				{
					28f,
					new FontSizeInfo(43f, 19f)
				},
				{
					36f,
					new FontSizeInfo(56f, 24f)
				},
				{
					48f,
					new FontSizeInfo(74f, 32f)
				},
				{
					72f,
					new FontSizeInfo(0f, 48f)
				},
				{
					96f,
					new FontSizeInfo(147f, 64f)
				},
				{
					128f,
					new FontSizeInfo(196f, 86f)
				},
				{
					256f,
					new FontSizeInfo(390f, 171f)
				}
			}
		},
		{
			"Sitka Small",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 8f)
				},
				{
					10f,
					new FontSizeInfo(22f, 9f)
				},
				{
					11f,
					new FontSizeInfo(27f, 10f)
				},
				{
					12f,
					new FontSizeInfo(28f, 11f)
				},
				{
					14f,
					new FontSizeInfo(32f, 13f)
				},
				{
					16f,
					new FontSizeInfo(36f, 15f)
				},
				{
					18f,
					new FontSizeInfo(40f, 17f)
				},
				{
					20f,
					new FontSizeInfo(46f, 19f)
				},
				{
					22f,
					new FontSizeInfo(49f, 20f)
				},
				{
					24f,
					new FontSizeInfo(55f, 22f)
				},
				{
					26f,
					new FontSizeInfo(59f, 24f)
				},
				{
					28f,
					new FontSizeInfo(61f, 26f)
				},
				{
					36f,
					new FontSizeInfo(80f, 33f)
				},
				{
					48f,
					new FontSizeInfo(106f, 44f)
				},
				{
					72f,
					new FontSizeInfo(0f, 66f)
				},
				{
					96f,
					new FontSizeInfo(212f, 88f)
				},
				{
					128f,
					new FontSizeInfo(284f, 118f)
				},
				{
					256f,
					new FontSizeInfo(564f, 235f)
				}
			}
		},
		{
			"Sitka Text",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 7f)
				},
				{
					10f,
					new FontSizeInfo(22f, 8f)
				},
				{
					11f,
					new FontSizeInfo(24f, 10f)
				},
				{
					12f,
					new FontSizeInfo(26f, 10f)
				},
				{
					14f,
					new FontSizeInfo(32f, 12f)
				},
				{
					16f,
					new FontSizeInfo(34f, 13f)
				},
				{
					18f,
					new FontSizeInfo(40f, 15f)
				},
				{
					20f,
					new FontSizeInfo(44f, 17f)
				},
				{
					22f,
					new FontSizeInfo(47f, 18f)
				},
				{
					24f,
					new FontSizeInfo(53f, 20f)
				},
				{
					26f,
					new FontSizeInfo(56f, 22f)
				},
				{
					28f,
					new FontSizeInfo(59f, 24f)
				},
				{
					36f,
					new FontSizeInfo(77f, 31f)
				},
				{
					48f,
					new FontSizeInfo(102f, 41f)
				},
				{
					72f,
					new FontSizeInfo(0f, 61f)
				},
				{
					96f,
					new FontSizeInfo(205f, 81f)
				},
				{
					128f,
					new FontSizeInfo(273f, 109f)
				},
				{
					256f,
					new FontSizeInfo(544f, 216f)
				}
			}
		},
		{
			"Sitka Subheading",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 7f)
				},
				{
					10f,
					new FontSizeInfo(22f, 8f)
				},
				{
					11f,
					new FontSizeInfo(24f, 9f)
				},
				{
					12f,
					new FontSizeInfo(26f, 9f)
				},
				{
					14f,
					new FontSizeInfo(32f, 11f)
				},
				{
					16f,
					new FontSizeInfo(34f, 12f)
				},
				{
					18f,
					new FontSizeInfo(40f, 14f)
				},
				{
					20f,
					new FontSizeInfo(44f, 16f)
				},
				{
					22f,
					new FontSizeInfo(47f, 17f)
				},
				{
					24f,
					new FontSizeInfo(53f, 19f)
				},
				{
					26f,
					new FontSizeInfo(56f, 21f)
				},
				{
					28f,
					new FontSizeInfo(59f, 22f)
				},
				{
					36f,
					new FontSizeInfo(77f, 28f)
				},
				{
					48f,
					new FontSizeInfo(102f, 38f)
				},
				{
					72f,
					new FontSizeInfo(0f, 57f)
				},
				{
					96f,
					new FontSizeInfo(205f, 76f)
				},
				{
					128f,
					new FontSizeInfo(273f, 101f)
				},
				{
					256f,
					new FontSizeInfo(544f, 201f)
				}
			}
		},
		{
			"Sitka Heading",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 6f)
				},
				{
					10f,
					new FontSizeInfo(22f, 7f)
				},
				{
					11f,
					new FontSizeInfo(24f, 8f)
				},
				{
					12f,
					new FontSizeInfo(26f, 9f)
				},
				{
					14f,
					new FontSizeInfo(32f, 11f)
				},
				{
					16f,
					new FontSizeInfo(34f, 12f)
				},
				{
					18f,
					new FontSizeInfo(40f, 13f)
				},
				{
					20f,
					new FontSizeInfo(44f, 15f)
				},
				{
					22f,
					new FontSizeInfo(47f, 16f)
				},
				{
					24f,
					new FontSizeInfo(53f, 18f)
				},
				{
					26f,
					new FontSizeInfo(56f, 20f)
				},
				{
					28f,
					new FontSizeInfo(59f, 21f)
				},
				{
					36f,
					new FontSizeInfo(77f, 27f)
				},
				{
					48f,
					new FontSizeInfo(102f, 36f)
				},
				{
					72f,
					new FontSizeInfo(0f, 54f)
				},
				{
					96f,
					new FontSizeInfo(205f, 71f)
				},
				{
					128f,
					new FontSizeInfo(273f, 95f)
				},
				{
					256f,
					new FontSizeInfo(544f, 190f)
				}
			}
		},
		{
			"Sitka Display",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 6f)
				},
				{
					10f,
					new FontSizeInfo(22f, 7f)
				},
				{
					11f,
					new FontSizeInfo(24f, 8f)
				},
				{
					12f,
					new FontSizeInfo(26f, 8f)
				},
				{
					14f,
					new FontSizeInfo(32f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 11f)
				},
				{
					18f,
					new FontSizeInfo(40f, 13f)
				},
				{
					20f,
					new FontSizeInfo(44f, 14f)
				},
				{
					22f,
					new FontSizeInfo(47f, 15f)
				},
				{
					24f,
					new FontSizeInfo(53f, 17f)
				},
				{
					26f,
					new FontSizeInfo(56f, 19f)
				},
				{
					28f,
					new FontSizeInfo(59f, 20f)
				},
				{
					36f,
					new FontSizeInfo(77f, 25f)
				},
				{
					48f,
					new FontSizeInfo(102f, 34f)
				},
				{
					72f,
					new FontSizeInfo(0f, 51f)
				},
				{
					96f,
					new FontSizeInfo(205f, 68f)
				},
				{
					128f,
					new FontSizeInfo(273f, 91f)
				},
				{
					256f,
					new FontSizeInfo(544f, 181f)
				}
			}
		},
		{
			"Sitka Banner",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 6f)
				},
				{
					10f,
					new FontSizeInfo(22f, 7f)
				},
				{
					11f,
					new FontSizeInfo(24f, 8f)
				},
				{
					12f,
					new FontSizeInfo(26f, 8f)
				},
				{
					14f,
					new FontSizeInfo(32f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 11f)
				},
				{
					18f,
					new FontSizeInfo(40f, 12f)
				},
				{
					20f,
					new FontSizeInfo(44f, 14f)
				},
				{
					22f,
					new FontSizeInfo(47f, 15f)
				},
				{
					24f,
					new FontSizeInfo(53f, 16f)
				},
				{
					26f,
					new FontSizeInfo(56f, 18f)
				},
				{
					28f,
					new FontSizeInfo(59f, 19f)
				},
				{
					36f,
					new FontSizeInfo(77f, 24f)
				},
				{
					48f,
					new FontSizeInfo(102f, 32f)
				},
				{
					72f,
					new FontSizeInfo(0f, 49f)
				},
				{
					96f,
					new FontSizeInfo(205f, 65f)
				},
				{
					128f,
					new FontSizeInfo(273f, 87f)
				},
				{
					256f,
					new FontSizeInfo(544f, 173f)
				}
			}
		},
		{
			"Tahoma",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(24f, 10f)
				},
				{
					16f,
					new FontSizeInfo(26f, 11f)
				},
				{
					18f,
					new FontSizeInfo(30f, 13f)
				},
				{
					20f,
					new FontSizeInfo(34f, 15f)
				},
				{
					22f,
					new FontSizeInfo(36f, 16f)
				},
				{
					24f,
					new FontSizeInfo(40f, 17f)
				},
				{
					26f,
					new FontSizeInfo(43f, 19f)
				},
				{
					28f,
					new FontSizeInfo(46f, 20f)
				},
				{
					36f,
					new FontSizeInfo(59f, 26f)
				},
				{
					48f,
					new FontSizeInfo(78f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 52f)
				},
				{
					96f,
					new FontSizeInfo(155f, 70f)
				},
				{
					128f,
					new FontSizeInfo(207f, 93f)
				},
				{
					256f,
					new FontSizeInfo(412f, 186f)
				}
			}
		},
		{
			"Trebuchet MS",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 6f)
				},
				{
					10f,
					new FontSizeInfo(21f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(24f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 13f)
				},
				{
					20f,
					new FontSizeInfo(37f, 14f)
				},
				{
					22f,
					new FontSizeInfo(38f, 15f)
				},
				{
					24f,
					new FontSizeInfo(41f, 17f)
				},
				{
					26f,
					new FontSizeInfo(45f, 18f)
				},
				{
					28f,
					new FontSizeInfo(48f, 19f)
				},
				{
					36f,
					new FontSizeInfo(62f, 25f)
				},
				{
					48f,
					new FontSizeInfo(82f, 34f)
				},
				{
					72f,
					new FontSizeInfo(0f, 50f)
				},
				{
					96f,
					new FontSizeInfo(164f, 67f)
				},
				{
					128f,
					new FontSizeInfo(219f, 90f)
				},
				{
					256f,
					new FontSizeInfo(418f, 179f)
				}
			}
		},
		{
			"Verdana",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 10f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 12f)
				},
				{
					16f,
					new FontSizeInfo(26f, 13f)
				},
				{
					18f,
					new FontSizeInfo(30f, 15f)
				},
				{
					20f,
					new FontSizeInfo(33f, 17f)
				},
				{
					22f,
					new FontSizeInfo(36f, 18f)
				},
				{
					24f,
					new FontSizeInfo(39f, 20f)
				},
				{
					26f,
					new FontSizeInfo(43f, 22f)
				},
				{
					28f,
					new FontSizeInfo(47f, 24f)
				},
				{
					36f,
					new FontSizeInfo(61f, 31f)
				},
				{
					48f,
					new FontSizeInfo(80f, 41f)
				},
				{
					72f,
					new FontSizeInfo(0f, 61f)
				},
				{
					96f,
					new FontSizeInfo(156f, 81f)
				},
				{
					128f,
					new FontSizeInfo(207f, 109f)
				},
				{
					256f,
					new FontSizeInfo(418f, 216f)
				}
			}
		},
		{
			"Webdings",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(24f, 8f)
				},
				{
					8f,
					new FontSizeInfo(22f, 11f)
				},
				{
					10f,
					new FontSizeInfo(21f, 13f)
				},
				{
					11f,
					new FontSizeInfo(21f, 15f)
				},
				{
					12f,
					new FontSizeInfo(21f, 16f)
				},
				{
					14f,
					new FontSizeInfo(26f, 19f)
				},
				{
					16f,
					new FontSizeInfo(27f, 21f)
				},
				{
					18f,
					new FontSizeInfo(31f, 24f)
				},
				{
					20f,
					new FontSizeInfo(35f, 27f)
				},
				{
					22f,
					new FontSizeInfo(36f, 29f)
				},
				{
					24f,
					new FontSizeInfo(40f, 32f)
				},
				{
					26f,
					new FontSizeInfo(43f, 35f)
				},
				{
					28f,
					new FontSizeInfo(46f, 37f)
				},
				{
					36f,
					new FontSizeInfo(59f, 48f)
				},
				{
					48f,
					new FontSizeInfo(78f, 64f)
				},
				{
					72f,
					new FontSizeInfo(0f, 96f)
				},
				{
					96f,
					new FontSizeInfo(155f, 128f)
				},
				{
					128f,
					new FontSizeInfo(206f, 171f)
				},
				{
					256f,
					new FontSizeInfo(410f, 341f)
				}
			}
		},
		{
			"Yu Gothic UI",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 11f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 15f)
				},
				{
					22f,
					new FontSizeInfo(44f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 17f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 26f)
				},
				{
					48f,
					new FontSizeInfo(92f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 52f)
				},
				{
					96f,
					new FontSizeInfo(180f, 69f)
				},
				{
					128f,
					new FontSizeInfo(241f, 92f)
				},
				{
					256f,
					new FontSizeInfo(482f, 184f)
				}
			}
		},
		{
			"Yu Gothic UI Semibold",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 9f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 11f)
				},
				{
					16f,
					new FontSizeInfo(34f, 12f)
				},
				{
					18f,
					new FontSizeInfo(35f, 14f)
				},
				{
					20f,
					new FontSizeInfo(41f, 16f)
				},
				{
					22f,
					new FontSizeInfo(44f, 17f)
				},
				{
					24f,
					new FontSizeInfo(50f, 18f)
				},
				{
					26f,
					new FontSizeInfo(51f, 20f)
				},
				{
					28f,
					new FontSizeInfo(54f, 21f)
				},
				{
					36f,
					new FontSizeInfo(70f, 28f)
				},
				{
					48f,
					new FontSizeInfo(92f, 37f)
				},
				{
					72f,
					new FontSizeInfo(0f, 55f)
				},
				{
					96f,
					new FontSizeInfo(180f, 74f)
				},
				{
					128f,
					new FontSizeInfo(241f, 99f)
				},
				{
					256f,
					new FontSizeInfo(482f, 196f)
				}
			}
		},
		{
			"Yu Gothic Light",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(22f, 7f)
				},
				{
					11f,
					new FontSizeInfo(24f, 8f)
				},
				{
					12f,
					new FontSizeInfo(26f, 9f)
				},
				{
					14f,
					new FontSizeInfo(32f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 11f)
				},
				{
					18f,
					new FontSizeInfo(40f, 13f)
				},
				{
					20f,
					new FontSizeInfo(44f, 14f)
				},
				{
					22f,
					new FontSizeInfo(47f, 16f)
				},
				{
					24f,
					new FontSizeInfo(53f, 17f)
				},
				{
					26f,
					new FontSizeInfo(56f, 19f)
				},
				{
					28f,
					new FontSizeInfo(61f, 20f)
				},
				{
					36f,
					new FontSizeInfo(77f, 26f)
				},
				{
					48f,
					new FontSizeInfo(102f, 34f)
				},
				{
					72f,
					new FontSizeInfo(0f, 51f)
				},
				{
					96f,
					new FontSizeInfo(205f, 69f)
				},
				{
					128f,
					new FontSizeInfo(275f, 92f)
				},
				{
					256f,
					new FontSizeInfo(550f, 183f)
				}
			}
		},
		{
			"Yu Gothic UI Light",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 11f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 14f)
				},
				{
					22f,
					new FontSizeInfo(44f, 15f)
				},
				{
					24f,
					new FontSizeInfo(50f, 17f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 25f)
				},
				{
					48f,
					new FontSizeInfo(92f, 34f)
				},
				{
					72f,
					new FontSizeInfo(0f, 51f)
				},
				{
					96f,
					new FontSizeInfo(180f, 68f)
				},
				{
					128f,
					new FontSizeInfo(241f, 91f)
				},
				{
					256f,
					new FontSizeInfo(482f, 181f)
				}
			}
		},
		{
			"Yu Gothic Medium",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(22f, 7f)
				},
				{
					11f,
					new FontSizeInfo(24f, 8f)
				},
				{
					12f,
					new FontSizeInfo(26f, 9f)
				},
				{
					14f,
					new FontSizeInfo(32f, 11f)
				},
				{
					16f,
					new FontSizeInfo(34f, 12f)
				},
				{
					18f,
					new FontSizeInfo(40f, 13f)
				},
				{
					20f,
					new FontSizeInfo(44f, 15f)
				},
				{
					22f,
					new FontSizeInfo(47f, 16f)
				},
				{
					24f,
					new FontSizeInfo(53f, 18f)
				},
				{
					26f,
					new FontSizeInfo(56f, 19f)
				},
				{
					28f,
					new FontSizeInfo(61f, 21f)
				},
				{
					36f,
					new FontSizeInfo(77f, 27f)
				},
				{
					48f,
					new FontSizeInfo(102f, 36f)
				},
				{
					72f,
					new FontSizeInfo(0f, 53f)
				},
				{
					96f,
					new FontSizeInfo(205f, 71f)
				},
				{
					128f,
					new FontSizeInfo(275f, 95f)
				},
				{
					256f,
					new FontSizeInfo(549f, 189f)
				}
			}
		},
		{
			"Yu Gothic UI Semilight",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 12f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 15f)
				},
				{
					22f,
					new FontSizeInfo(44f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 18f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 26f)
				},
				{
					48f,
					new FontSizeInfo(92f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 53f)
				},
				{
					96f,
					new FontSizeInfo(180f, 70f)
				},
				{
					128f,
					new FontSizeInfo(241f, 94f)
				},
				{
					256f,
					new FontSizeInfo(482f, 187f)
				}
			}
		},
		{
			"MT Extra",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 8f)
				},
				{
					8f,
					new FontSizeInfo(20f, 11f)
				},
				{
					10f,
					new FontSizeInfo(20f, 13f)
				},
				{
					11f,
					new FontSizeInfo(21f, 15f)
				},
				{
					12f,
					new FontSizeInfo(24f, 16f)
				},
				{
					14f,
					new FontSizeInfo(28f, 19f)
				},
				{
					16f,
					new FontSizeInfo(30f, 21f)
				},
				{
					18f,
					new FontSizeInfo(36f, 24f)
				},
				{
					20f,
					new FontSizeInfo(40f, 28f)
				},
				{
					22f,
					new FontSizeInfo(42f, 30f)
				},
				{
					24f,
					new FontSizeInfo(48f, 33f)
				},
				{
					26f,
					new FontSizeInfo(51f, 36f)
				},
				{
					28f,
					new FontSizeInfo(56f, 38f)
				},
				{
					36f,
					new FontSizeInfo(71f, 49f)
				},
				{
					48f,
					new FontSizeInfo(95f, 65f)
				},
				{
					72f,
					new FontSizeInfo(0f, 98f)
				},
				{
					96f,
					new FontSizeInfo(192f, 131f)
				},
				{
					128f,
					new FontSizeInfo(257f, 174f)
				},
				{
					256f,
					new FontSizeInfo(511f, 347f)
				}
			}
		},
		{
			"Agency FB",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 5f)
				},
				{
					11f,
					new FontSizeInfo(20f, 6f)
				},
				{
					12f,
					new FontSizeInfo(21f, 6f)
				},
				{
					14f,
					new FontSizeInfo(26f, 8f)
				},
				{
					16f,
					new FontSizeInfo(26f, 9f)
				},
				{
					18f,
					new FontSizeInfo(30f, 10f)
				},
				{
					20f,
					new FontSizeInfo(34f, 11f)
				},
				{
					22f,
					new FontSizeInfo(36f, 11f)
				},
				{
					24f,
					new FontSizeInfo(44f, 14f)
				},
				{
					26f,
					new FontSizeInfo(45f, 15f)
				},
				{
					28f,
					new FontSizeInfo(48f, 16f)
				},
				{
					36f,
					new FontSizeInfo(62f, 21f)
				},
				{
					48f,
					new FontSizeInfo(81f, 28f)
				},
				{
					72f,
					new FontSizeInfo(0f, 42f)
				},
				{
					96f,
					new FontSizeInfo(161f, 56f)
				},
				{
					128f,
					new FontSizeInfo(216f, 74f)
				},
				{
					256f,
					new FontSizeInfo(430f, 148f)
				}
			}
		},
		{
			"Algerian",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(21f, 9f)
				},
				{
					12f,
					new FontSizeInfo(23f, 10f)
				},
				{
					14f,
					new FontSizeInfo(26f, 11f)
				},
				{
					16f,
					new FontSizeInfo(29f, 13f)
				},
				{
					18f,
					new FontSizeInfo(34f, 14f)
				},
				{
					20f,
					new FontSizeInfo(38f, 16f)
				},
				{
					22f,
					new FontSizeInfo(40f, 17f)
				},
				{
					24f,
					new FontSizeInfo(46f, 19f)
				},
				{
					26f,
					new FontSizeInfo(50f, 21f)
				},
				{
					28f,
					new FontSizeInfo(52f, 22f)
				},
				{
					36f,
					new FontSizeInfo(68f, 29f)
				},
				{
					48f,
					new FontSizeInfo(91f, 38f)
				},
				{
					72f,
					new FontSizeInfo(0f, 58f)
				},
				{
					96f,
					new FontSizeInfo(184f, 77f)
				},
				{
					128f,
					new FontSizeInfo(244f, 103f)
				},
				{
					256f,
					new FontSizeInfo(488f, 205f)
				}
			}
		},
		{
			"Book Antiqua",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 12f)
				},
				{
					20f,
					new FontSizeInfo(35f, 14f)
				},
				{
					22f,
					new FontSizeInfo(38f, 15f)
				},
				{
					24f,
					new FontSizeInfo(42f, 16f)
				},
				{
					26f,
					new FontSizeInfo(45f, 18f)
				},
				{
					28f,
					new FontSizeInfo(48f, 19f)
				},
				{
					36f,
					new FontSizeInfo(62f, 24f)
				},
				{
					48f,
					new FontSizeInfo(83f, 32f)
				},
				{
					72f,
					new FontSizeInfo(0f, 48f)
				},
				{
					96f,
					new FontSizeInfo(164f, 64f)
				},
				{
					128f,
					new FontSizeInfo(220f, 86f)
				},
				{
					256f,
					new FontSizeInfo(438f, 171f)
				}
			}
		},
		{
			"Arial Narrow",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(22f, 7f)
				},
				{
					12f,
					new FontSizeInfo(21f, 7f)
				},
				{
					14f,
					new FontSizeInfo(24f, 9f)
				},
				{
					16f,
					new FontSizeInfo(27f, 10f)
				},
				{
					18f,
					new FontSizeInfo(31f, 11f)
				},
				{
					20f,
					new FontSizeInfo(34f, 12f)
				},
				{
					22f,
					new FontSizeInfo(36f, 13f)
				},
				{
					24f,
					new FontSizeInfo(40f, 15f)
				},
				{
					26f,
					new FontSizeInfo(45f, 16f)
				},
				{
					28f,
					new FontSizeInfo(47f, 17f)
				},
				{
					36f,
					new FontSizeInfo(61f, 22f)
				},
				{
					48f,
					new FontSizeInfo(80f, 29f)
				},
				{
					72f,
					new FontSizeInfo(0f, 44f)
				},
				{
					96f,
					new FontSizeInfo(157f, 58f)
				},
				{
					128f,
					new FontSizeInfo(210f, 78f)
				},
				{
					256f,
					new FontSizeInfo(418f, 156f)
				}
			}
		},
		{
			"Arial Rounded MT Bold",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(26f, 12f)
				},
				{
					18f,
					new FontSizeInfo(30f, 14f)
				},
				{
					20f,
					new FontSizeInfo(34f, 16f)
				},
				{
					22f,
					new FontSizeInfo(36f, 17f)
				},
				{
					24f,
					new FontSizeInfo(40f, 19f)
				},
				{
					26f,
					new FontSizeInfo(43f, 21f)
				},
				{
					28f,
					new FontSizeInfo(46f, 22f)
				},
				{
					36f,
					new FontSizeInfo(59f, 29f)
				},
				{
					48f,
					new FontSizeInfo(79f, 38f)
				},
				{
					72f,
					new FontSizeInfo(0f, 57f)
				},
				{
					96f,
					new FontSizeInfo(156f, 76f)
				},
				{
					128f,
					new FontSizeInfo(208f, 102f)
				},
				{
					256f,
					new FontSizeInfo(414f, 202f)
				}
			}
		},
		{
			"Baskerville Old Face",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(20f, 7f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 9f)
				},
				{
					16f,
					new FontSizeInfo(27f, 10f)
				},
				{
					18f,
					new FontSizeInfo(31f, 12f)
				},
				{
					20f,
					new FontSizeInfo(35f, 13f)
				},
				{
					22f,
					new FontSizeInfo(38f, 14f)
				},
				{
					24f,
					new FontSizeInfo(41f, 16f)
				},
				{
					26f,
					new FontSizeInfo(45f, 17f)
				},
				{
					28f,
					new FontSizeInfo(48f, 18f)
				},
				{
					36f,
					new FontSizeInfo(61f, 24f)
				},
				{
					48f,
					new FontSizeInfo(82f, 31f)
				},
				{
					72f,
					new FontSizeInfo(0f, 47f)
				},
				{
					96f,
					new FontSizeInfo(162f, 63f)
				},
				{
					128f,
					new FontSizeInfo(216f, 84f)
				},
				{
					256f,
					new FontSizeInfo(430f, 168f)
				}
			}
		},
		{
			"Bauhaus 93",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(23f, 9f)
				},
				{
					12f,
					new FontSizeInfo(25f, 9f)
				},
				{
					14f,
					new FontSizeInfo(28f, 11f)
				},
				{
					16f,
					new FontSizeInfo(33f, 12f)
				},
				{
					18f,
					new FontSizeInfo(37f, 14f)
				},
				{
					20f,
					new FontSizeInfo(42f, 15f)
				},
				{
					22f,
					new FontSizeInfo(45f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 18f)
				},
				{
					26f,
					new FontSizeInfo(54f, 20f)
				},
				{
					28f,
					new FontSizeInfo(57f, 21f)
				},
				{
					36f,
					new FontSizeInfo(74f, 27f)
				},
				{
					48f,
					new FontSizeInfo(100f, 36f)
				},
				{
					72f,
					new FontSizeInfo(0f, 55f)
				},
				{
					96f,
					new FontSizeInfo(199f, 73f)
				},
				{
					128f,
					new FontSizeInfo(266f, 97f)
				},
				{
					256f,
					new FontSizeInfo(531f, 194f)
				}
			}
		},
		{
			"Bell MT",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(21f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 12f)
				},
				{
					20f,
					new FontSizeInfo(35f, 14f)
				},
				{
					22f,
					new FontSizeInfo(38f, 15f)
				},
				{
					24f,
					new FontSizeInfo(42f, 16f)
				},
				{
					26f,
					new FontSizeInfo(45f, 18f)
				},
				{
					28f,
					new FontSizeInfo(48f, 19f)
				},
				{
					36f,
					new FontSizeInfo(62f, 24f)
				},
				{
					48f,
					new FontSizeInfo(82f, 32f)
				},
				{
					72f,
					new FontSizeInfo(0f, 48f)
				},
				{
					96f,
					new FontSizeInfo(163f, 64f)
				},
				{
					128f,
					new FontSizeInfo(218f, 86f)
				},
				{
					256f,
					new FontSizeInfo(433f, 171f)
				}
			}
		},
		{
			"Bernard MT Condensed",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(20f, 7f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(24f, 9f)
				},
				{
					16f,
					new FontSizeInfo(26f, 10f)
				},
				{
					18f,
					new FontSizeInfo(30f, 12f)
				},
				{
					20f,
					new FontSizeInfo(33f, 13f)
				},
				{
					22f,
					new FontSizeInfo(36f, 14f)
				},
				{
					24f,
					new FontSizeInfo(39f, 16f)
				},
				{
					26f,
					new FontSizeInfo(43f, 17f)
				},
				{
					28f,
					new FontSizeInfo(45f, 18f)
				},
				{
					36f,
					new FontSizeInfo(59f, 24f)
				},
				{
					48f,
					new FontSizeInfo(79f, 31f)
				},
				{
					72f,
					new FontSizeInfo(0f, 47f)
				},
				{
					96f,
					new FontSizeInfo(156f, 63f)
				},
				{
					128f,
					new FontSizeInfo(208f, 84f)
				},
				{
					256f,
					new FontSizeInfo(416f, 167f)
				}
			}
		},
		{
			"Bodoni MT",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(20f, 7f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(26f, 9f)
				},
				{
					16f,
					new FontSizeInfo(27f, 10f)
				},
				{
					18f,
					new FontSizeInfo(32f, 12f)
				},
				{
					20f,
					new FontSizeInfo(35f, 13f)
				},
				{
					22f,
					new FontSizeInfo(39f, 14f)
				},
				{
					24f,
					new FontSizeInfo(42f, 16f)
				},
				{
					26f,
					new FontSizeInfo(45f, 17f)
				},
				{
					28f,
					new FontSizeInfo(48f, 18f)
				},
				{
					36f,
					new FontSizeInfo(61f, 24f)
				},
				{
					48f,
					new FontSizeInfo(83f, 31f)
				},
				{
					72f,
					new FontSizeInfo(0f, 47f)
				},
				{
					96f,
					new FontSizeInfo(164f, 63f)
				},
				{
					128f,
					new FontSizeInfo(220f, 84f)
				},
				{
					256f,
					new FontSizeInfo(438f, 167f)
				}
			}
		},
		{
			"Bodoni MT Black",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 9f)
				},
				{
					11f,
					new FontSizeInfo(20f, 10f)
				},
				{
					12f,
					new FontSizeInfo(21f, 11f)
				},
				{
					14f,
					new FontSizeInfo(25f, 12f)
				},
				{
					16f,
					new FontSizeInfo(27f, 14f)
				},
				{
					18f,
					new FontSizeInfo(31f, 16f)
				},
				{
					20f,
					new FontSizeInfo(35f, 18f)
				},
				{
					22f,
					new FontSizeInfo(37f, 19f)
				},
				{
					24f,
					new FontSizeInfo(41f, 21f)
				},
				{
					26f,
					new FontSizeInfo(45f, 23f)
				},
				{
					28f,
					new FontSizeInfo(47f, 24f)
				},
				{
					36f,
					new FontSizeInfo(61f, 31f)
				},
				{
					48f,
					new FontSizeInfo(81f, 42f)
				},
				{
					72f,
					new FontSizeInfo(0f, 63f)
				},
				{
					96f,
					new FontSizeInfo(161f, 84f)
				},
				{
					128f,
					new FontSizeInfo(215f, 112f)
				},
				{
					256f,
					new FontSizeInfo(428f, 224f)
				}
			}
		},
		{
			"Bodoni MT Condensed",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(20f, 4f)
				},
				{
					11f,
					new FontSizeInfo(20f, 5f)
				},
				{
					12f,
					new FontSizeInfo(21f, 5f)
				},
				{
					14f,
					new FontSizeInfo(25f, 6f)
				},
				{
					16f,
					new FontSizeInfo(27f, 7f)
				},
				{
					18f,
					new FontSizeInfo(31f, 8f)
				},
				{
					20f,
					new FontSizeInfo(35f, 9f)
				},
				{
					22f,
					new FontSizeInfo(38f, 10f)
				},
				{
					24f,
					new FontSizeInfo(41f, 11f)
				},
				{
					26f,
					new FontSizeInfo(45f, 12f)
				},
				{
					28f,
					new FontSizeInfo(48f, 13f)
				},
				{
					36f,
					new FontSizeInfo(61f, 16f)
				},
				{
					48f,
					new FontSizeInfo(82f, 22f)
				},
				{
					72f,
					new FontSizeInfo(0f, 32f)
				},
				{
					96f,
					new FontSizeInfo(162f, 43f)
				},
				{
					128f,
					new FontSizeInfo(216f, 58f)
				},
				{
					256f,
					new FontSizeInfo(431f, 115f)
				}
			}
		},
		{
			"Bodoni MT Poster Compressed",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(20f, 5f)
				},
				{
					11f,
					new FontSizeInfo(20f, 5f)
				},
				{
					12f,
					new FontSizeInfo(21f, 6f)
				},
				{
					14f,
					new FontSizeInfo(24f, 6f)
				},
				{
					16f,
					new FontSizeInfo(27f, 7f)
				},
				{
					18f,
					new FontSizeInfo(30f, 8f)
				},
				{
					20f,
					new FontSizeInfo(34f, 8f)
				},
				{
					22f,
					new FontSizeInfo(36f, 9f)
				},
				{
					24f,
					new FontSizeInfo(40f, 9f)
				},
				{
					26f,
					new FontSizeInfo(44f, 10f)
				},
				{
					28f,
					new FontSizeInfo(46f, 10f)
				},
				{
					36f,
					new FontSizeInfo(59f, 14f)
				},
				{
					48f,
					new FontSizeInfo(79f, 18f)
				},
				{
					72f,
					new FontSizeInfo(0f, 28f)
				},
				{
					96f,
					new FontSizeInfo(157f, 38f)
				},
				{
					128f,
					new FontSizeInfo(209f, 49f)
				},
				{
					256f,
					new FontSizeInfo(416f, 96f)
				}
			}
		},
		{
			"Bookman Old Style",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(21f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 12f)
				},
				{
					16f,
					new FontSizeInfo(27f, 13f)
				},
				{
					18f,
					new FontSizeInfo(31f, 15f)
				},
				{
					20f,
					new FontSizeInfo(34f, 17f)
				},
				{
					22f,
					new FontSizeInfo(37f, 18f)
				},
				{
					24f,
					new FontSizeInfo(42f, 20f)
				},
				{
					26f,
					new FontSizeInfo(44f, 22f)
				},
				{
					28f,
					new FontSizeInfo(47f, 23f)
				},
				{
					36f,
					new FontSizeInfo(61f, 30f)
				},
				{
					48f,
					new FontSizeInfo(82f, 40f)
				},
				{
					72f,
					new FontSizeInfo(0f, 60f)
				},
				{
					96f,
					new FontSizeInfo(163f, 79f)
				},
				{
					128f,
					new FontSizeInfo(218f, 106f)
				},
				{
					256f,
					new FontSizeInfo(437f, 211f)
				}
			}
		},
		{
			"Bradley Hand ITC",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(21f, 8f)
				},
				{
					11f,
					new FontSizeInfo(22f, 10f)
				},
				{
					12f,
					new FontSizeInfo(23f, 10f)
				},
				{
					14f,
					new FontSizeInfo(28f, 12f)
				},
				{
					16f,
					new FontSizeInfo(30f, 14f)
				},
				{
					18f,
					new FontSizeInfo(35f, 15f)
				},
				{
					20f,
					new FontSizeInfo(39f, 17f)
				},
				{
					22f,
					new FontSizeInfo(42f, 19f)
				},
				{
					24f,
					new FontSizeInfo(46f, 21f)
				},
				{
					26f,
					new FontSizeInfo(50f, 23f)
				},
				{
					28f,
					new FontSizeInfo(53f, 24f)
				},
				{
					36f,
					new FontSizeInfo(68f, 31f)
				},
				{
					48f,
					new FontSizeInfo(90f, 41f)
				},
				{
					72f,
					new FontSizeInfo(0f, 62f)
				},
				{
					96f,
					new FontSizeInfo(180f, 82f)
				},
				{
					128f,
					new FontSizeInfo(240f, 110f)
				},
				{
					256f,
					new FontSizeInfo(478f, 219f)
				}
			}
		},
		{
			"Britannic Bold",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 12f)
				},
				{
					16f,
					new FontSizeInfo(26f, 13f)
				},
				{
					18f,
					new FontSizeInfo(30f, 15f)
				},
				{
					20f,
					new FontSizeInfo(34f, 17f)
				},
				{
					22f,
					new FontSizeInfo(36f, 18f)
				},
				{
					24f,
					new FontSizeInfo(40f, 20f)
				},
				{
					26f,
					new FontSizeInfo(43f, 21f)
				},
				{
					28f,
					new FontSizeInfo(46f, 23f)
				},
				{
					36f,
					new FontSizeInfo(59f, 29f)
				},
				{
					48f,
					new FontSizeInfo(79f, 39f)
				},
				{
					72f,
					new FontSizeInfo(0f, 59f)
				},
				{
					96f,
					new FontSizeInfo(156f, 78f)
				},
				{
					128f,
					new FontSizeInfo(208f, 105f)
				},
				{
					256f,
					new FontSizeInfo(414f, 209f)
				}
			}
		},
		{
			"Berlin Sans FB",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(26f, 13f)
				},
				{
					18f,
					new FontSizeInfo(30f, 14f)
				},
				{
					20f,
					new FontSizeInfo(33f, 16f)
				},
				{
					22f,
					new FontSizeInfo(36f, 17f)
				},
				{
					24f,
					new FontSizeInfo(39f, 19f)
				},
				{
					26f,
					new FontSizeInfo(43f, 21f)
				},
				{
					28f,
					new FontSizeInfo(45f, 22f)
				},
				{
					36f,
					new FontSizeInfo(59f, 29f)
				},
				{
					48f,
					new FontSizeInfo(78f, 38f)
				},
				{
					72f,
					new FontSizeInfo(0f, 57f)
				},
				{
					96f,
					new FontSizeInfo(155f, 76f)
				},
				{
					128f,
					new FontSizeInfo(207f, 102f)
				},
				{
					256f,
					new FontSizeInfo(411f, 203f)
				}
			}
		},
		{
			"Berlin Sans FB Demi",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 12f)
				},
				{
					16f,
					new FontSizeInfo(26f, 13f)
				},
				{
					18f,
					new FontSizeInfo(30f, 15f)
				},
				{
					20f,
					new FontSizeInfo(34f, 17f)
				},
				{
					22f,
					new FontSizeInfo(36f, 18f)
				},
				{
					24f,
					new FontSizeInfo(40f, 20f)
				},
				{
					26f,
					new FontSizeInfo(43f, 22f)
				},
				{
					28f,
					new FontSizeInfo(46f, 23f)
				},
				{
					36f,
					new FontSizeInfo(59f, 30f)
				},
				{
					48f,
					new FontSizeInfo(78f, 40f)
				},
				{
					72f,
					new FontSizeInfo(0f, 60f)
				},
				{
					96f,
					new FontSizeInfo(155f, 81f)
				},
				{
					128f,
					new FontSizeInfo(207f, 108f)
				},
				{
					256f,
					new FontSizeInfo(412f, 215f)
				}
			}
		},
		{
			"Broadway",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 10f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 12f)
				},
				{
					16f,
					new FontSizeInfo(27f, 14f)
				},
				{
					18f,
					new FontSizeInfo(30f, 16f)
				},
				{
					20f,
					new FontSizeInfo(34f, 17f)
				},
				{
					22f,
					new FontSizeInfo(36f, 19f)
				},
				{
					24f,
					new FontSizeInfo(40f, 21f)
				},
				{
					26f,
					new FontSizeInfo(44f, 23f)
				},
				{
					28f,
					new FontSizeInfo(46f, 24f)
				},
				{
					36f,
					new FontSizeInfo(60f, 31f)
				},
				{
					48f,
					new FontSizeInfo(79f, 41f)
				},
				{
					72f,
					new FontSizeInfo(0f, 62f)
				},
				{
					96f,
					new FontSizeInfo(157f, 83f)
				},
				{
					128f,
					new FontSizeInfo(210f, 111f)
				},
				{
					256f,
					new FontSizeInfo(417f, 219f)
				}
			}
		},
		{
			"Brush Script MT",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(21f, 8f)
				},
				{
					12f,
					new FontSizeInfo(22f, 8f)
				},
				{
					14f,
					new FontSizeInfo(26f, 10f)
				},
				{
					16f,
					new FontSizeInfo(29f, 11f)
				},
				{
					18f,
					new FontSizeInfo(33f, 12f)
				},
				{
					20f,
					new FontSizeInfo(37f, 14f)
				},
				{
					22f,
					new FontSizeInfo(40f, 15f)
				},
				{
					24f,
					new FontSizeInfo(44f, 16f)
				},
				{
					26f,
					new FontSizeInfo(48f, 18f)
				},
				{
					28f,
					new FontSizeInfo(51f, 19f)
				},
				{
					36f,
					new FontSizeInfo(65f, 24f)
				},
				{
					48f,
					new FontSizeInfo(87f, 33f)
				},
				{
					72f,
					new FontSizeInfo(0f, 49f)
				},
				{
					96f,
					new FontSizeInfo(172f, 65f)
				},
				{
					128f,
					new FontSizeInfo(230f, 87f)
				},
				{
					256f,
					new FontSizeInfo(457f, 174f)
				}
			}
		},
		{
			"Bookshelf Symbol 7",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 8f)
				},
				{
					10f,
					new FontSizeInfo(20f, 10f)
				},
				{
					11f,
					new FontSizeInfo(20f, 11f)
				},
				{
					12f,
					new FontSizeInfo(21f, 12f)
				},
				{
					14f,
					new FontSizeInfo(24f, 14f)
				},
				{
					16f,
					new FontSizeInfo(26f, 16f)
				},
				{
					18f,
					new FontSizeInfo(29f, 18f)
				},
				{
					20f,
					new FontSizeInfo(32f, 20f)
				},
				{
					22f,
					new FontSizeInfo(34f, 21f)
				},
				{
					24f,
					new FontSizeInfo(38f, 24f)
				},
				{
					26f,
					new FontSizeInfo(41f, 26f)
				},
				{
					28f,
					new FontSizeInfo(43f, 27f)
				},
				{
					36f,
					new FontSizeInfo(56f, 35f)
				},
				{
					48f,
					new FontSizeInfo(74f, 47f)
				},
				{
					72f,
					new FontSizeInfo(0f, 71f)
				},
				{
					96f,
					new FontSizeInfo(147f, 95f)
				},
				{
					128f,
					new FontSizeInfo(196f, 126f)
				},
				{
					256f,
					new FontSizeInfo(390f, 251f)
				}
			}
		},
		{
			"Californian FB",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 11f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 13f)
				},
				{
					20f,
					new FontSizeInfo(35f, 14f)
				},
				{
					22f,
					new FontSizeInfo(37f, 15f)
				},
				{
					24f,
					new FontSizeInfo(41f, 17f)
				},
				{
					26f,
					new FontSizeInfo(45f, 18f)
				},
				{
					28f,
					new FontSizeInfo(47f, 19f)
				},
				{
					36f,
					new FontSizeInfo(61f, 26f)
				},
				{
					48f,
					new FontSizeInfo(81f, 34f)
				},
				{
					72f,
					new FontSizeInfo(0f, 51f)
				},
				{
					96f,
					new FontSizeInfo(162f, 68f)
				},
				{
					128f,
					new FontSizeInfo(216f, 91f)
				},
				{
					256f,
					new FontSizeInfo(429f, 181f)
				}
			}
		},
		{
			"Calisto MT",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(24f, 10f)
				},
				{
					16f,
					new FontSizeInfo(27f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 12f)
				},
				{
					20f,
					new FontSizeInfo(34f, 14f)
				},
				{
					22f,
					new FontSizeInfo(37f, 15f)
				},
				{
					24f,
					new FontSizeInfo(40f, 16f)
				},
				{
					26f,
					new FontSizeInfo(44f, 18f)
				},
				{
					28f,
					new FontSizeInfo(46f, 19f)
				},
				{
					36f,
					new FontSizeInfo(60f, 24f)
				},
				{
					48f,
					new FontSizeInfo(80f, 33f)
				},
				{
					72f,
					new FontSizeInfo(0f, 49f)
				},
				{
					96f,
					new FontSizeInfo(158f, 65f)
				},
				{
					128f,
					new FontSizeInfo(211f, 87f)
				},
				{
					256f,
					new FontSizeInfo(420f, 174f)
				}
			}
		},
		{
			"Castellar",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 8f)
				},
				{
					10f,
					new FontSizeInfo(20f, 9f)
				},
				{
					11f,
					new FontSizeInfo(20f, 11f)
				},
				{
					12f,
					new FontSizeInfo(21f, 12f)
				},
				{
					14f,
					new FontSizeInfo(25f, 14f)
				},
				{
					16f,
					new FontSizeInfo(28f, 15f)
				},
				{
					18f,
					new FontSizeInfo(32f, 17f)
				},
				{
					20f,
					new FontSizeInfo(36f, 19f)
				},
				{
					22f,
					new FontSizeInfo(38f, 21f)
				},
				{
					24f,
					new FontSizeInfo(42f, 23f)
				},
				{
					26f,
					new FontSizeInfo(46f, 25f)
				},
				{
					28f,
					new FontSizeInfo(48f, 27f)
				},
				{
					36f,
					new FontSizeInfo(62f, 35f)
				},
				{
					48f,
					new FontSizeInfo(83f, 46f)
				},
				{
					72f,
					new FontSizeInfo(0f, 69f)
				},
				{
					96f,
					new FontSizeInfo(165f, 92f)
				},
				{
					128f,
					new FontSizeInfo(220f, 123f)
				},
				{
					256f,
					new FontSizeInfo(437f, 245f)
				}
			}
		},
		{
			"Century Schoolbook",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(30f, 13f)
				},
				{
					20f,
					new FontSizeInfo(34f, 15f)
				},
				{
					22f,
					new FontSizeInfo(36f, 16f)
				},
				{
					24f,
					new FontSizeInfo(40f, 18f)
				},
				{
					26f,
					new FontSizeInfo(44f, 19f)
				},
				{
					28f,
					new FontSizeInfo(46f, 21f)
				},
				{
					36f,
					new FontSizeInfo(59f, 27f)
				},
				{
					48f,
					new FontSizeInfo(79f, 36f)
				},
				{
					72f,
					new FontSizeInfo(0f, 53f)
				},
				{
					96f,
					new FontSizeInfo(157f, 71f)
				},
				{
					128f,
					new FontSizeInfo(209f, 95f)
				},
				{
					256f,
					new FontSizeInfo(416f, 190f)
				}
			}
		},
		{
			"Centaur",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(20f, 7f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 9f)
				},
				{
					16f,
					new FontSizeInfo(28f, 10f)
				},
				{
					18f,
					new FontSizeInfo(32f, 12f)
				},
				{
					20f,
					new FontSizeInfo(36f, 13f)
				},
				{
					22f,
					new FontSizeInfo(38f, 14f)
				},
				{
					24f,
					new FontSizeInfo(42f, 16f)
				},
				{
					26f,
					new FontSizeInfo(46f, 17f)
				},
				{
					28f,
					new FontSizeInfo(48f, 18f)
				},
				{
					36f,
					new FontSizeInfo(62f, 24f)
				},
				{
					48f,
					new FontSizeInfo(83f, 31f)
				},
				{
					72f,
					new FontSizeInfo(0f, 47f)
				},
				{
					96f,
					new FontSizeInfo(165f, 63f)
				},
				{
					128f,
					new FontSizeInfo(220f, 84f)
				},
				{
					256f,
					new FontSizeInfo(440f, 167f)
				}
			}
		},
		{
			"Chiller",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(21f, 7f)
				},
				{
					12f,
					new FontSizeInfo(22f, 8f)
				},
				{
					14f,
					new FontSizeInfo(26f, 9f)
				},
				{
					16f,
					new FontSizeInfo(28f, 10f)
				},
				{
					18f,
					new FontSizeInfo(32f, 11f)
				},
				{
					20f,
					new FontSizeInfo(36f, 13f)
				},
				{
					22f,
					new FontSizeInfo(39f, 14f)
				},
				{
					24f,
					new FontSizeInfo(43f, 15f)
				},
				{
					26f,
					new FontSizeInfo(47f, 16f)
				},
				{
					28f,
					new FontSizeInfo(49f, 17f)
				},
				{
					36f,
					new FontSizeInfo(64f, 23f)
				},
				{
					48f,
					new FontSizeInfo(85f, 30f)
				},
				{
					72f,
					new FontSizeInfo(0f, 45f)
				},
				{
					96f,
					new FontSizeInfo(168f, 60f)
				},
				{
					128f,
					new FontSizeInfo(224f, 80f)
				},
				{
					256f,
					new FontSizeInfo(446f, 160f)
				}
			}
		},
		{
			"Colonna MT",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(21f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 8f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(29f, 11f)
				},
				{
					18f,
					new FontSizeInfo(33f, 12f)
				},
				{
					20f,
					new FontSizeInfo(37f, 14f)
				},
				{
					22f,
					new FontSizeInfo(40f, 15f)
				},
				{
					24f,
					new FontSizeInfo(44f, 16f)
				},
				{
					26f,
					new FontSizeInfo(48f, 18f)
				},
				{
					28f,
					new FontSizeInfo(51f, 19f)
				},
				{
					36f,
					new FontSizeInfo(66f, 24f)
				},
				{
					48f,
					new FontSizeInfo(87f, 32f)
				},
				{
					72f,
					new FontSizeInfo(0f, 48f)
				},
				{
					96f,
					new FontSizeInfo(173f, 64f)
				},
				{
					128f,
					new FontSizeInfo(231f, 86f)
				},
				{
					256f,
					new FontSizeInfo(459f, 171f)
				}
			}
		},
		{
			"Cooper Black",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(27f, 13f)
				},
				{
					18f,
					new FontSizeInfo(30f, 14f)
				},
				{
					20f,
					new FontSizeInfo(34f, 16f)
				},
				{
					22f,
					new FontSizeInfo(37f, 17f)
				},
				{
					24f,
					new FontSizeInfo(40f, 19f)
				},
				{
					26f,
					new FontSizeInfo(44f, 21f)
				},
				{
					28f,
					new FontSizeInfo(46f, 22f)
				},
				{
					36f,
					new FontSizeInfo(60f, 29f)
				},
				{
					48f,
					new FontSizeInfo(80f, 38f)
				},
				{
					72f,
					new FontSizeInfo(0f, 58f)
				},
				{
					96f,
					new FontSizeInfo(158f, 77f)
				},
				{
					128f,
					new FontSizeInfo(211f, 103f)
				},
				{
					256f,
					new FontSizeInfo(420f, 205f)
				}
			}
		},
		{
			"Copperplate Gothic Bold",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 8f)
				},
				{
					10f,
					new FontSizeInfo(20f, 9f)
				},
				{
					11f,
					new FontSizeInfo(20f, 11f)
				},
				{
					12f,
					new FontSizeInfo(21f, 11f)
				},
				{
					14f,
					new FontSizeInfo(24f, 14f)
				},
				{
					16f,
					new FontSizeInfo(27f, 15f)
				},
				{
					18f,
					new FontSizeInfo(30f, 17f)
				},
				{
					20f,
					new FontSizeInfo(34f, 19f)
				},
				{
					22f,
					new FontSizeInfo(37f, 21f)
				},
				{
					24f,
					new FontSizeInfo(40f, 23f)
				},
				{
					26f,
					new FontSizeInfo(44f, 25f)
				},
				{
					28f,
					new FontSizeInfo(46f, 26f)
				},
				{
					36f,
					new FontSizeInfo(60f, 34f)
				},
				{
					48f,
					new FontSizeInfo(80f, 46f)
				},
				{
					72f,
					new FontSizeInfo(0f, 68f)
				},
				{
					96f,
					new FontSizeInfo(158f, 91f)
				},
				{
					128f,
					new FontSizeInfo(211f, 122f)
				},
				{
					256f,
					new FontSizeInfo(420f, 243f)
				}
			}
		},
		{
			"Copperplate Gothic Light",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 8f)
				},
				{
					10f,
					new FontSizeInfo(20f, 9f)
				},
				{
					11f,
					new FontSizeInfo(20f, 11f)
				},
				{
					12f,
					new FontSizeInfo(21f, 12f)
				},
				{
					14f,
					new FontSizeInfo(24f, 14f)
				},
				{
					16f,
					new FontSizeInfo(26f, 15f)
				},
				{
					18f,
					new FontSizeInfo(30f, 17f)
				},
				{
					20f,
					new FontSizeInfo(34f, 19f)
				},
				{
					22f,
					new FontSizeInfo(36f, 21f)
				},
				{
					24f,
					new FontSizeInfo(40f, 23f)
				},
				{
					26f,
					new FontSizeInfo(43f, 25f)
				},
				{
					28f,
					new FontSizeInfo(46f, 27f)
				},
				{
					36f,
					new FontSizeInfo(59f, 35f)
				},
				{
					48f,
					new FontSizeInfo(78f, 46f)
				},
				{
					72f,
					new FontSizeInfo(0f, 69f)
				},
				{
					96f,
					new FontSizeInfo(157f, 92f)
				},
				{
					128f,
					new FontSizeInfo(209f, 123f)
				},
				{
					256f,
					new FontSizeInfo(418f, 246f)
				}
			}
		},
		{
			"Curlz MT",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(22f, 9f)
				},
				{
					14f,
					new FontSizeInfo(28f, 10f)
				},
				{
					16f,
					new FontSizeInfo(30f, 11f)
				},
				{
					18f,
					new FontSizeInfo(34f, 13f)
				},
				{
					20f,
					new FontSizeInfo(38f, 15f)
				},
				{
					22f,
					new FontSizeInfo(41f, 16f)
				},
				{
					24f,
					new FontSizeInfo(45f, 17f)
				},
				{
					26f,
					new FontSizeInfo(48f, 19f)
				},
				{
					28f,
					new FontSizeInfo(51f, 20f)
				},
				{
					36f,
					new FontSizeInfo(65f, 26f)
				},
				{
					48f,
					new FontSizeInfo(88f, 34f)
				},
				{
					72f,
					new FontSizeInfo(0f, 52f)
				},
				{
					96f,
					new FontSizeInfo(173f, 69f)
				},
				{
					128f,
					new FontSizeInfo(233f, 92f)
				},
				{
					256f,
					new FontSizeInfo(462f, 183f)
				}
			}
		},
		{
			"Elephant",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 8f)
				},
				{
					10f,
					new FontSizeInfo(20f, 10f)
				},
				{
					11f,
					new FontSizeInfo(21f, 11f)
				},
				{
					12f,
					new FontSizeInfo(22f, 12f)
				},
				{
					14f,
					new FontSizeInfo(26f, 14f)
				},
				{
					16f,
					new FontSizeInfo(28f, 16f)
				},
				{
					18f,
					new FontSizeInfo(32f, 18f)
				},
				{
					20f,
					new FontSizeInfo(36f, 20f)
				},
				{
					22f,
					new FontSizeInfo(39f, 22f)
				},
				{
					24f,
					new FontSizeInfo(43f, 24f)
				},
				{
					26f,
					new FontSizeInfo(47f, 27f)
				},
				{
					28f,
					new FontSizeInfo(49f, 28f)
				},
				{
					36f,
					new FontSizeInfo(64f, 36f)
				},
				{
					48f,
					new FontSizeInfo(85f, 49f)
				},
				{
					72f,
					new FontSizeInfo(0f, 73f)
				},
				{
					96f,
					new FontSizeInfo(168f, 97f)
				},
				{
					128f,
					new FontSizeInfo(224f, 130f)
				},
				{
					256f,
					new FontSizeInfo(446f, 258f)
				}
			}
		},
		{
			"Engravers MT",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 10f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 12f)
				},
				{
					16f,
					new FontSizeInfo(27f, 14f)
				},
				{
					18f,
					new FontSizeInfo(30f, 16f)
				},
				{
					20f,
					new FontSizeInfo(34f, 17f)
				},
				{
					22f,
					new FontSizeInfo(37f, 19f)
				},
				{
					24f,
					new FontSizeInfo(40f, 21f)
				},
				{
					26f,
					new FontSizeInfo(44f, 23f)
				},
				{
					28f,
					new FontSizeInfo(46f, 24f)
				},
				{
					36f,
					new FontSizeInfo(60f, 31f)
				},
				{
					48f,
					new FontSizeInfo(80f, 41f)
				},
				{
					72f,
					new FontSizeInfo(0f, 62f)
				},
				{
					96f,
					new FontSizeInfo(158f, 83f)
				},
				{
					128f,
					new FontSizeInfo(211f, 110f)
				},
				{
					256f,
					new FontSizeInfo(419f, 219f)
				}
			}
		},
		{
			"Eras Bold ITC",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 10f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(25f, 12f)
				},
				{
					16f,
					new FontSizeInfo(27f, 14f)
				},
				{
					18f,
					new FontSizeInfo(31f, 16f)
				},
				{
					20f,
					new FontSizeInfo(35f, 18f)
				},
				{
					22f,
					new FontSizeInfo(37f, 19f)
				},
				{
					24f,
					new FontSizeInfo(41f, 21f)
				},
				{
					26f,
					new FontSizeInfo(45f, 23f)
				},
				{
					28f,
					new FontSizeInfo(47f, 24f)
				},
				{
					36f,
					new FontSizeInfo(61f, 31f)
				},
				{
					48f,
					new FontSizeInfo(81f, 42f)
				},
				{
					72f,
					new FontSizeInfo(0f, 63f)
				},
				{
					96f,
					new FontSizeInfo(161f, 83f)
				},
				{
					128f,
					new FontSizeInfo(215f, 111f)
				},
				{
					256f,
					new FontSizeInfo(425f, 222f)
				}
			}
		},
		{
			"Eras Demi ITC",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(25f, 12f)
				},
				{
					16f,
					new FontSizeInfo(27f, 13f)
				},
				{
					18f,
					new FontSizeInfo(31f, 15f)
				},
				{
					20f,
					new FontSizeInfo(35f, 16f)
				},
				{
					22f,
					new FontSizeInfo(37f, 18f)
				},
				{
					24f,
					new FontSizeInfo(41f, 20f)
				},
				{
					26f,
					new FontSizeInfo(45f, 21f)
				},
				{
					28f,
					new FontSizeInfo(47f, 23f)
				},
				{
					36f,
					new FontSizeInfo(61f, 29f)
				},
				{
					48f,
					new FontSizeInfo(81f, 39f)
				},
				{
					72f,
					new FontSizeInfo(0f, 59f)
				},
				{
					96f,
					new FontSizeInfo(161f, 78f)
				},
				{
					128f,
					new FontSizeInfo(215f, 104f)
				},
				{
					256f,
					new FontSizeInfo(425f, 208f)
				}
			}
		},
		{
			"Eras Light ITC",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(22f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 13f)
				},
				{
					20f,
					new FontSizeInfo(35f, 14f)
				},
				{
					22f,
					new FontSizeInfo(37f, 16f)
				},
				{
					24f,
					new FontSizeInfo(41f, 17f)
				},
				{
					26f,
					new FontSizeInfo(45f, 19f)
				},
				{
					28f,
					new FontSizeInfo(47f, 20f)
				},
				{
					36f,
					new FontSizeInfo(61f, 26f)
				},
				{
					48f,
					new FontSizeInfo(81f, 34f)
				},
				{
					72f,
					new FontSizeInfo(0f, 51f)
				},
				{
					96f,
					new FontSizeInfo(161f, 69f)
				},
				{
					128f,
					new FontSizeInfo(215f, 92f)
				},
				{
					256f,
					new FontSizeInfo(427f, 183f)
				}
			}
		},
		{
			"Eras Medium ITC",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 11f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 14f)
				},
				{
					20f,
					new FontSizeInfo(35f, 15f)
				},
				{
					22f,
					new FontSizeInfo(37f, 17f)
				},
				{
					24f,
					new FontSizeInfo(41f, 18f)
				},
				{
					26f,
					new FontSizeInfo(45f, 20f)
				},
				{
					28f,
					new FontSizeInfo(47f, 21f)
				},
				{
					36f,
					new FontSizeInfo(61f, 27f)
				},
				{
					48f,
					new FontSizeInfo(81f, 37f)
				},
				{
					72f,
					new FontSizeInfo(0f, 55f)
				},
				{
					96f,
					new FontSizeInfo(161f, 73f)
				},
				{
					128f,
					new FontSizeInfo(215f, 98f)
				},
				{
					256f,
					new FontSizeInfo(426f, 195f)
				}
			}
		},
		{
			"Felix Titling",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 11f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 14f)
				},
				{
					20f,
					new FontSizeInfo(35f, 16f)
				},
				{
					22f,
					new FontSizeInfo(37f, 17f)
				},
				{
					24f,
					new FontSizeInfo(41f, 19f)
				},
				{
					26f,
					new FontSizeInfo(44f, 20f)
				},
				{
					28f,
					new FontSizeInfo(47f, 22f)
				},
				{
					36f,
					new FontSizeInfo(61f, 28f)
				},
				{
					48f,
					new FontSizeInfo(80f, 37f)
				},
				{
					72f,
					new FontSizeInfo(0f, 56f)
				},
				{
					96f,
					new FontSizeInfo(160f, 75f)
				},
				{
					128f,
					new FontSizeInfo(213f, 100f)
				},
				{
					256f,
					new FontSizeInfo(424f, 198f)
				}
			}
		},
		{
			"Forte",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(21f, 8f)
				},
				{
					12f,
					new FontSizeInfo(22f, 8f)
				},
				{
					14f,
					new FontSizeInfo(28f, 10f)
				},
				{
					16f,
					new FontSizeInfo(30f, 11f)
				},
				{
					18f,
					new FontSizeInfo(34f, 13f)
				},
				{
					20f,
					new FontSizeInfo(40f, 14f)
				},
				{
					22f,
					new FontSizeInfo(42f, 15f)
				},
				{
					24f,
					new FontSizeInfo(46f, 17f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(71f, 25f)
				},
				{
					48f,
					new FontSizeInfo(95f, 34f)
				},
				{
					72f,
					new FontSizeInfo(0f, 51f)
				},
				{
					96f,
					new FontSizeInfo(188f, 68f)
				},
				{
					128f,
					new FontSizeInfo(253f, 91f)
				},
				{
					256f,
					new FontSizeInfo(503f, 181f)
				}
			}
		},
		{
			"Franklin Gothic Book",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(21f, 9f)
				},
				{
					12f,
					new FontSizeInfo(22f, 9f)
				},
				{
					14f,
					new FontSizeInfo(26f, 11f)
				},
				{
					16f,
					new FontSizeInfo(28f, 12f)
				},
				{
					18f,
					new FontSizeInfo(32f, 14f)
				},
				{
					20f,
					new FontSizeInfo(36f, 16f)
				},
				{
					22f,
					new FontSizeInfo(39f, 17f)
				},
				{
					24f,
					new FontSizeInfo(40f, 19f)
				},
				{
					26f,
					new FontSizeInfo(44f, 21f)
				},
				{
					28f,
					new FontSizeInfo(46f, 22f)
				},
				{
					36f,
					new FontSizeInfo(64f, 28f)
				},
				{
					48f,
					new FontSizeInfo(85f, 38f)
				},
				{
					72f,
					new FontSizeInfo(0f, 56f)
				},
				{
					96f,
					new FontSizeInfo(168f, 75f)
				},
				{
					128f,
					new FontSizeInfo(224f, 100f)
				},
				{
					256f,
					new FontSizeInfo(416f, 200f)
				}
			}
		},
		{
			"Franklin Gothic Demi",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(21f, 9f)
				},
				{
					12f,
					new FontSizeInfo(22f, 9f)
				},
				{
					14f,
					new FontSizeInfo(26f, 11f)
				},
				{
					16f,
					new FontSizeInfo(28f, 12f)
				},
				{
					18f,
					new FontSizeInfo(32f, 14f)
				},
				{
					20f,
					new FontSizeInfo(36f, 16f)
				},
				{
					22f,
					new FontSizeInfo(39f, 17f)
				},
				{
					24f,
					new FontSizeInfo(40f, 19f)
				},
				{
					26f,
					new FontSizeInfo(44f, 21f)
				},
				{
					28f,
					new FontSizeInfo(46f, 22f)
				},
				{
					36f,
					new FontSizeInfo(64f, 28f)
				},
				{
					48f,
					new FontSizeInfo(85f, 38f)
				},
				{
					72f,
					new FontSizeInfo(0f, 56f)
				},
				{
					96f,
					new FontSizeInfo(168f, 75f)
				},
				{
					128f,
					new FontSizeInfo(224f, 100f)
				},
				{
					256f,
					new FontSizeInfo(416f, 200f)
				}
			}
		},
		{
			"Franklin Gothic Demi Cond",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(21f, 8f)
				},
				{
					12f,
					new FontSizeInfo(22f, 8f)
				},
				{
					14f,
					new FontSizeInfo(26f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(32f, 12f)
				},
				{
					20f,
					new FontSizeInfo(36f, 14f)
				},
				{
					22f,
					new FontSizeInfo(39f, 15f)
				},
				{
					24f,
					new FontSizeInfo(40f, 16f)
				},
				{
					26f,
					new FontSizeInfo(44f, 18f)
				},
				{
					28f,
					new FontSizeInfo(46f, 19f)
				},
				{
					36f,
					new FontSizeInfo(64f, 25f)
				},
				{
					48f,
					new FontSizeInfo(85f, 33f)
				},
				{
					72f,
					new FontSizeInfo(0f, 49f)
				},
				{
					96f,
					new FontSizeInfo(168f, 65f)
				},
				{
					128f,
					new FontSizeInfo(224f, 87f)
				},
				{
					256f,
					new FontSizeInfo(416f, 174f)
				}
			}
		},
		{
			"Franklin Gothic Heavy",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(21f, 9f)
				},
				{
					12f,
					new FontSizeInfo(22f, 9f)
				},
				{
					14f,
					new FontSizeInfo(26f, 11f)
				},
				{
					16f,
					new FontSizeInfo(28f, 12f)
				},
				{
					18f,
					new FontSizeInfo(32f, 14f)
				},
				{
					20f,
					new FontSizeInfo(36f, 16f)
				},
				{
					22f,
					new FontSizeInfo(39f, 17f)
				},
				{
					24f,
					new FontSizeInfo(40f, 19f)
				},
				{
					26f,
					new FontSizeInfo(44f, 21f)
				},
				{
					28f,
					new FontSizeInfo(46f, 22f)
				},
				{
					36f,
					new FontSizeInfo(64f, 28f)
				},
				{
					48f,
					new FontSizeInfo(85f, 38f)
				},
				{
					72f,
					new FontSizeInfo(0f, 56f)
				},
				{
					96f,
					new FontSizeInfo(168f, 75f)
				},
				{
					128f,
					new FontSizeInfo(224f, 100f)
				},
				{
					256f,
					new FontSizeInfo(416f, 200f)
				}
			}
		},
		{
			"Franklin Gothic Medium Cond",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(21f, 8f)
				},
				{
					12f,
					new FontSizeInfo(22f, 8f)
				},
				{
					14f,
					new FontSizeInfo(26f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(32f, 12f)
				},
				{
					20f,
					new FontSizeInfo(36f, 14f)
				},
				{
					22f,
					new FontSizeInfo(39f, 15f)
				},
				{
					24f,
					new FontSizeInfo(40f, 16f)
				},
				{
					26f,
					new FontSizeInfo(44f, 18f)
				},
				{
					28f,
					new FontSizeInfo(46f, 19f)
				},
				{
					36f,
					new FontSizeInfo(64f, 25f)
				},
				{
					48f,
					new FontSizeInfo(85f, 33f)
				},
				{
					72f,
					new FontSizeInfo(0f, 49f)
				},
				{
					96f,
					new FontSizeInfo(168f, 65f)
				},
				{
					128f,
					new FontSizeInfo(224f, 87f)
				},
				{
					256f,
					new FontSizeInfo(416f, 174f)
				}
			}
		},
		{
			"Freestyle Script",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(20f, 5f)
				},
				{
					11f,
					new FontSizeInfo(21f, 6f)
				},
				{
					12f,
					new FontSizeInfo(23f, 6f)
				},
				{
					14f,
					new FontSizeInfo(26f, 7f)
				},
				{
					16f,
					new FontSizeInfo(29f, 8f)
				},
				{
					18f,
					new FontSizeInfo(33f, 9f)
				},
				{
					20f,
					new FontSizeInfo(37f, 10f)
				},
				{
					22f,
					new FontSizeInfo(40f, 10f)
				},
				{
					24f,
					new FontSizeInfo(43f, 12f)
				},
				{
					26f,
					new FontSizeInfo(48f, 13f)
				},
				{
					28f,
					new FontSizeInfo(50f, 13f)
				},
				{
					36f,
					new FontSizeInfo(65f, 17f)
				},
				{
					48f,
					new FontSizeInfo(86f, 23f)
				},
				{
					72f,
					new FontSizeInfo(0f, 33f)
				},
				{
					96f,
					new FontSizeInfo(172f, 44f)
				},
				{
					128f,
					new FontSizeInfo(228f, 58f)
				},
				{
					256f,
					new FontSizeInfo(454f, 117f)
				}
			}
		},
		{
			"French Script MT",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(20f, 5f)
				},
				{
					11f,
					new FontSizeInfo(21f, 5f)
				},
				{
					12f,
					new FontSizeInfo(22f, 6f)
				},
				{
					14f,
					new FontSizeInfo(26f, 7f)
				},
				{
					16f,
					new FontSizeInfo(28f, 8f)
				},
				{
					18f,
					new FontSizeInfo(32f, 9f)
				},
				{
					20f,
					new FontSizeInfo(36f, 10f)
				},
				{
					22f,
					new FontSizeInfo(39f, 11f)
				},
				{
					24f,
					new FontSizeInfo(43f, 12f)
				},
				{
					26f,
					new FontSizeInfo(47f, 13f)
				},
				{
					28f,
					new FontSizeInfo(49f, 14f)
				},
				{
					36f,
					new FontSizeInfo(64f, 18f)
				},
				{
					48f,
					new FontSizeInfo(85f, 23f)
				},
				{
					72f,
					new FontSizeInfo(0f, 35f)
				},
				{
					96f,
					new FontSizeInfo(168f, 47f)
				},
				{
					128f,
					new FontSizeInfo(224f, 62f)
				},
				{
					256f,
					new FontSizeInfo(446f, 125f)
				}
			}
		},
		{
			"Footlight MT Light",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(30f, 13f)
				},
				{
					20f,
					new FontSizeInfo(34f, 15f)
				},
				{
					22f,
					new FontSizeInfo(37f, 16f)
				},
				{
					24f,
					new FontSizeInfo(40f, 18f)
				},
				{
					26f,
					new FontSizeInfo(44f, 19f)
				},
				{
					28f,
					new FontSizeInfo(46f, 20f)
				},
				{
					36f,
					new FontSizeInfo(60f, 27f)
				},
				{
					48f,
					new FontSizeInfo(79f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 53f)
				},
				{
					96f,
					new FontSizeInfo(158f, 71f)
				},
				{
					128f,
					new FontSizeInfo(210f, 94f)
				},
				{
					256f,
					new FontSizeInfo(419f, 188f)
				}
			}
		},
		{
			"Garamond",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(20f, 7f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 9f)
				},
				{
					16f,
					new FontSizeInfo(28f, 10f)
				},
				{
					18f,
					new FontSizeInfo(31f, 11f)
				},
				{
					20f,
					new FontSizeInfo(35f, 13f)
				},
				{
					22f,
					new FontSizeInfo(38f, 14f)
				},
				{
					24f,
					new FontSizeInfo(41f, 15f)
				},
				{
					26f,
					new FontSizeInfo(45f, 16f)
				},
				{
					28f,
					new FontSizeInfo(48f, 17f)
				},
				{
					36f,
					new FontSizeInfo(62f, 23f)
				},
				{
					48f,
					new FontSizeInfo(82f, 30f)
				},
				{
					72f,
					new FontSizeInfo(0f, 45f)
				},
				{
					96f,
					new FontSizeInfo(163f, 60f)
				},
				{
					128f,
					new FontSizeInfo(217f, 80f)
				},
				{
					256f,
					new FontSizeInfo(432f, 160f)
				}
			}
		},
		{
			"Gigi",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 5f)
				},
				{
					10f,
					new FontSizeInfo(22f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 7f)
				},
				{
					12f,
					new FontSizeInfo(24f, 7f)
				},
				{
					14f,
					new FontSizeInfo(28f, 10f)
				},
				{
					16f,
					new FontSizeInfo(31f, 12f)
				},
				{
					18f,
					new FontSizeInfo(35f, 12f)
				},
				{
					20f,
					new FontSizeInfo(39f, 14f)
				},
				{
					22f,
					new FontSizeInfo(42f, 14f)
				},
				{
					24f,
					new FontSizeInfo(46f, 17f)
				},
				{
					26f,
					new FontSizeInfo(50f, 19f)
				},
				{
					28f,
					new FontSizeInfo(53f, 19f)
				},
				{
					36f,
					new FontSizeInfo(69f, 24f)
				},
				{
					48f,
					new FontSizeInfo(91f, 33f)
				},
				{
					72f,
					new FontSizeInfo(0f, 50f)
				},
				{
					96f,
					new FontSizeInfo(182f, 66f)
				},
				{
					128f,
					new FontSizeInfo(242f, 88f)
				},
				{
					256f,
					new FontSizeInfo(482f, 175f)
				}
			}
		},
		{
			"Gill Sans MT",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 6f)
				},
				{
					10f,
					new FontSizeInfo(21f, 7f)
				},
				{
					11f,
					new FontSizeInfo(23f, 8f)
				},
				{
					12f,
					new FontSizeInfo(26f, 8f)
				},
				{
					14f,
					new FontSizeInfo(29f, 10f)
				},
				{
					16f,
					new FontSizeInfo(33f, 11f)
				},
				{
					18f,
					new FontSizeInfo(37f, 12f)
				},
				{
					20f,
					new FontSizeInfo(41f, 14f)
				},
				{
					22f,
					new FontSizeInfo(43f, 15f)
				},
				{
					24f,
					new FontSizeInfo(48f, 16f)
				},
				{
					26f,
					new FontSizeInfo(51f, 18f)
				},
				{
					28f,
					new FontSizeInfo(56f, 19f)
				},
				{
					36f,
					new FontSizeInfo(71f, 24f)
				},
				{
					48f,
					new FontSizeInfo(91f, 32f)
				},
				{
					72f,
					new FontSizeInfo(0f, 48f)
				},
				{
					96f,
					new FontSizeInfo(184f, 64f)
				},
				{
					128f,
					new FontSizeInfo(243f, 86f)
				},
				{
					256f,
					new FontSizeInfo(421f, 171f)
				}
			}
		},
		{
			"Gill Sans MT Condensed",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(21f, 5f)
				},
				{
					11f,
					new FontSizeInfo(21f, 5f)
				},
				{
					12f,
					new FontSizeInfo(22f, 6f)
				},
				{
					14f,
					new FontSizeInfo(27f, 7f)
				},
				{
					16f,
					new FontSizeInfo(29f, 8f)
				},
				{
					18f,
					new FontSizeInfo(32f, 9f)
				},
				{
					20f,
					new FontSizeInfo(37f, 10f)
				},
				{
					22f,
					new FontSizeInfo(39f, 11f)
				},
				{
					24f,
					new FontSizeInfo(43f, 12f)
				},
				{
					26f,
					new FontSizeInfo(47f, 13f)
				},
				{
					28f,
					new FontSizeInfo(50f, 14f)
				},
				{
					36f,
					new FontSizeInfo(64f, 18f)
				},
				{
					48f,
					new FontSizeInfo(86f, 23f)
				},
				{
					72f,
					new FontSizeInfo(0f, 35f)
				},
				{
					96f,
					new FontSizeInfo(169f, 47f)
				},
				{
					128f,
					new FontSizeInfo(226f, 62f)
				},
				{
					256f,
					new FontSizeInfo(434f, 125f)
				}
			}
		},
		{
			"Gill Sans Ultra Bold Condensed",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(22f, 8f)
				},
				{
					11f,
					new FontSizeInfo(21f, 9f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 11f)
				},
				{
					16f,
					new FontSizeInfo(30f, 12f)
				},
				{
					18f,
					new FontSizeInfo(33f, 14f)
				},
				{
					20f,
					new FontSizeInfo(38f, 16f)
				},
				{
					22f,
					new FontSizeInfo(41f, 17f)
				},
				{
					24f,
					new FontSizeInfo(45f, 19f)
				},
				{
					26f,
					new FontSizeInfo(49f, 20f)
				},
				{
					28f,
					new FontSizeInfo(51f, 22f)
				},
				{
					36f,
					new FontSizeInfo(69f, 28f)
				},
				{
					48f,
					new FontSizeInfo(89f, 37f)
				},
				{
					72f,
					new FontSizeInfo(0f, 56f)
				},
				{
					96f,
					new FontSizeInfo(176f, 74f)
				},
				{
					128f,
					new FontSizeInfo(235f, 99f)
				},
				{
					256f,
					new FontSizeInfo(428f, 198f)
				}
			}
		},
		{
			"Gill Sans Ultra Bold",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(23f, 9f)
				},
				{
					10f,
					new FontSizeInfo(22f, 11f)
				},
				{
					11f,
					new FontSizeInfo(21f, 13f)
				},
				{
					12f,
					new FontSizeInfo(24f, 14f)
				},
				{
					14f,
					new FontSizeInfo(27f, 16f)
				},
				{
					16f,
					new FontSizeInfo(31f, 18f)
				},
				{
					18f,
					new FontSizeInfo(35f, 21f)
				},
				{
					20f,
					new FontSizeInfo(41f, 23f)
				},
				{
					22f,
					new FontSizeInfo(44f, 25f)
				},
				{
					24f,
					new FontSizeInfo(45f, 27f)
				},
				{
					26f,
					new FontSizeInfo(50f, 30f)
				},
				{
					28f,
					new FontSizeInfo(53f, 32f)
				},
				{
					36f,
					new FontSizeInfo(71f, 41f)
				},
				{
					48f,
					new FontSizeInfo(93f, 55f)
				},
				{
					72f,
					new FontSizeInfo(0f, 82f)
				},
				{
					96f,
					new FontSizeInfo(181f, 109f)
				},
				{
					128f,
					new FontSizeInfo(241f, 146f)
				},
				{
					256f,
					new FontSizeInfo(428f, 291f)
				}
			}
		},
		{
			"Gloucester MT Extra Condensed",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(20f, 5f)
				},
				{
					11f,
					new FontSizeInfo(20f, 6f)
				},
				{
					12f,
					new FontSizeInfo(21f, 6f)
				},
				{
					14f,
					new FontSizeInfo(24f, 7f)
				},
				{
					16f,
					new FontSizeInfo(27f, 8f)
				},
				{
					18f,
					new FontSizeInfo(30f, 9f)
				},
				{
					20f,
					new FontSizeInfo(34f, 10f)
				},
				{
					22f,
					new FontSizeInfo(36f, 11f)
				},
				{
					24f,
					new FontSizeInfo(40f, 12f)
				},
				{
					26f,
					new FontSizeInfo(44f, 14f)
				},
				{
					28f,
					new FontSizeInfo(46f, 14f)
				},
				{
					36f,
					new FontSizeInfo(59f, 19f)
				},
				{
					48f,
					new FontSizeInfo(79f, 25f)
				},
				{
					72f,
					new FontSizeInfo(0f, 37f)
				},
				{
					96f,
					new FontSizeInfo(157f, 49f)
				},
				{
					128f,
					new FontSizeInfo(209f, 66f)
				},
				{
					256f,
					new FontSizeInfo(416f, 132f)
				}
			}
		},
		{
			"Gill Sans MT Ext Condensed Bold",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(20f, 4f)
				},
				{
					11f,
					new FontSizeInfo(20f, 5f)
				},
				{
					12f,
					new FontSizeInfo(23f, 6f)
				},
				{
					14f,
					new FontSizeInfo(25f, 7f)
				},
				{
					16f,
					new FontSizeInfo(27f, 7f)
				},
				{
					18f,
					new FontSizeInfo(33f, 8f)
				},
				{
					20f,
					new FontSizeInfo(35f, 9f)
				},
				{
					22f,
					new FontSizeInfo(37f, 10f)
				},
				{
					24f,
					new FontSizeInfo(41f, 11f)
				},
				{
					26f,
					new FontSizeInfo(45f, 12f)
				},
				{
					28f,
					new FontSizeInfo(47f, 13f)
				},
				{
					36f,
					new FontSizeInfo(61f, 17f)
				},
				{
					48f,
					new FontSizeInfo(81f, 22f)
				},
				{
					72f,
					new FontSizeInfo(0f, 33f)
				},
				{
					96f,
					new FontSizeInfo(161f, 44f)
				},
				{
					128f,
					new FontSizeInfo(214f, 59f)
				},
				{
					256f,
					new FontSizeInfo(410f, 117f)
				}
			}
		},
		{
			"Century Gothic",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(26f, 12f)
				},
				{
					18f,
					new FontSizeInfo(32f, 13f)
				},
				{
					20f,
					new FontSizeInfo(35f, 15f)
				},
				{
					22f,
					new FontSizeInfo(38f, 16f)
				},
				{
					24f,
					new FontSizeInfo(41f, 18f)
				},
				{
					26f,
					new FontSizeInfo(44f, 19f)
				},
				{
					28f,
					new FontSizeInfo(46f, 21f)
				},
				{
					36f,
					new FontSizeInfo(61f, 27f)
				},
				{
					48f,
					new FontSizeInfo(82f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 53f)
				},
				{
					96f,
					new FontSizeInfo(158f, 71f)
				},
				{
					128f,
					new FontSizeInfo(214f, 95f)
				},
				{
					256f,
					new FontSizeInfo(427f, 189f)
				}
			}
		},
		{
			"Goudy Old Style",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(21f, 8f)
				},
				{
					12f,
					new FontSizeInfo(22f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(28f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 12f)
				},
				{
					20f,
					new FontSizeInfo(35f, 14f)
				},
				{
					22f,
					new FontSizeInfo(40f, 15f)
				},
				{
					24f,
					new FontSizeInfo(43f, 16f)
				},
				{
					26f,
					new FontSizeInfo(47f, 18f)
				},
				{
					28f,
					new FontSizeInfo(50f, 19f)
				},
				{
					36f,
					new FontSizeInfo(63f, 24f)
				},
				{
					48f,
					new FontSizeInfo(86f, 32f)
				},
				{
					72f,
					new FontSizeInfo(0f, 48f)
				},
				{
					96f,
					new FontSizeInfo(173f, 64f)
				},
				{
					128f,
					new FontSizeInfo(231f, 86f)
				},
				{
					256f,
					new FontSizeInfo(451f, 171f)
				}
			}
		},
		{
			"Goudy Stout",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 9f)
				},
				{
					10f,
					new FontSizeInfo(21f, 11f)
				},
				{
					11f,
					new FontSizeInfo(21f, 12f)
				},
				{
					12f,
					new FontSizeInfo(23f, 13f)
				},
				{
					14f,
					new FontSizeInfo(27f, 16f)
				},
				{
					16f,
					new FontSizeInfo(29f, 17f)
				},
				{
					18f,
					new FontSizeInfo(33f, 20f)
				},
				{
					20f,
					new FontSizeInfo(39f, 22f)
				},
				{
					22f,
					new FontSizeInfo(42f, 24f)
				},
				{
					24f,
					new FontSizeInfo(46f, 26f)
				},
				{
					26f,
					new FontSizeInfo(50f, 29f)
				},
				{
					28f,
					new FontSizeInfo(53f, 31f)
				},
				{
					36f,
					new FontSizeInfo(68f, 40f)
				},
				{
					48f,
					new FontSizeInfo(89f, 53f)
				},
				{
					72f,
					new FontSizeInfo(0f, 79f)
				},
				{
					96f,
					new FontSizeInfo(178f, 106f)
				},
				{
					128f,
					new FontSizeInfo(238f, 141f)
				},
				{
					256f,
					new FontSizeInfo(473f, 280f)
				}
			}
		},
		{
			"Harlow Solid Italic",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(21f, 6f)
				},
				{
					11f,
					new FontSizeInfo(22f, 7f)
				},
				{
					12f,
					new FontSizeInfo(23f, 7f)
				},
				{
					14f,
					new FontSizeInfo(27f, 9f)
				},
				{
					16f,
					new FontSizeInfo(30f, 10f)
				},
				{
					18f,
					new FontSizeInfo(34f, 11f)
				},
				{
					20f,
					new FontSizeInfo(38f, 13f)
				},
				{
					22f,
					new FontSizeInfo(41f, 14f)
				},
				{
					24f,
					new FontSizeInfo(45f, 15f)
				},
				{
					26f,
					new FontSizeInfo(50f, 16f)
				},
				{
					28f,
					new FontSizeInfo(52f, 17f)
				},
				{
					36f,
					new FontSizeInfo(68f, 22f)
				},
				{
					48f,
					new FontSizeInfo(90f, 30f)
				},
				{
					72f,
					new FontSizeInfo(0f, 45f)
				},
				{
					96f,
					new FontSizeInfo(179f, 60f)
				},
				{
					128f,
					new FontSizeInfo(238f, 80f)
				},
				{
					256f,
					new FontSizeInfo(474f, 159f)
				}
			}
		},
		{
			"Harrington",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 13f)
				},
				{
					20f,
					new FontSizeInfo(34f, 15f)
				},
				{
					22f,
					new FontSizeInfo(37f, 16f)
				},
				{
					24f,
					new FontSizeInfo(40f, 18f)
				},
				{
					26f,
					new FontSizeInfo(44f, 19f)
				},
				{
					28f,
					new FontSizeInfo(47f, 21f)
				},
				{
					36f,
					new FontSizeInfo(60f, 27f)
				},
				{
					48f,
					new FontSizeInfo(80f, 36f)
				},
				{
					72f,
					new FontSizeInfo(0f, 53f)
				},
				{
					96f,
					new FontSizeInfo(159f, 71f)
				},
				{
					128f,
					new FontSizeInfo(211f, 95f)
				},
				{
					256f,
					new FontSizeInfo(421f, 189f)
				}
			}
		},
		{
			"Haettenschweiler",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(20f, 7f)
				},
				{
					12f,
					new FontSizeInfo(21f, 7f)
				},
				{
					14f,
					new FontSizeInfo(24f, 8f)
				},
				{
					16f,
					new FontSizeInfo(26f, 9f)
				},
				{
					18f,
					new FontSizeInfo(29f, 11f)
				},
				{
					20f,
					new FontSizeInfo(32f, 12f)
				},
				{
					22f,
					new FontSizeInfo(34f, 13f)
				},
				{
					24f,
					new FontSizeInfo(37f, 14f)
				},
				{
					26f,
					new FontSizeInfo(40f, 15f)
				},
				{
					28f,
					new FontSizeInfo(43f, 16f)
				},
				{
					36f,
					new FontSizeInfo(55f, 21f)
				},
				{
					48f,
					new FontSizeInfo(73f, 28f)
				},
				{
					72f,
					new FontSizeInfo(0f, 42f)
				},
				{
					96f,
					new FontSizeInfo(148f, 56f)
				},
				{
					128f,
					new FontSizeInfo(198f, 75f)
				},
				{
					256f,
					new FontSizeInfo(391f, 150f)
				}
			}
		},
		{
			"High Tower Text",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(20f, 7f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 9f)
				},
				{
					16f,
					new FontSizeInfo(27f, 10f)
				},
				{
					18f,
					new FontSizeInfo(31f, 12f)
				},
				{
					20f,
					new FontSizeInfo(35f, 13f)
				},
				{
					22f,
					new FontSizeInfo(37f, 14f)
				},
				{
					24f,
					new FontSizeInfo(41f, 15f)
				},
				{
					26f,
					new FontSizeInfo(45f, 17f)
				},
				{
					28f,
					new FontSizeInfo(48f, 18f)
				},
				{
					36f,
					new FontSizeInfo(62f, 23f)
				},
				{
					48f,
					new FontSizeInfo(82f, 31f)
				},
				{
					72f,
					new FontSizeInfo(0f, 46f)
				},
				{
					96f,
					new FontSizeInfo(163f, 62f)
				},
				{
					128f,
					new FontSizeInfo(217f, 82f)
				},
				{
					256f,
					new FontSizeInfo(432f, 164f)
				}
			}
		},
		{
			"Imprint MT Shadow",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(27f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 12f)
				},
				{
					20f,
					new FontSizeInfo(35f, 14f)
				},
				{
					22f,
					new FontSizeInfo(37f, 15f)
				},
				{
					24f,
					new FontSizeInfo(41f, 16f)
				},
				{
					26f,
					new FontSizeInfo(45f, 18f)
				},
				{
					28f,
					new FontSizeInfo(47f, 19f)
				},
				{
					36f,
					new FontSizeInfo(61f, 24f)
				},
				{
					48f,
					new FontSizeInfo(81f, 33f)
				},
				{
					72f,
					new FontSizeInfo(0f, 49f)
				},
				{
					96f,
					new FontSizeInfo(161f, 65f)
				},
				{
					128f,
					new FontSizeInfo(215f, 87f)
				},
				{
					256f,
					new FontSizeInfo(427f, 174f)
				}
			}
		},
		{
			"Informal Roman",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(29f, 11f)
				},
				{
					18f,
					new FontSizeInfo(33f, 13f)
				},
				{
					20f,
					new FontSizeInfo(39f, 15f)
				},
				{
					22f,
					new FontSizeInfo(42f, 16f)
				},
				{
					24f,
					new FontSizeInfo(45f, 17f)
				},
				{
					26f,
					new FontSizeInfo(49f, 19f)
				},
				{
					28f,
					new FontSizeInfo(53f, 20f)
				},
				{
					36f,
					new FontSizeInfo(69f, 25f)
				},
				{
					48f,
					new FontSizeInfo(91f, 34f)
				},
				{
					72f,
					new FontSizeInfo(0f, 51f)
				},
				{
					96f,
					new FontSizeInfo(183f, 68f)
				},
				{
					128f,
					new FontSizeInfo(245f, 91f)
				},
				{
					256f,
					new FontSizeInfo(479f, 181f)
				}
			}
		},
		{
			"Blackadder ITC",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 4f)
				},
				{
					10f,
					new FontSizeInfo(21f, 5f)
				},
				{
					11f,
					new FontSizeInfo(22f, 6f)
				},
				{
					12f,
					new FontSizeInfo(23f, 6f)
				},
				{
					14f,
					new FontSizeInfo(28f, 8f)
				},
				{
					16f,
					new FontSizeInfo(30f, 8f)
				},
				{
					18f,
					new FontSizeInfo(34f, 9f)
				},
				{
					20f,
					new FontSizeInfo(38f, 11f)
				},
				{
					22f,
					new FontSizeInfo(41f, 11f)
				},
				{
					24f,
					new FontSizeInfo(45f, 13f)
				},
				{
					26f,
					new FontSizeInfo(50f, 14f)
				},
				{
					28f,
					new FontSizeInfo(52f, 15f)
				},
				{
					36f,
					new FontSizeInfo(68f, 19f)
				},
				{
					48f,
					new FontSizeInfo(90f, 25f)
				},
				{
					72f,
					new FontSizeInfo(0f, 38f)
				},
				{
					96f,
					new FontSizeInfo(179f, 51f)
				},
				{
					128f,
					new FontSizeInfo(238f, 67f)
				},
				{
					256f,
					new FontSizeInfo(474f, 135f)
				}
			}
		},
		{
			"Edwardian Script ITC",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(21f, 7f)
				},
				{
					12f,
					new FontSizeInfo(22f, 8f)
				},
				{
					14f,
					new FontSizeInfo(26f, 9f)
				},
				{
					16f,
					new FontSizeInfo(29f, 10f)
				},
				{
					18f,
					new FontSizeInfo(33f, 11f)
				},
				{
					20f,
					new FontSizeInfo(37f, 13f)
				},
				{
					22f,
					new FontSizeInfo(40f, 14f)
				},
				{
					24f,
					new FontSizeInfo(44f, 15f)
				},
				{
					26f,
					new FontSizeInfo(47f, 17f)
				},
				{
					28f,
					new FontSizeInfo(50f, 18f)
				},
				{
					36f,
					new FontSizeInfo(65f, 23f)
				},
				{
					48f,
					new FontSizeInfo(86f, 31f)
				},
				{
					72f,
					new FontSizeInfo(0f, 46f)
				},
				{
					96f,
					new FontSizeInfo(171f, 61f)
				},
				{
					128f,
					new FontSizeInfo(228f, 82f)
				},
				{
					256f,
					new FontSizeInfo(454f, 163f)
				}
			}
		},
		{
			"Kristen ITC",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 7f)
				},
				{
					10f,
					new FontSizeInfo(21f, 8f)
				},
				{
					11f,
					new FontSizeInfo(24f, 9f)
				},
				{
					12f,
					new FontSizeInfo(25f, 10f)
				},
				{
					14f,
					new FontSizeInfo(29f, 11f)
				},
				{
					16f,
					new FontSizeInfo(31f, 13f)
				},
				{
					18f,
					new FontSizeInfo(35f, 14f)
				},
				{
					20f,
					new FontSizeInfo(37f, 16f)
				},
				{
					22f,
					new FontSizeInfo(42f, 17f)
				},
				{
					24f,
					new FontSizeInfo(46f, 19f)
				},
				{
					26f,
					new FontSizeInfo(49f, 21f)
				},
				{
					28f,
					new FontSizeInfo(53f, 22f)
				},
				{
					36f,
					new FontSizeInfo(68f, 29f)
				},
				{
					48f,
					new FontSizeInfo(91f, 38f)
				},
				{
					72f,
					new FontSizeInfo(0f, 58f)
				},
				{
					96f,
					new FontSizeInfo(178f, 77f)
				},
				{
					128f,
					new FontSizeInfo(240f, 103f)
				},
				{
					256f,
					new FontSizeInfo(473f, 205f)
				}
			}
		},
		{
			"Jokerman",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(22f, 8f)
				},
				{
					10f,
					new FontSizeInfo(23f, 9f)
				},
				{
					11f,
					new FontSizeInfo(25f, 11f)
				},
				{
					12f,
					new FontSizeInfo(27f, 11f)
				},
				{
					14f,
					new FontSizeInfo(31f, 14f)
				},
				{
					16f,
					new FontSizeInfo(35f, 15f)
				},
				{
					18f,
					new FontSizeInfo(39f, 17f)
				},
				{
					20f,
					new FontSizeInfo(43f, 19f)
				},
				{
					22f,
					new FontSizeInfo(48f, 21f)
				},
				{
					24f,
					new FontSizeInfo(52f, 23f)
				},
				{
					26f,
					new FontSizeInfo(56f, 25f)
				},
				{
					28f,
					new FontSizeInfo(61f, 26f)
				},
				{
					36f,
					new FontSizeInfo(76f, 34f)
				},
				{
					48f,
					new FontSizeInfo(101f, 46f)
				},
				{
					72f,
					new FontSizeInfo(0f, 69f)
				},
				{
					96f,
					new FontSizeInfo(202f, 91f)
				},
				{
					128f,
					new FontSizeInfo(270f, 122f)
				},
				{
					256f,
					new FontSizeInfo(569f, 244f)
				}
			}
		},
		{
			"Juice ITC",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(20f, 5f)
				},
				{
					11f,
					new FontSizeInfo(20f, 6f)
				},
				{
					12f,
					new FontSizeInfo(21f, 6f)
				},
				{
					14f,
					new FontSizeInfo(24f, 7f)
				},
				{
					16f,
					new FontSizeInfo(27f, 8f)
				},
				{
					18f,
					new FontSizeInfo(30f, 8f)
				},
				{
					20f,
					new FontSizeInfo(34f, 10f)
				},
				{
					22f,
					new FontSizeInfo(36f, 11f)
				},
				{
					24f,
					new FontSizeInfo(40f, 11f)
				},
				{
					26f,
					new FontSizeInfo(44f, 13f)
				},
				{
					28f,
					new FontSizeInfo(46f, 14f)
				},
				{
					36f,
					new FontSizeInfo(62f, 18f)
				},
				{
					48f,
					new FontSizeInfo(81f, 24f)
				},
				{
					72f,
					new FontSizeInfo(0f, 36f)
				},
				{
					96f,
					new FontSizeInfo(164f, 48f)
				},
				{
					128f,
					new FontSizeInfo(218f, 64f)
				},
				{
					256f,
					new FontSizeInfo(436f, 128f)
				}
			}
		},
		{
			"Kunstler Script",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(20f, 5f)
				},
				{
					11f,
					new FontSizeInfo(20f, 5f)
				},
				{
					12f,
					new FontSizeInfo(21f, 6f)
				},
				{
					14f,
					new FontSizeInfo(25f, 7f)
				},
				{
					16f,
					new FontSizeInfo(28f, 7f)
				},
				{
					18f,
					new FontSizeInfo(31f, 8f)
				},
				{
					20f,
					new FontSizeInfo(35f, 10f)
				},
				{
					22f,
					new FontSizeInfo(38f, 10f)
				},
				{
					24f,
					new FontSizeInfo(42f, 11f)
				},
				{
					26f,
					new FontSizeInfo(45f, 12f)
				},
				{
					28f,
					new FontSizeInfo(48f, 13f)
				},
				{
					36f,
					new FontSizeInfo(62f, 17f)
				},
				{
					48f,
					new FontSizeInfo(82f, 23f)
				},
				{
					72f,
					new FontSizeInfo(0f, 34f)
				},
				{
					96f,
					new FontSizeInfo(163f, 45f)
				},
				{
					128f,
					new FontSizeInfo(218f, 60f)
				},
				{
					256f,
					new FontSizeInfo(437f, 120f)
				}
			}
		},
		{
			"Wide Latin",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 14f)
				},
				{
					10f,
					new FontSizeInfo(20f, 16f)
				},
				{
					11f,
					new FontSizeInfo(20f, 19f)
				},
				{
					12f,
					new FontSizeInfo(21f, 20f)
				},
				{
					14f,
					new FontSizeInfo(25f, 24f)
				},
				{
					16f,
					new FontSizeInfo(28f, 26f)
				},
				{
					18f,
					new FontSizeInfo(31f, 30f)
				},
				{
					20f,
					new FontSizeInfo(35f, 33f)
				},
				{
					22f,
					new FontSizeInfo(38f, 36f)
				},
				{
					24f,
					new FontSizeInfo(42f, 40f)
				},
				{
					26f,
					new FontSizeInfo(45f, 43f)
				},
				{
					28f,
					new FontSizeInfo(48f, 46f)
				},
				{
					36f,
					new FontSizeInfo(62f, 59f)
				},
				{
					48f,
					new FontSizeInfo(82f, 79f)
				},
				{
					72f,
					new FontSizeInfo(0f, 119f)
				},
				{
					96f,
					new FontSizeInfo(163f, 158f)
				},
				{
					128f,
					new FontSizeInfo(218f, 212f)
				},
				{
					256f,
					new FontSizeInfo(433f, 422f)
				}
			}
		},
		{
			"Lucida Bright",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 12f)
				},
				{
					16f,
					new FontSizeInfo(26f, 13f)
				},
				{
					18f,
					new FontSizeInfo(30f, 15f)
				},
				{
					20f,
					new FontSizeInfo(34f, 16f)
				},
				{
					22f,
					new FontSizeInfo(36f, 18f)
				},
				{
					24f,
					new FontSizeInfo(39f, 19f)
				},
				{
					26f,
					new FontSizeInfo(43f, 21f)
				},
				{
					28f,
					new FontSizeInfo(45f, 23f)
				},
				{
					36f,
					new FontSizeInfo(59f, 29f)
				},
				{
					48f,
					new FontSizeInfo(80f, 39f)
				},
				{
					72f,
					new FontSizeInfo(0f, 58f)
				},
				{
					96f,
					new FontSizeInfo(159f, 78f)
				},
				{
					128f,
					new FontSizeInfo(213f, 104f)
				},
				{
					256f,
					new FontSizeInfo(426f, 207f)
				}
			}
		},
		{
			"Lucida Calligraphy",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 8f)
				},
				{
					10f,
					new FontSizeInfo(20f, 10f)
				},
				{
					11f,
					new FontSizeInfo(23f, 11f)
				},
				{
					12f,
					new FontSizeInfo(24f, 12f)
				},
				{
					14f,
					new FontSizeInfo(26f, 14f)
				},
				{
					16f,
					new FontSizeInfo(29f, 15f)
				},
				{
					18f,
					new FontSizeInfo(35f, 17f)
				},
				{
					20f,
					new FontSizeInfo(39f, 19f)
				},
				{
					22f,
					new FontSizeInfo(41f, 20f)
				},
				{
					24f,
					new FontSizeInfo(45f, 22f)
				},
				{
					26f,
					new FontSizeInfo(49f, 24f)
				},
				{
					28f,
					new FontSizeInfo(52f, 26f)
				},
				{
					36f,
					new FontSizeInfo(69f, 33f)
				},
				{
					48f,
					new FontSizeInfo(92f, 44f)
				},
				{
					72f,
					new FontSizeInfo(0f, 65f)
				},
				{
					96f,
					new FontSizeInfo(183f, 86f)
				},
				{
					128f,
					new FontSizeInfo(246f, 115f)
				},
				{
					256f,
					new FontSizeInfo(489f, 227f)
				}
			}
		},
		{
			"Leelawadee",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 11f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 14f)
				},
				{
					20f,
					new FontSizeInfo(34f, 16f)
				},
				{
					22f,
					new FontSizeInfo(37f, 17f)
				},
				{
					24f,
					new FontSizeInfo(41f, 18f)
				},
				{
					26f,
					new FontSizeInfo(44f, 20f)
				},
				{
					28f,
					new FontSizeInfo(47f, 21f)
				},
				{
					36f,
					new FontSizeInfo(60f, 28f)
				},
				{
					48f,
					new FontSizeInfo(80f, 37f)
				},
				{
					72f,
					new FontSizeInfo(0f, 55f)
				},
				{
					96f,
					new FontSizeInfo(160f, 74f)
				},
				{
					128f,
					new FontSizeInfo(213f, 98f)
				},
				{
					256f,
					new FontSizeInfo(423f, 196f)
				}
			}
		},
		{
			"Lucida Fax",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 9f)
				},
				{
					11f,
					new FontSizeInfo(20f, 10f)
				},
				{
					12f,
					new FontSizeInfo(21f, 11f)
				},
				{
					14f,
					new FontSizeInfo(24f, 13f)
				},
				{
					16f,
					new FontSizeInfo(27f, 14f)
				},
				{
					18f,
					new FontSizeInfo(30f, 16f)
				},
				{
					20f,
					new FontSizeInfo(34f, 18f)
				},
				{
					22f,
					new FontSizeInfo(36f, 19f)
				},
				{
					24f,
					new FontSizeInfo(40f, 21f)
				},
				{
					26f,
					new FontSizeInfo(44f, 23f)
				},
				{
					28f,
					new FontSizeInfo(46f, 25f)
				},
				{
					36f,
					new FontSizeInfo(59f, 32f)
				},
				{
					48f,
					new FontSizeInfo(78f, 42f)
				},
				{
					72f,
					new FontSizeInfo(0f, 64f)
				},
				{
					96f,
					new FontSizeInfo(157f, 85f)
				},
				{
					128f,
					new FontSizeInfo(211f, 113f)
				},
				{
					256f,
					new FontSizeInfo(422f, 226f)
				}
			}
		},
		{
			"Lucida Handwriting",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 8f)
				},
				{
					10f,
					new FontSizeInfo(20f, 10f)
				},
				{
					11f,
					new FontSizeInfo(21f, 11f)
				},
				{
					12f,
					new FontSizeInfo(24f, 12f)
				},
				{
					14f,
					new FontSizeInfo(26f, 14f)
				},
				{
					16f,
					new FontSizeInfo(29f, 15f)
				},
				{
					18f,
					new FontSizeInfo(33f, 17f)
				},
				{
					20f,
					new FontSizeInfo(39f, 19f)
				},
				{
					22f,
					new FontSizeInfo(41f, 21f)
				},
				{
					24f,
					new FontSizeInfo(45f, 23f)
				},
				{
					26f,
					new FontSizeInfo(49f, 25f)
				},
				{
					28f,
					new FontSizeInfo(52f, 26f)
				},
				{
					36f,
					new FontSizeInfo(69f, 33f)
				},
				{
					48f,
					new FontSizeInfo(90f, 44f)
				},
				{
					72f,
					new FontSizeInfo(0f, 66f)
				},
				{
					96f,
					new FontSizeInfo(181f, 87f)
				},
				{
					128f,
					new FontSizeInfo(242f, 116f)
				},
				{
					256f,
					new FontSizeInfo(483f, 231f)
				}
			}
		},
		{
			"Lucida Sans",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 9f)
				},
				{
					11f,
					new FontSizeInfo(20f, 10f)
				},
				{
					12f,
					new FontSizeInfo(21f, 11f)
				},
				{
					14f,
					new FontSizeInfo(24f, 13f)
				},
				{
					16f,
					new FontSizeInfo(27f, 14f)
				},
				{
					18f,
					new FontSizeInfo(30f, 16f)
				},
				{
					20f,
					new FontSizeInfo(34f, 18f)
				},
				{
					22f,
					new FontSizeInfo(36f, 19f)
				},
				{
					24f,
					new FontSizeInfo(40f, 21f)
				},
				{
					26f,
					new FontSizeInfo(43f, 23f)
				},
				{
					28f,
					new FontSizeInfo(46f, 24f)
				},
				{
					36f,
					new FontSizeInfo(59f, 32f)
				},
				{
					48f,
					new FontSizeInfo(80f, 42f)
				},
				{
					72f,
					new FontSizeInfo(0f, 63f)
				},
				{
					96f,
					new FontSizeInfo(159f, 84f)
				},
				{
					128f,
					new FontSizeInfo(213f, 113f)
				},
				{
					256f,
					new FontSizeInfo(426f, 225f)
				}
			}
		},
		{
			"Lucida Sans Typewriter",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(26f, 13f)
				},
				{
					18f,
					new FontSizeInfo(30f, 14f)
				},
				{
					20f,
					new FontSizeInfo(34f, 16f)
				},
				{
					22f,
					new FontSizeInfo(36f, 17f)
				},
				{
					24f,
					new FontSizeInfo(40f, 19f)
				},
				{
					26f,
					new FontSizeInfo(43f, 21f)
				},
				{
					28f,
					new FontSizeInfo(46f, 22f)
				},
				{
					36f,
					new FontSizeInfo(59f, 29f)
				},
				{
					48f,
					new FontSizeInfo(80f, 39f)
				},
				{
					72f,
					new FontSizeInfo(0f, 58f)
				},
				{
					96f,
					new FontSizeInfo(159f, 77f)
				},
				{
					128f,
					new FontSizeInfo(213f, 103f)
				},
				{
					256f,
					new FontSizeInfo(424f, 205f)
				}
			}
		},
		{
			"Magneto",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 8f)
				},
				{
					10f,
					new FontSizeInfo(20f, 10f)
				},
				{
					11f,
					new FontSizeInfo(21f, 12f)
				},
				{
					12f,
					new FontSizeInfo(22f, 12f)
				},
				{
					14f,
					new FontSizeInfo(26f, 15f)
				},
				{
					16f,
					new FontSizeInfo(28f, 16f)
				},
				{
					18f,
					new FontSizeInfo(32f, 19f)
				},
				{
					20f,
					new FontSizeInfo(36f, 21f)
				},
				{
					22f,
					new FontSizeInfo(39f, 22f)
				},
				{
					24f,
					new FontSizeInfo(43f, 25f)
				},
				{
					26f,
					new FontSizeInfo(47f, 27f)
				},
				{
					28f,
					new FontSizeInfo(49f, 29f)
				},
				{
					36f,
					new FontSizeInfo(64f, 37f)
				},
				{
					48f,
					new FontSizeInfo(84f, 49f)
				},
				{
					72f,
					new FontSizeInfo(0f, 74f)
				},
				{
					96f,
					new FontSizeInfo(168f, 99f)
				},
				{
					128f,
					new FontSizeInfo(224f, 132f)
				},
				{
					256f,
					new FontSizeInfo(445f, 263f)
				}
			}
		},
		{
			"Maiandra GD",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 9f)
				},
				{
					11f,
					new FontSizeInfo(20f, 10f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(24f, 12f)
				},
				{
					16f,
					new FontSizeInfo(27f, 13f)
				},
				{
					18f,
					new FontSizeInfo(31f, 15f)
				},
				{
					20f,
					new FontSizeInfo(34f, 17f)
				},
				{
					22f,
					new FontSizeInfo(37f, 18f)
				},
				{
					24f,
					new FontSizeInfo(41f, 20f)
				},
				{
					26f,
					new FontSizeInfo(44f, 22f)
				},
				{
					28f,
					new FontSizeInfo(47f, 23f)
				},
				{
					36f,
					new FontSizeInfo(60f, 29f)
				},
				{
					48f,
					new FontSizeInfo(80f, 39f)
				},
				{
					72f,
					new FontSizeInfo(0f, 57f)
				},
				{
					96f,
					new FontSizeInfo(159f, 76f)
				},
				{
					128f,
					new FontSizeInfo(212f, 102f)
				},
				{
					256f,
					new FontSizeInfo(422f, 200f)
				}
			}
		},
		{
			"Matura MT Script Capitals",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(21f, 8f)
				},
				{
					10f,
					new FontSizeInfo(22f, 10f)
				},
				{
					11f,
					new FontSizeInfo(23f, 11f)
				},
				{
					12f,
					new FontSizeInfo(25f, 12f)
				},
				{
					14f,
					new FontSizeInfo(29f, 13f)
				},
				{
					16f,
					new FontSizeInfo(32f, 15f)
				},
				{
					18f,
					new FontSizeInfo(36f, 17f)
				},
				{
					20f,
					new FontSizeInfo(41f, 19f)
				},
				{
					22f,
					new FontSizeInfo(44f, 20f)
				},
				{
					24f,
					new FontSizeInfo(48f, 22f)
				},
				{
					26f,
					new FontSizeInfo(53f, 24f)
				},
				{
					28f,
					new FontSizeInfo(56f, 25f)
				},
				{
					36f,
					new FontSizeInfo(72f, 33f)
				},
				{
					48f,
					new FontSizeInfo(95f, 42f)
				},
				{
					72f,
					new FontSizeInfo(0f, 63f)
				},
				{
					96f,
					new FontSizeInfo(190f, 84f)
				},
				{
					128f,
					new FontSizeInfo(253f, 111f)
				},
				{
					256f,
					new FontSizeInfo(504f, 221f)
				}
			}
		},
		{
			"Mistral",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(21f, 8f)
				},
				{
					12f,
					new FontSizeInfo(22f, 8f)
				},
				{
					14f,
					new FontSizeInfo(26f, 10f)
				},
				{
					16f,
					new FontSizeInfo(29f, 10f)
				},
				{
					18f,
					new FontSizeInfo(33f, 12f)
				},
				{
					20f,
					new FontSizeInfo(37f, 13f)
				},
				{
					22f,
					new FontSizeInfo(39f, 14f)
				},
				{
					24f,
					new FontSizeInfo(43f, 15f)
				},
				{
					26f,
					new FontSizeInfo(47f, 17f)
				},
				{
					28f,
					new FontSizeInfo(50f, 18f)
				},
				{
					36f,
					new FontSizeInfo(64f, 23f)
				},
				{
					48f,
					new FontSizeInfo(85f, 30f)
				},
				{
					72f,
					new FontSizeInfo(0f, 44f)
				},
				{
					96f,
					new FontSizeInfo(169f, 59f)
				},
				{
					128f,
					new FontSizeInfo(226f, 78f)
				},
				{
					256f,
					new FontSizeInfo(450f, 154f)
				}
			}
		},
		{
			"Modern No. 20",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(24f, 10f)
				},
				{
					16f,
					new FontSizeInfo(27f, 11f)
				},
				{
					18f,
					new FontSizeInfo(30f, 12f)
				},
				{
					20f,
					new FontSizeInfo(34f, 13f)
				},
				{
					22f,
					new FontSizeInfo(36f, 14f)
				},
				{
					24f,
					new FontSizeInfo(40f, 16f)
				},
				{
					26f,
					new FontSizeInfo(43f, 17f)
				},
				{
					28f,
					new FontSizeInfo(46f, 18f)
				},
				{
					36f,
					new FontSizeInfo(59f, 23f)
				},
				{
					48f,
					new FontSizeInfo(80f, 30f)
				},
				{
					72f,
					new FontSizeInfo(0f, 45f)
				},
				{
					96f,
					new FontSizeInfo(160f, 59f)
				},
				{
					128f,
					new FontSizeInfo(214f, 79f)
				},
				{
					256f,
					new FontSizeInfo(427f, 156f)
				}
			}
		},
		{
			"Microsoft Uighur",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(23f, 4f)
				},
				{
					10f,
					new FontSizeInfo(23f, 5f)
				},
				{
					11f,
					new FontSizeInfo(24f, 6f)
				},
				{
					12f,
					new FontSizeInfo(25f, 6f)
				},
				{
					14f,
					new FontSizeInfo(30f, 7f)
				},
				{
					16f,
					new FontSizeInfo(33f, 8f)
				},
				{
					18f,
					new FontSizeInfo(37f, 9f)
				},
				{
					20f,
					new FontSizeInfo(42f, 10f)
				},
				{
					22f,
					new FontSizeInfo(45f, 11f)
				},
				{
					24f,
					new FontSizeInfo(52f, 12f)
				},
				{
					26f,
					new FontSizeInfo(56f, 13f)
				},
				{
					28f,
					new FontSizeInfo(59f, 14f)
				},
				{
					36f,
					new FontSizeInfo(76f, 18f)
				},
				{
					48f,
					new FontSizeInfo(102f, 24f)
				},
				{
					72f,
					new FontSizeInfo(0f, 36f)
				},
				{
					96f,
					new FontSizeInfo(202f, 48f)
				},
				{
					128f,
					new FontSizeInfo(272f, 64f)
				},
				{
					256f,
					new FontSizeInfo(476f, 128f)
				}
			}
		},
		{
			"Monotype Corsiva",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 9f)
				},
				{
					16f,
					new FontSizeInfo(28f, 10f)
				},
				{
					18f,
					new FontSizeInfo(31f, 12f)
				},
				{
					20f,
					new FontSizeInfo(36f, 13f)
				},
				{
					22f,
					new FontSizeInfo(39f, 14f)
				},
				{
					24f,
					new FontSizeInfo(42f, 15f)
				},
				{
					26f,
					new FontSizeInfo(47f, 16f)
				},
				{
					28f,
					new FontSizeInfo(49f, 17f)
				},
				{
					36f,
					new FontSizeInfo(62f, 22f)
				},
				{
					48f,
					new FontSizeInfo(85f, 29f)
				},
				{
					72f,
					new FontSizeInfo(0f, 43f)
				},
				{
					96f,
					new FontSizeInfo(170f, 57f)
				},
				{
					128f,
					new FontSizeInfo(225f, 76f)
				},
				{
					256f,
					new FontSizeInfo(453f, 151f)
				}
			}
		},
		{
			"Niagara Engraved",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(21f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 5f)
				},
				{
					11f,
					new FontSizeInfo(20f, 5f)
				},
				{
					12f,
					new FontSizeInfo(21f, 6f)
				},
				{
					14f,
					new FontSizeInfo(24f, 7f)
				},
				{
					16f,
					new FontSizeInfo(26f, 8f)
				},
				{
					18f,
					new FontSizeInfo(29f, 9f)
				},
				{
					20f,
					new FontSizeInfo(33f, 9f)
				},
				{
					22f,
					new FontSizeInfo(35f, 10f)
				},
				{
					24f,
					new FontSizeInfo(39f, 11f)
				},
				{
					26f,
					new FontSizeInfo(43f, 11f)
				},
				{
					28f,
					new FontSizeInfo(45f, 12f)
				},
				{
					36f,
					new FontSizeInfo(58f, 15f)
				},
				{
					48f,
					new FontSizeInfo(77f, 20f)
				},
				{
					72f,
					new FontSizeInfo(0f, 30f)
				},
				{
					96f,
					new FontSizeInfo(154f, 39f)
				},
				{
					128f,
					new FontSizeInfo(207f, 52f)
				},
				{
					256f,
					new FontSizeInfo(407f, 103f)
				}
			}
		},
		{
			"Niagara Solid",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(21f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 5f)
				},
				{
					11f,
					new FontSizeInfo(20f, 6f)
				},
				{
					12f,
					new FontSizeInfo(21f, 6f)
				},
				{
					14f,
					new FontSizeInfo(24f, 7f)
				},
				{
					16f,
					new FontSizeInfo(26f, 8f)
				},
				{
					18f,
					new FontSizeInfo(29f, 9f)
				},
				{
					20f,
					new FontSizeInfo(33f, 9f)
				},
				{
					22f,
					new FontSizeInfo(35f, 10f)
				},
				{
					24f,
					new FontSizeInfo(39f, 11f)
				},
				{
					26f,
					new FontSizeInfo(43f, 11f)
				},
				{
					28f,
					new FontSizeInfo(45f, 12f)
				},
				{
					36f,
					new FontSizeInfo(58f, 15f)
				},
				{
					48f,
					new FontSizeInfo(77f, 20f)
				},
				{
					72f,
					new FontSizeInfo(0f, 30f)
				},
				{
					96f,
					new FontSizeInfo(154f, 39f)
				},
				{
					128f,
					new FontSizeInfo(207f, 52f)
				},
				{
					256f,
					new FontSizeInfo(407f, 103f)
				}
			}
		},
		{
			"OCR A Extended",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 8f)
				},
				{
					10f,
					new FontSizeInfo(20f, 9f)
				},
				{
					11f,
					new FontSizeInfo(20f, 10f)
				},
				{
					12f,
					new FontSizeInfo(21f, 11f)
				},
				{
					14f,
					new FontSizeInfo(24f, 12f)
				},
				{
					16f,
					new FontSizeInfo(26f, 14f)
				},
				{
					18f,
					new FontSizeInfo(29f, 15f)
				},
				{
					20f,
					new FontSizeInfo(33f, 17f)
				},
				{
					22f,
					new FontSizeInfo(35f, 19f)
				},
				{
					24f,
					new FontSizeInfo(39f, 20f)
				},
				{
					26f,
					new FontSizeInfo(42f, 22f)
				},
				{
					28f,
					new FontSizeInfo(45f, 23f)
				},
				{
					36f,
					new FontSizeInfo(57f, 30f)
				},
				{
					48f,
					new FontSizeInfo(76f, 40f)
				},
				{
					72f,
					new FontSizeInfo(0f, 59f)
				},
				{
					96f,
					new FontSizeInfo(152f, 78f)
				},
				{
					128f,
					new FontSizeInfo(202f, 104f)
				},
				{
					256f,
					new FontSizeInfo(402f, 207f)
				}
			}
		},
		{
			"Old English Text MT",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(21f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(30f, 13f)
				},
				{
					20f,
					new FontSizeInfo(34f, 15f)
				},
				{
					22f,
					new FontSizeInfo(36f, 16f)
				},
				{
					24f,
					new FontSizeInfo(40f, 17f)
				},
				{
					26f,
					new FontSizeInfo(44f, 19f)
				},
				{
					28f,
					new FontSizeInfo(46f, 20f)
				},
				{
					36f,
					new FontSizeInfo(59f, 25f)
				},
				{
					48f,
					new FontSizeInfo(79f, 34f)
				},
				{
					72f,
					new FontSizeInfo(0f, 50f)
				},
				{
					96f,
					new FontSizeInfo(157f, 66f)
				},
				{
					128f,
					new FontSizeInfo(209f, 88f)
				},
				{
					256f,
					new FontSizeInfo(418f, 175f)
				}
			}
		},
		{
			"Onyx",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(20f, 6f)
				},
				{
					12f,
					new FontSizeInfo(21f, 7f)
				},
				{
					14f,
					new FontSizeInfo(24f, 7f)
				},
				{
					16f,
					new FontSizeInfo(26f, 8f)
				},
				{
					18f,
					new FontSizeInfo(29f, 9f)
				},
				{
					20f,
					new FontSizeInfo(33f, 9f)
				},
				{
					22f,
					new FontSizeInfo(35f, 10f)
				},
				{
					24f,
					new FontSizeInfo(39f, 10f)
				},
				{
					26f,
					new FontSizeInfo(42f, 11f)
				},
				{
					28f,
					new FontSizeInfo(45f, 11f)
				},
				{
					36f,
					new FontSizeInfo(58f, 15f)
				},
				{
					48f,
					new FontSizeInfo(77f, 19f)
				},
				{
					72f,
					new FontSizeInfo(0f, 29f)
				},
				{
					96f,
					new FontSizeInfo(152f, 39f)
				},
				{
					128f,
					new FontSizeInfo(203f, 50f)
				},
				{
					256f,
					new FontSizeInfo(416f, 97f)
				}
			}
		},
		{
			"MS Outlook",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(26f, 12f)
				},
				{
					18f,
					new FontSizeInfo(29f, 13f)
				},
				{
					20f,
					new FontSizeInfo(34f, 15f)
				},
				{
					22f,
					new FontSizeInfo(36f, 16f)
				},
				{
					24f,
					new FontSizeInfo(40f, 17f)
				},
				{
					26f,
					new FontSizeInfo(43f, 19f)
				},
				{
					28f,
					new FontSizeInfo(46f, 20f)
				},
				{
					36f,
					new FontSizeInfo(59f, 25f)
				},
				{
					48f,
					new FontSizeInfo(78f, 33f)
				},
				{
					72f,
					new FontSizeInfo(0f, 49f)
				},
				{
					96f,
					new FontSizeInfo(154f, 65f)
				},
				{
					128f,
					new FontSizeInfo(206f, 87f)
				},
				{
					256f,
					new FontSizeInfo(410f, 172f)
				}
			}
		},
		{
			"Palace Script MT",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(20f, 5f)
				},
				{
					11f,
					new FontSizeInfo(20f, 6f)
				},
				{
					12f,
					new FontSizeInfo(21f, 6f)
				},
				{
					14f,
					new FontSizeInfo(25f, 7f)
				},
				{
					16f,
					new FontSizeInfo(27f, 8f)
				},
				{
					18f,
					new FontSizeInfo(31f, 9f)
				},
				{
					20f,
					new FontSizeInfo(35f, 9f)
				},
				{
					22f,
					new FontSizeInfo(37f, 10f)
				},
				{
					24f,
					new FontSizeInfo(41f, 11f)
				},
				{
					26f,
					new FontSizeInfo(44f, 12f)
				},
				{
					28f,
					new FontSizeInfo(47f, 13f)
				},
				{
					36f,
					new FontSizeInfo(61f, 16f)
				},
				{
					48f,
					new FontSizeInfo(81f, 21f)
				},
				{
					72f,
					new FontSizeInfo(0f, 31f)
				},
				{
					96f,
					new FontSizeInfo(160f, 41f)
				},
				{
					128f,
					new FontSizeInfo(213f, 55f)
				},
				{
					256f,
					new FontSizeInfo(425f, 108f)
				}
			}
		},
		{
			"Papyrus",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(23f, 7f)
				},
				{
					10f,
					new FontSizeInfo(24f, 9f)
				},
				{
					11f,
					new FontSizeInfo(26f, 10f)
				},
				{
					12f,
					new FontSizeInfo(27f, 10f)
				},
				{
					14f,
					new FontSizeInfo(32f, 12f)
				},
				{
					16f,
					new FontSizeInfo(35f, 14f)
				},
				{
					18f,
					new FontSizeInfo(40f, 15f)
				},
				{
					20f,
					new FontSizeInfo(44f, 16f)
				},
				{
					22f,
					new FontSizeInfo(49f, 18f)
				},
				{
					24f,
					new FontSizeInfo(54f, 20f)
				},
				{
					26f,
					new FontSizeInfo(58f, 22f)
				},
				{
					28f,
					new FontSizeInfo(61f, 22f)
				},
				{
					36f,
					new FontSizeInfo(80f, 29f)
				},
				{
					48f,
					new FontSizeInfo(104f, 38f)
				},
				{
					72f,
					new FontSizeInfo(0f, 56f)
				},
				{
					96f,
					new FontSizeInfo(210f, 74f)
				},
				{
					128f,
					new FontSizeInfo(279f, 99f)
				},
				{
					256f,
					new FontSizeInfo(556f, 197f)
				}
			}
		},
		{
			"Parchment",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 6f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(21f, 4f)
				},
				{
					11f,
					new FontSizeInfo(22f, 4f)
				},
				{
					12f,
					new FontSizeInfo(23f, 4f)
				},
				{
					14f,
					new FontSizeInfo(27f, 5f)
				},
				{
					16f,
					new FontSizeInfo(29f, 5f)
				},
				{
					18f,
					new FontSizeInfo(33f, 6f)
				},
				{
					20f,
					new FontSizeInfo(36f, 7f)
				},
				{
					22f,
					new FontSizeInfo(39f, 7f)
				},
				{
					24f,
					new FontSizeInfo(43f, 8f)
				},
				{
					26f,
					new FontSizeInfo(46f, 8f)
				},
				{
					28f,
					new FontSizeInfo(49f, 9f)
				},
				{
					36f,
					new FontSizeInfo(64f, 11f)
				},
				{
					48f,
					new FontSizeInfo(85f, 14f)
				},
				{
					72f,
					new FontSizeInfo(0f, 21f)
				},
				{
					96f,
					new FontSizeInfo(169f, 28f)
				},
				{
					128f,
					new FontSizeInfo(226f, 37f)
				},
				{
					256f,
					new FontSizeInfo(450f, 73f)
				}
			}
		},
		{
			"Perpetua",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(21f, 7f)
				},
				{
					12f,
					new FontSizeInfo(22f, 7f)
				},
				{
					14f,
					new FontSizeInfo(26f, 9f)
				},
				{
					16f,
					new FontSizeInfo(29f, 10f)
				},
				{
					18f,
					new FontSizeInfo(33f, 11f)
				},
				{
					20f,
					new FontSizeInfo(37f, 12f)
				},
				{
					22f,
					new FontSizeInfo(39f, 13f)
				},
				{
					24f,
					new FontSizeInfo(43f, 15f)
				},
				{
					26f,
					new FontSizeInfo(47f, 16f)
				},
				{
					28f,
					new FontSizeInfo(50f, 17f)
				},
				{
					36f,
					new FontSizeInfo(65f, 22f)
				},
				{
					48f,
					new FontSizeInfo(86f, 29f)
				},
				{
					72f,
					new FontSizeInfo(0f, 44f)
				},
				{
					96f,
					new FontSizeInfo(171f, 59f)
				},
				{
					128f,
					new FontSizeInfo(228f, 78f)
				},
				{
					256f,
					new FontSizeInfo(453f, 156f)
				}
			}
		},
		{
			"Perpetua Titling MT",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 9f)
				},
				{
					11f,
					new FontSizeInfo(20f, 10f)
				},
				{
					12f,
					new FontSizeInfo(21f, 11f)
				},
				{
					14f,
					new FontSizeInfo(24f, 13f)
				},
				{
					16f,
					new FontSizeInfo(27f, 14f)
				},
				{
					18f,
					new FontSizeInfo(30f, 16f)
				},
				{
					20f,
					new FontSizeInfo(34f, 18f)
				},
				{
					22f,
					new FontSizeInfo(36f, 19f)
				},
				{
					24f,
					new FontSizeInfo(40f, 21f)
				},
				{
					26f,
					new FontSizeInfo(44f, 23f)
				},
				{
					28f,
					new FontSizeInfo(46f, 25f)
				},
				{
					36f,
					new FontSizeInfo(60f, 32f)
				},
				{
					48f,
					new FontSizeInfo(79f, 43f)
				},
				{
					72f,
					new FontSizeInfo(0f, 64f)
				},
				{
					96f,
					new FontSizeInfo(158f, 85f)
				},
				{
					128f,
					new FontSizeInfo(210f, 114f)
				},
				{
					256f,
					new FontSizeInfo(418f, 227f)
				}
			}
		},
		{
			"Playbill",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(20f, 5f)
				},
				{
					11f,
					new FontSizeInfo(20f, 5f)
				},
				{
					12f,
					new FontSizeInfo(21f, 6f)
				},
				{
					14f,
					new FontSizeInfo(24f, 7f)
				},
				{
					16f,
					new FontSizeInfo(26f, 7f)
				},
				{
					18f,
					new FontSizeInfo(29f, 9f)
				},
				{
					20f,
					new FontSizeInfo(33f, 10f)
				},
				{
					22f,
					new FontSizeInfo(35f, 10f)
				},
				{
					24f,
					new FontSizeInfo(39f, 11f)
				},
				{
					26f,
					new FontSizeInfo(42f, 12f)
				},
				{
					28f,
					new FontSizeInfo(45f, 13f)
				},
				{
					36f,
					new FontSizeInfo(58f, 17f)
				},
				{
					48f,
					new FontSizeInfo(77f, 23f)
				},
				{
					72f,
					new FontSizeInfo(0f, 34f)
				},
				{
					96f,
					new FontSizeInfo(153f, 46f)
				},
				{
					128f,
					new FontSizeInfo(203f, 61f)
				},
				{
					256f,
					new FontSizeInfo(405f, 122f)
				}
			}
		},
		{
			"Poor Richard",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(30f, 13f)
				},
				{
					20f,
					new FontSizeInfo(34f, 15f)
				},
				{
					22f,
					new FontSizeInfo(37f, 16f)
				},
				{
					24f,
					new FontSizeInfo(40f, 18f)
				},
				{
					26f,
					new FontSizeInfo(45f, 19f)
				},
				{
					28f,
					new FontSizeInfo(48f, 20f)
				},
				{
					36f,
					new FontSizeInfo(60f, 26f)
				},
				{
					48f,
					new FontSizeInfo(79f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 53f)
				},
				{
					96f,
					new FontSizeInfo(157f, 70f)
				},
				{
					128f,
					new FontSizeInfo(209f, 94f)
				},
				{
					256f,
					new FontSizeInfo(416f, 187f)
				}
			}
		},
		{
			"Pristina",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 5f)
				},
				{
					10f,
					new FontSizeInfo(22f, 7f)
				},
				{
					11f,
					new FontSizeInfo(23f, 8f)
				},
				{
					12f,
					new FontSizeInfo(24f, 8f)
				},
				{
					14f,
					new FontSizeInfo(29f, 9f)
				},
				{
					16f,
					new FontSizeInfo(32f, 10f)
				},
				{
					18f,
					new FontSizeInfo(36f, 12f)
				},
				{
					20f,
					new FontSizeInfo(40f, 13f)
				},
				{
					22f,
					new FontSizeInfo(43f, 14f)
				},
				{
					24f,
					new FontSizeInfo(47f, 15f)
				},
				{
					26f,
					new FontSizeInfo(52f, 16f)
				},
				{
					28f,
					new FontSizeInfo(55f, 17f)
				},
				{
					36f,
					new FontSizeInfo(71f, 23f)
				},
				{
					48f,
					new FontSizeInfo(94f, 30f)
				},
				{
					72f,
					new FontSizeInfo(0f, 45f)
				},
				{
					96f,
					new FontSizeInfo(187f, 61f)
				},
				{
					128f,
					new FontSizeInfo(248f, 81f)
				},
				{
					256f,
					new FontSizeInfo(494f, 162f)
				}
			}
		},
		{
			"Rage Italic",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(21f, 6f)
				},
				{
					10f,
					new FontSizeInfo(21f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(24f, 8f)
				},
				{
					14f,
					new FontSizeInfo(28f, 10f)
				},
				{
					16f,
					new FontSizeInfo(31f, 11f)
				},
				{
					18f,
					new FontSizeInfo(35f, 12f)
				},
				{
					20f,
					new FontSizeInfo(39f, 14f)
				},
				{
					22f,
					new FontSizeInfo(42f, 15f)
				},
				{
					24f,
					new FontSizeInfo(46f, 17f)
				},
				{
					26f,
					new FontSizeInfo(50f, 18f)
				},
				{
					28f,
					new FontSizeInfo(53f, 19f)
				},
				{
					36f,
					new FontSizeInfo(69f, 25f)
				},
				{
					48f,
					new FontSizeInfo(91f, 33f)
				},
				{
					72f,
					new FontSizeInfo(0f, 50f)
				},
				{
					96f,
					new FontSizeInfo(182f, 66f)
				},
				{
					128f,
					new FontSizeInfo(242f, 89f)
				},
				{
					256f,
					new FontSizeInfo(482f, 177f)
				}
			}
		},
		{
			"Ravie",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(22f, 10f)
				},
				{
					10f,
					new FontSizeInfo(22f, 12f)
				},
				{
					11f,
					new FontSizeInfo(24f, 14f)
				},
				{
					12f,
					new FontSizeInfo(25f, 15f)
				},
				{
					14f,
					new FontSizeInfo(29f, 18f)
				},
				{
					16f,
					new FontSizeInfo(34f, 20f)
				},
				{
					18f,
					new FontSizeInfo(37f, 23f)
				},
				{
					20f,
					new FontSizeInfo(38f, 26f)
				},
				{
					22f,
					new FontSizeInfo(42f, 28f)
				},
				{
					24f,
					new FontSizeInfo(46f, 30f)
				},
				{
					26f,
					new FontSizeInfo(52f, 33f)
				},
				{
					28f,
					new FontSizeInfo(53f, 35f)
				},
				{
					36f,
					new FontSizeInfo(67f, 46f)
				},
				{
					48f,
					new FontSizeInfo(91f, 61f)
				},
				{
					72f,
					new FontSizeInfo(0f, 91f)
				},
				{
					96f,
					new FontSizeInfo(179f, 122f)
				},
				{
					128f,
					new FontSizeInfo(238f, 162f)
				},
				{
					256f,
					new FontSizeInfo(472f, 324f)
				}
			}
		},
		{
			"MS Reference Sans Serif",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 10f)
				},
				{
					12f,
					new FontSizeInfo(21f, 10f)
				},
				{
					14f,
					new FontSizeInfo(26f, 12f)
				},
				{
					16f,
					new FontSizeInfo(27f, 13f)
				},
				{
					18f,
					new FontSizeInfo(30f, 15f)
				},
				{
					20f,
					new FontSizeInfo(36f, 17f)
				},
				{
					22f,
					new FontSizeInfo(36f, 18f)
				},
				{
					24f,
					new FontSizeInfo(42f, 20f)
				},
				{
					26f,
					new FontSizeInfo(46f, 22f)
				},
				{
					28f,
					new FontSizeInfo(48f, 24f)
				},
				{
					36f,
					new FontSizeInfo(62f, 31f)
				},
				{
					48f,
					new FontSizeInfo(81f, 41f)
				},
				{
					72f,
					new FontSizeInfo(0f, 61f)
				},
				{
					96f,
					new FontSizeInfo(159f, 81f)
				},
				{
					128f,
					new FontSizeInfo(211f, 109f)
				},
				{
					256f,
					new FontSizeInfo(418f, 216f)
				}
			}
		},
		{
			"MS Reference Specialty",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 10f)
				},
				{
					8f,
					new FontSizeInfo(20f, 13f)
				},
				{
					10f,
					new FontSizeInfo(20f, 16f)
				},
				{
					11f,
					new FontSizeInfo(20f, 18f)
				},
				{
					12f,
					new FontSizeInfo(21f, 19f)
				},
				{
					14f,
					new FontSizeInfo(24f, 23f)
				},
				{
					16f,
					new FontSizeInfo(27f, 25f)
				},
				{
					18f,
					new FontSizeInfo(32f, 29f)
				},
				{
					20f,
					new FontSizeInfo(34f, 32f)
				},
				{
					22f,
					new FontSizeInfo(36f, 35f)
				},
				{
					24f,
					new FontSizeInfo(42f, 38f)
				},
				{
					26f,
					new FontSizeInfo(44f, 42f)
				},
				{
					28f,
					new FontSizeInfo(48f, 44f)
				},
				{
					36f,
					new FontSizeInfo(59f, 58f)
				},
				{
					48f,
					new FontSizeInfo(81f, 77f)
				},
				{
					72f,
					new FontSizeInfo(0f, 115f)
				},
				{
					96f,
					new FontSizeInfo(159f, 153f)
				},
				{
					128f,
					new FontSizeInfo(213f, 205f)
				},
				{
					256f,
					new FontSizeInfo(423f, 408f)
				}
			}
		},
		{
			"Rockwell Condensed",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(20f, 5f)
				},
				{
					11f,
					new FontSizeInfo(20f, 5f)
				},
				{
					12f,
					new FontSizeInfo(21f, 6f)
				},
				{
					14f,
					new FontSizeInfo(24f, 7f)
				},
				{
					16f,
					new FontSizeInfo(27f, 8f)
				},
				{
					18f,
					new FontSizeInfo(31f, 9f)
				},
				{
					20f,
					new FontSizeInfo(34f, 10f)
				},
				{
					22f,
					new FontSizeInfo(37f, 11f)
				},
				{
					24f,
					new FontSizeInfo(41f, 12f)
				},
				{
					26f,
					new FontSizeInfo(44f, 13f)
				},
				{
					28f,
					new FontSizeInfo(47f, 14f)
				},
				{
					36f,
					new FontSizeInfo(60f, 18f)
				},
				{
					48f,
					new FontSizeInfo(80f, 23f)
				},
				{
					72f,
					new FontSizeInfo(0f, 35f)
				},
				{
					96f,
					new FontSizeInfo(159f, 47f)
				},
				{
					128f,
					new FontSizeInfo(212f, 62f)
				},
				{
					256f,
					new FontSizeInfo(422f, 125f)
				}
			}
		},
		{
			"Rockwell",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(24f, 10f)
				},
				{
					16f,
					new FontSizeInfo(27f, 11f)
				},
				{
					18f,
					new FontSizeInfo(30f, 13f)
				},
				{
					20f,
					new FontSizeInfo(34f, 15f)
				},
				{
					22f,
					new FontSizeInfo(36f, 16f)
				},
				{
					24f,
					new FontSizeInfo(40f, 17f)
				},
				{
					26f,
					new FontSizeInfo(44f, 19f)
				},
				{
					28f,
					new FontSizeInfo(46f, 20f)
				},
				{
					36f,
					new FontSizeInfo(60f, 26f)
				},
				{
					48f,
					new FontSizeInfo(80f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 52f)
				},
				{
					96f,
					new FontSizeInfo(158f, 69f)
				},
				{
					128f,
					new FontSizeInfo(210f, 93f)
				},
				{
					256f,
					new FontSizeInfo(420f, 185f)
				}
			}
		},
		{
			"Rockwell Extra Bold",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 8f)
				},
				{
					10f,
					new FontSizeInfo(20f, 9f)
				},
				{
					11f,
					new FontSizeInfo(20f, 10f)
				},
				{
					12f,
					new FontSizeInfo(21f, 11f)
				},
				{
					14f,
					new FontSizeInfo(24f, 13f)
				},
				{
					16f,
					new FontSizeInfo(27f, 14f)
				},
				{
					18f,
					new FontSizeInfo(30f, 17f)
				},
				{
					20f,
					new FontSizeInfo(34f, 19f)
				},
				{
					22f,
					new FontSizeInfo(37f, 20f)
				},
				{
					24f,
					new FontSizeInfo(40f, 22f)
				},
				{
					26f,
					new FontSizeInfo(44f, 24f)
				},
				{
					28f,
					new FontSizeInfo(46f, 25f)
				},
				{
					36f,
					new FontSizeInfo(60f, 33f)
				},
				{
					48f,
					new FontSizeInfo(80f, 44f)
				},
				{
					72f,
					new FontSizeInfo(0f, 66f)
				},
				{
					96f,
					new FontSizeInfo(158f, 88f)
				},
				{
					128f,
					new FontSizeInfo(211f, 118f)
				},
				{
					256f,
					new FontSizeInfo(420f, 234f)
				}
			}
		},
		{
			"Script MT Bold",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(27f, 11f)
				},
				{
					18f,
					new FontSizeInfo(31f, 13f)
				},
				{
					20f,
					new FontSizeInfo(35f, 14f)
				},
				{
					22f,
					new FontSizeInfo(37f, 15f)
				},
				{
					24f,
					new FontSizeInfo(41f, 17f)
				},
				{
					26f,
					new FontSizeInfo(45f, 18f)
				},
				{
					28f,
					new FontSizeInfo(47f, 20f)
				},
				{
					36f,
					new FontSizeInfo(61f, 25f)
				},
				{
					48f,
					new FontSizeInfo(81f, 34f)
				},
				{
					72f,
					new FontSizeInfo(0f, 51f)
				},
				{
					96f,
					new FontSizeInfo(161f, 68f)
				},
				{
					128f,
					new FontSizeInfo(215f, 90f)
				},
				{
					256f,
					new FontSizeInfo(428f, 180f)
				}
			}
		},
		{
			"Showcard Gothic",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(22f, 9f)
				},
				{
					14f,
					new FontSizeInfo(24f, 11f)
				},
				{
					16f,
					new FontSizeInfo(30f, 12f)
				},
				{
					18f,
					new FontSizeInfo(32f, 14f)
				},
				{
					20f,
					new FontSizeInfo(35f, 16f)
				},
				{
					22f,
					new FontSizeInfo(38f, 17f)
				},
				{
					24f,
					new FontSizeInfo(42f, 19f)
				},
				{
					26f,
					new FontSizeInfo(47f, 21f)
				},
				{
					28f,
					new FontSizeInfo(48f, 22f)
				},
				{
					36f,
					new FontSizeInfo(63f, 28f)
				},
				{
					48f,
					new FontSizeInfo(82f, 38f)
				},
				{
					72f,
					new FontSizeInfo(0f, 57f)
				},
				{
					96f,
					new FontSizeInfo(165f, 76f)
				},
				{
					128f,
					new FontSizeInfo(219f, 101f)
				},
				{
					256f,
					new FontSizeInfo(435f, 200f)
				}
			}
		},
		{
			"Snap ITC",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 10f)
				},
				{
					10f,
					new FontSizeInfo(20f, 11f)
				},
				{
					11f,
					new FontSizeInfo(20f, 13f)
				},
				{
					12f,
					new FontSizeInfo(24f, 14f)
				},
				{
					14f,
					new FontSizeInfo(26f, 16f)
				},
				{
					16f,
					new FontSizeInfo(28f, 18f)
				},
				{
					18f,
					new FontSizeInfo(32f, 21f)
				},
				{
					20f,
					new FontSizeInfo(37f, 23f)
				},
				{
					22f,
					new FontSizeInfo(38f, 25f)
				},
				{
					24f,
					new FontSizeInfo(44f, 28f)
				},
				{
					26f,
					new FontSizeInfo(46f, 30f)
				},
				{
					28f,
					new FontSizeInfo(50f, 32f)
				},
				{
					36f,
					new FontSizeInfo(65f, 42f)
				},
				{
					48f,
					new FontSizeInfo(84f, 55f)
				},
				{
					72f,
					new FontSizeInfo(0f, 83f)
				},
				{
					96f,
					new FontSizeInfo(167f, 111f)
				},
				{
					128f,
					new FontSizeInfo(225f, 148f)
				},
				{
					256f,
					new FontSizeInfo(443f, 295f)
				}
			}
		},
		{
			"Stencil",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 9f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 11f)
				},
				{
					16f,
					new FontSizeInfo(28f, 12f)
				},
				{
					18f,
					new FontSizeInfo(32f, 14f)
				},
				{
					20f,
					new FontSizeInfo(35f, 16f)
				},
				{
					22f,
					new FontSizeInfo(38f, 17f)
				},
				{
					24f,
					new FontSizeInfo(42f, 18f)
				},
				{
					26f,
					new FontSizeInfo(46f, 20f)
				},
				{
					28f,
					new FontSizeInfo(48f, 21f)
				},
				{
					36f,
					new FontSizeInfo(62f, 28f)
				},
				{
					48f,
					new FontSizeInfo(83f, 37f)
				},
				{
					72f,
					new FontSizeInfo(0f, 55f)
				},
				{
					96f,
					new FontSizeInfo(164f, 74f)
				},
				{
					128f,
					new FontSizeInfo(219f, 98f)
				},
				{
					256f,
					new FontSizeInfo(436f, 196f)
				}
			}
		},
		{
			"Tw Cen MT",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(20f, 8f)
				},
				{
					12f,
					new FontSizeInfo(21f, 9f)
				},
				{
					14f,
					new FontSizeInfo(25f, 10f)
				},
				{
					16f,
					new FontSizeInfo(27f, 12f)
				},
				{
					18f,
					new FontSizeInfo(31f, 13f)
				},
				{
					20f,
					new FontSizeInfo(34f, 15f)
				},
				{
					22f,
					new FontSizeInfo(37f, 16f)
				},
				{
					24f,
					new FontSizeInfo(40f, 18f)
				},
				{
					26f,
					new FontSizeInfo(44f, 19f)
				},
				{
					28f,
					new FontSizeInfo(47f, 20f)
				},
				{
					36f,
					new FontSizeInfo(60f, 26f)
				},
				{
					48f,
					new FontSizeInfo(80f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 53f)
				},
				{
					96f,
					new FontSizeInfo(159f, 71f)
				},
				{
					128f,
					new FontSizeInfo(212f, 94f)
				},
				{
					256f,
					new FontSizeInfo(421f, 188f)
				}
			}
		},
		{
			"Tw Cen MT Condensed",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 4f)
				},
				{
					10f,
					new FontSizeInfo(20f, 5f)
				},
				{
					11f,
					new FontSizeInfo(20f, 5f)
				},
				{
					12f,
					new FontSizeInfo(21f, 6f)
				},
				{
					14f,
					new FontSizeInfo(25f, 7f)
				},
				{
					16f,
					new FontSizeInfo(27f, 8f)
				},
				{
					18f,
					new FontSizeInfo(31f, 9f)
				},
				{
					20f,
					new FontSizeInfo(34f, 10f)
				},
				{
					22f,
					new FontSizeInfo(37f, 11f)
				},
				{
					24f,
					new FontSizeInfo(41f, 12f)
				},
				{
					26f,
					new FontSizeInfo(45f, 13f)
				},
				{
					28f,
					new FontSizeInfo(48f, 14f)
				},
				{
					36f,
					new FontSizeInfo(61f, 18f)
				},
				{
					48f,
					new FontSizeInfo(80f, 23f)
				},
				{
					72f,
					new FontSizeInfo(0f, 35f)
				},
				{
					96f,
					new FontSizeInfo(160f, 47f)
				},
				{
					128f,
					new FontSizeInfo(214f, 62f)
				},
				{
					256f,
					new FontSizeInfo(424f, 125f)
				}
			}
		},
		{
			"Tw Cen MT Condensed Extra Bold",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(20f, 7f)
				},
				{
					12f,
					new FontSizeInfo(22f, 8f)
				},
				{
					14f,
					new FontSizeInfo(25f, 9f)
				},
				{
					16f,
					new FontSizeInfo(27f, 10f)
				},
				{
					18f,
					new FontSizeInfo(31f, 12f)
				},
				{
					20f,
					new FontSizeInfo(35f, 13f)
				},
				{
					22f,
					new FontSizeInfo(38f, 14f)
				},
				{
					24f,
					new FontSizeInfo(42f, 16f)
				},
				{
					26f,
					new FontSizeInfo(45f, 17f)
				},
				{
					28f,
					new FontSizeInfo(48f, 18f)
				},
				{
					36f,
					new FontSizeInfo(62f, 24f)
				},
				{
					48f,
					new FontSizeInfo(81f, 31f)
				},
				{
					72f,
					new FontSizeInfo(0f, 47f)
				},
				{
					96f,
					new FontSizeInfo(165f, 63f)
				},
				{
					128f,
					new FontSizeInfo(218f, 84f)
				},
				{
					256f,
					new FontSizeInfo(414f, 167f)
				}
			}
		},
		{
			"Tempus Sans ITC",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 9f)
				},
				{
					11f,
					new FontSizeInfo(21f, 10f)
				},
				{
					12f,
					new FontSizeInfo(22f, 11f)
				},
				{
					14f,
					new FontSizeInfo(26f, 13f)
				},
				{
					16f,
					new FontSizeInfo(29f, 14f)
				},
				{
					18f,
					new FontSizeInfo(33f, 16f)
				},
				{
					20f,
					new FontSizeInfo(37f, 18f)
				},
				{
					22f,
					new FontSizeInfo(40f, 19f)
				},
				{
					24f,
					new FontSizeInfo(44f, 21f)
				},
				{
					26f,
					new FontSizeInfo(48f, 23f)
				},
				{
					28f,
					new FontSizeInfo(51f, 25f)
				},
				{
					36f,
					new FontSizeInfo(65f, 32f)
				},
				{
					48f,
					new FontSizeInfo(87f, 43f)
				},
				{
					72f,
					new FontSizeInfo(0f, 64f)
				},
				{
					96f,
					new FontSizeInfo(173f, 85f)
				},
				{
					128f,
					new FontSizeInfo(230f, 114f)
				},
				{
					256f,
					new FontSizeInfo(458f, 226f)
				}
			}
		},
		{
			"Viner Hand ITC",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(23f, 7f)
				},
				{
					10f,
					new FontSizeInfo(24f, 9f)
				},
				{
					11f,
					new FontSizeInfo(25f, 10f)
				},
				{
					12f,
					new FontSizeInfo(27f, 11f)
				},
				{
					14f,
					new FontSizeInfo(32f, 13f)
				},
				{
					16f,
					new FontSizeInfo(35f, 14f)
				},
				{
					18f,
					new FontSizeInfo(40f, 16f)
				},
				{
					20f,
					new FontSizeInfo(45f, 18f)
				},
				{
					22f,
					new FontSizeInfo(48f, 20f)
				},
				{
					24f,
					new FontSizeInfo(53f, 22f)
				},
				{
					26f,
					new FontSizeInfo(58f, 24f)
				},
				{
					28f,
					new FontSizeInfo(61f, 25f)
				},
				{
					36f,
					new FontSizeInfo(79f, 33f)
				},
				{
					48f,
					new FontSizeInfo(104f, 43f)
				},
				{
					72f,
					new FontSizeInfo(0f, 65f)
				},
				{
					96f,
					new FontSizeInfo(208f, 87f)
				},
				{
					128f,
					new FontSizeInfo(277f, 116f)
				},
				{
					256f,
					new FontSizeInfo(552f, 232f)
				}
			}
		},
		{
			"Vivaldi",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(20f, 6f)
				},
				{
					12f,
					new FontSizeInfo(22f, 7f)
				},
				{
					14f,
					new FontSizeInfo(25f, 8f)
				},
				{
					16f,
					new FontSizeInfo(28f, 9f)
				},
				{
					18f,
					new FontSizeInfo(32f, 10f)
				},
				{
					20f,
					new FontSizeInfo(36f, 12f)
				},
				{
					22f,
					new FontSizeInfo(38f, 12f)
				},
				{
					24f,
					new FontSizeInfo(42f, 14f)
				},
				{
					26f,
					new FontSizeInfo(46f, 15f)
				},
				{
					28f,
					new FontSizeInfo(49f, 16f)
				},
				{
					36f,
					new FontSizeInfo(63f, 20f)
				},
				{
					48f,
					new FontSizeInfo(83f, 27f)
				},
				{
					72f,
					new FontSizeInfo(0f, 41f)
				},
				{
					96f,
					new FontSizeInfo(168f, 55f)
				},
				{
					128f,
					new FontSizeInfo(225f, 73f)
				},
				{
					256f,
					new FontSizeInfo(448f, 146f)
				}
			}
		},
		{
			"Vladimir Script",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(21f, 7f)
				},
				{
					12f,
					new FontSizeInfo(22f, 7f)
				},
				{
					14f,
					new FontSizeInfo(26f, 9f)
				},
				{
					16f,
					new FontSizeInfo(29f, 10f)
				},
				{
					18f,
					new FontSizeInfo(32f, 11f)
				},
				{
					20f,
					new FontSizeInfo(36f, 13f)
				},
				{
					22f,
					new FontSizeInfo(39f, 13f)
				},
				{
					24f,
					new FontSizeInfo(43f, 15f)
				},
				{
					26f,
					new FontSizeInfo(47f, 16f)
				},
				{
					28f,
					new FontSizeInfo(50f, 17f)
				},
				{
					36f,
					new FontSizeInfo(64f, 22f)
				},
				{
					48f,
					new FontSizeInfo(85f, 30f)
				},
				{
					72f,
					new FontSizeInfo(0f, 44f)
				},
				{
					96f,
					new FontSizeInfo(169f, 59f)
				},
				{
					128f,
					new FontSizeInfo(225f, 79f)
				},
				{
					256f,
					new FontSizeInfo(448f, 158f)
				}
			}
		},
		{
			"Wingdings 2",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 9f)
				},
				{
					8f,
					new FontSizeInfo(20f, 12f)
				},
				{
					10f,
					new FontSizeInfo(20f, 15f)
				},
				{
					11f,
					new FontSizeInfo(20f, 17f)
				},
				{
					12f,
					new FontSizeInfo(21f, 18f)
				},
				{
					14f,
					new FontSizeInfo(24f, 22f)
				},
				{
					16f,
					new FontSizeInfo(26f, 24f)
				},
				{
					18f,
					new FontSizeInfo(30f, 27f)
				},
				{
					20f,
					new FontSizeInfo(34f, 31f)
				},
				{
					22f,
					new FontSizeInfo(36f, 33f)
				},
				{
					24f,
					new FontSizeInfo(40f, 36f)
				},
				{
					26f,
					new FontSizeInfo(43f, 40f)
				},
				{
					28f,
					new FontSizeInfo(46f, 42f)
				},
				{
					36f,
					new FontSizeInfo(59f, 54f)
				},
				{
					48f,
					new FontSizeInfo(79f, 73f)
				},
				{
					72f,
					new FontSizeInfo(0f, 109f)
				},
				{
					96f,
					new FontSizeInfo(156f, 145f)
				},
				{
					128f,
					new FontSizeInfo(208f, 194f)
				},
				{
					256f,
					new FontSizeInfo(414f, 386f)
				}
			}
		},
		{
			"Wingdings 3",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 4f)
				},
				{
					8f,
					new FontSizeInfo(20f, 10f)
				},
				{
					10f,
					new FontSizeInfo(20f, 12f)
				},
				{
					11f,
					new FontSizeInfo(20f, 13f)
				},
				{
					12f,
					new FontSizeInfo(21f, 14f)
				},
				{
					14f,
					new FontSizeInfo(25f, 17f)
				},
				{
					16f,
					new FontSizeInfo(27f, 19f)
				},
				{
					18f,
					new FontSizeInfo(30f, 21f)
				},
				{
					20f,
					new FontSizeInfo(34f, 24f)
				},
				{
					22f,
					new FontSizeInfo(35f, 26f)
				},
				{
					24f,
					new FontSizeInfo(40f, 29f)
				},
				{
					26f,
					new FontSizeInfo(43f, 19f)
				},
				{
					28f,
					new FontSizeInfo(46f, 33f)
				},
				{
					36f,
					new FontSizeInfo(59f, 27f)
				},
				{
					48f,
					new FontSizeInfo(79f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 53f)
				},
				{
					96f,
					new FontSizeInfo(156f, 71f)
				},
				{
					128f,
					new FontSizeInfo(208f, 95f)
				},
				{
					256f,
					new FontSizeInfo(414f, 189f)
				}
			}
		},
		{
			"Buxton Sketch",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(21f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 9f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 11f)
				},
				{
					16f,
					new FontSizeInfo(30f, 12f)
				},
				{
					18f,
					new FontSizeInfo(34f, 14f)
				},
				{
					20f,
					new FontSizeInfo(38f, 16f)
				},
				{
					22f,
					new FontSizeInfo(41f, 17f)
				},
				{
					24f,
					new FontSizeInfo(45f, 18f)
				},
				{
					26f,
					new FontSizeInfo(49f, 20f)
				},
				{
					28f,
					new FontSizeInfo(52f, 21f)
				},
				{
					36f,
					new FontSizeInfo(67f, 28f)
				},
				{
					48f,
					new FontSizeInfo(90f, 37f)
				},
				{
					72f,
					new FontSizeInfo(0f, 55f)
				},
				{
					96f,
					new FontSizeInfo(178f, 74f)
				},
				{
					128f,
					new FontSizeInfo(238f, 99f)
				},
				{
					256f,
					new FontSizeInfo(473f, 196f)
				}
			}
		},
		{
			"Segoe Marker",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 5f)
				},
				{
					10f,
					new FontSizeInfo(20f, 6f)
				},
				{
					11f,
					new FontSizeInfo(20f, 6f)
				},
				{
					12f,
					new FontSizeInfo(21f, 7f)
				},
				{
					14f,
					new FontSizeInfo(24f, 8f)
				},
				{
					16f,
					new FontSizeInfo(27f, 9f)
				},
				{
					18f,
					new FontSizeInfo(31f, 10f)
				},
				{
					20f,
					new FontSizeInfo(34f, 12f)
				},
				{
					22f,
					new FontSizeInfo(37f, 12f)
				},
				{
					24f,
					new FontSizeInfo(41f, 14f)
				},
				{
					26f,
					new FontSizeInfo(44f, 15f)
				},
				{
					28f,
					new FontSizeInfo(47f, 16f)
				},
				{
					36f,
					new FontSizeInfo(60f, 20f)
				},
				{
					48f,
					new FontSizeInfo(80f, 27f)
				},
				{
					72f,
					new FontSizeInfo(0f, 41f)
				},
				{
					96f,
					new FontSizeInfo(159f, 55f)
				},
				{
					128f,
					new FontSizeInfo(212f, 73f)
				},
				{
					256f,
					new FontSizeInfo(422f, 146f)
				}
			}
		},
		{
			"SketchFlow Print",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 8f)
				},
				{
					10f,
					new FontSizeInfo(20f, 10f)
				},
				{
					11f,
					new FontSizeInfo(20f, 11f)
				},
				{
					12f,
					new FontSizeInfo(21f, 12f)
				},
				{
					14f,
					new FontSizeInfo(25f, 14f)
				},
				{
					16f,
					new FontSizeInfo(28f, 15f)
				},
				{
					18f,
					new FontSizeInfo(31f, 17f)
				},
				{
					20f,
					new FontSizeInfo(35f, 20f)
				},
				{
					22f,
					new FontSizeInfo(38f, 21f)
				},
				{
					24f,
					new FontSizeInfo(41f, 23f)
				},
				{
					26f,
					new FontSizeInfo(45f, 25f)
				},
				{
					28f,
					new FontSizeInfo(48f, 27f)
				},
				{
					36f,
					new FontSizeInfo(61f, 35f)
				},
				{
					48f,
					new FontSizeInfo(81f, 47f)
				},
				{
					72f,
					new FontSizeInfo(0f, 70f)
				},
				{
					96f,
					new FontSizeInfo(161f, 93f)
				},
				{
					128f,
					new FontSizeInfo(215f, 124f)
				},
				{
					256f,
					new FontSizeInfo(427f, 248f)
				}
			}
		},
		{
			"Microsoft MHei",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(29f, 10f)
				},
				{
					16f,
					new FontSizeInfo(31f, 12f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(39f, 15f)
				},
				{
					22f,
					new FontSizeInfo(41f, 16f)
				},
				{
					24f,
					new FontSizeInfo(47f, 18f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(53f, 20f)
				},
				{
					36f,
					new FontSizeInfo(69f, 26f)
				},
				{
					48f,
					new FontSizeInfo(91f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 53f)
				},
				{
					96f,
					new FontSizeInfo(181f, 70f)
				},
				{
					128f,
					new FontSizeInfo(243f, 94f)
				},
				{
					256f,
					new FontSizeInfo(482f, 187f)
				}
			}
		},
		{
			"Microsoft NeoGothic",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(29f, 11f)
				},
				{
					16f,
					new FontSizeInfo(31f, 12f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(39f, 15f)
				},
				{
					22f,
					new FontSizeInfo(41f, 16f)
				},
				{
					24f,
					new FontSizeInfo(47f, 18f)
				},
				{
					26f,
					new FontSizeInfo(51f, 20f)
				},
				{
					28f,
					new FontSizeInfo(53f, 21f)
				},
				{
					36f,
					new FontSizeInfo(69f, 27f)
				},
				{
					48f,
					new FontSizeInfo(91f, 36f)
				},
				{
					72f,
					new FontSizeInfo(0f, 54f)
				},
				{
					96f,
					new FontSizeInfo(181f, 72f)
				},
				{
					128f,
					new FontSizeInfo(243f, 96f)
				},
				{
					256f,
					new FontSizeInfo(482f, 191f)
				}
			}
		},
		{
			"Segoe WP Black",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 7f)
				},
				{
					10f,
					new FontSizeInfo(20f, 8f)
				},
				{
					11f,
					new FontSizeInfo(22f, 9f)
				},
				{
					12f,
					new FontSizeInfo(23f, 10f)
				},
				{
					14f,
					new FontSizeInfo(29f, 12f)
				},
				{
					16f,
					new FontSizeInfo(31f, 13f)
				},
				{
					18f,
					new FontSizeInfo(35f, 15f)
				},
				{
					20f,
					new FontSizeInfo(39f, 17f)
				},
				{
					22f,
					new FontSizeInfo(41f, 18f)
				},
				{
					24f,
					new FontSizeInfo(47f, 20f)
				},
				{
					26f,
					new FontSizeInfo(51f, 22f)
				},
				{
					28f,
					new FontSizeInfo(53f, 23f)
				},
				{
					36f,
					new FontSizeInfo(69f, 30f)
				},
				{
					48f,
					new FontSizeInfo(91f, 40f)
				},
				{
					72f,
					new FontSizeInfo(0f, 60f)
				},
				{
					96f,
					new FontSizeInfo(181f, 79f)
				},
				{
					128f,
					new FontSizeInfo(243f, 106f)
				},
				{
					256f,
					new FontSizeInfo(482f, 212f)
				}
			}
		},
		{
			"Segoe WP",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 11f)
				},
				{
					16f,
					new FontSizeInfo(34f, 12f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 15f)
				},
				{
					22f,
					new FontSizeInfo(44f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 18f)
				},
				{
					26f,
					new FontSizeInfo(51f, 20f)
				},
				{
					28f,
					new FontSizeInfo(54f, 21f)
				},
				{
					36f,
					new FontSizeInfo(70f, 27f)
				},
				{
					48f,
					new FontSizeInfo(92f, 36f)
				},
				{
					72f,
					new FontSizeInfo(0f, 54f)
				},
				{
					96f,
					new FontSizeInfo(180f, 72f)
				},
				{
					128f,
					new FontSizeInfo(241f, 96f)
				},
				{
					256f,
					new FontSizeInfo(482f, 191f)
				}
			}
		},
		{
			"Segoe WP Semibold",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 9f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 11f)
				},
				{
					16f,
					new FontSizeInfo(34f, 12f)
				},
				{
					18f,
					new FontSizeInfo(35f, 14f)
				},
				{
					20f,
					new FontSizeInfo(41f, 16f)
				},
				{
					22f,
					new FontSizeInfo(44f, 17f)
				},
				{
					24f,
					new FontSizeInfo(50f, 18f)
				},
				{
					26f,
					new FontSizeInfo(51f, 20f)
				},
				{
					28f,
					new FontSizeInfo(54f, 21f)
				},
				{
					36f,
					new FontSizeInfo(70f, 28f)
				},
				{
					48f,
					new FontSizeInfo(92f, 37f)
				},
				{
					72f,
					new FontSizeInfo(0f, 55f)
				},
				{
					96f,
					new FontSizeInfo(180f, 74f)
				},
				{
					128f,
					new FontSizeInfo(241f, 99f)
				},
				{
					256f,
					new FontSizeInfo(482f, 196f)
				}
			}
		},
		{
			"Segoe WP Light",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 8f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 11f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 14f)
				},
				{
					22f,
					new FontSizeInfo(44f, 15f)
				},
				{
					24f,
					new FontSizeInfo(50f, 17f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 25f)
				},
				{
					48f,
					new FontSizeInfo(92f, 34f)
				},
				{
					72f,
					new FontSizeInfo(0f, 51f)
				},
				{
					96f,
					new FontSizeInfo(180f, 68f)
				},
				{
					128f,
					new FontSizeInfo(241f, 91f)
				},
				{
					256f,
					new FontSizeInfo(482f, 181f)
				}
			}
		},
		{
			"Segoe WP SemiLight",
			new Dictionary<float, FontSizeInfo>
			{
				{
					6f,
					new FontSizeInfo(20f, 5f)
				},
				{
					8f,
					new FontSizeInfo(20f, 6f)
				},
				{
					10f,
					new FontSizeInfo(20f, 7f)
				},
				{
					11f,
					new FontSizeInfo(22f, 8f)
				},
				{
					12f,
					new FontSizeInfo(23f, 9f)
				},
				{
					14f,
					new FontSizeInfo(27f, 10f)
				},
				{
					16f,
					new FontSizeInfo(34f, 12f)
				},
				{
					18f,
					new FontSizeInfo(35f, 13f)
				},
				{
					20f,
					new FontSizeInfo(41f, 15f)
				},
				{
					22f,
					new FontSizeInfo(44f, 16f)
				},
				{
					24f,
					new FontSizeInfo(50f, 18f)
				},
				{
					26f,
					new FontSizeInfo(51f, 19f)
				},
				{
					28f,
					new FontSizeInfo(54f, 20f)
				},
				{
					36f,
					new FontSizeInfo(70f, 26f)
				},
				{
					48f,
					new FontSizeInfo(92f, 35f)
				},
				{
					72f,
					new FontSizeInfo(0f, 53f)
				},
				{
					96f,
					new FontSizeInfo(180f, 70f)
				},
				{
					128f,
					new FontSizeInfo(241f, 94f)
				},
				{
					256f,
					new FontSizeInfo(482f, 187f)
				}
			}
		}
	};
}
