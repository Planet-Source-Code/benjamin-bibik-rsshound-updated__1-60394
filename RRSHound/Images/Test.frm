VERSION 5.00
Begin VB.Form Test
   BorderStyle = 0        'None
   ClientHeight = 5220
   ClientLeft = 0
   ClientTop = 0
   ClientWidth = 9105
   ControlBox = 0          'False
   LinkTopic = "Form1"
   MaxButton = 0           'False
   MinButton = 0           'False
   ScaleHeight = 4020
   ScaleWidth = 6435
   ShowInTaskbar = 0       'False
   StartUpPosition = 3    'Windows Default
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'API Delcares
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Dim bytRegion(5551) As Byte
Dim nBytes As Long

Private Sub Form_Load()
Dim rgnMain as Long

nBytes = 5552

LoadBytes

rgnMain = ExtCreateRegion(ByVal 0&, nBytes, bytRegion(0))
SetWindowRgn Me.hwnd, rgnMain, True

End Sub
Private Sub LoadBytes()
bytRegion(0) = 32
bytRegion(4) = 1
bytRegion(8) = 89
bytRegion(9) = 1
bytRegion(12) = 144
bytRegion(13) = 21
bytRegion(16) = 9
bytRegion(20) = 14
bytRegion(24) = 71
bytRegion(25) = 2
bytRegion(28) = 92
bytRegion(29) = 1
bytRegion(32) = 12
bytRegion(36) = 14
bytRegion(40) = 55
bytRegion(44) = 15
bytRegion(48) = 86
bytRegion(52) = 14
bytRegion(56) = 123
bytRegion(60) = 15
bytRegion(64) = 156
bytRegion(68) = 14
bytRegion(72) = 193
bytRegion(76) = 15
bytRegion(81) = 1
bytRegion(84) = 14
bytRegion(88) = 10
bytRegion(89) = 1
bytRegion(92) = 15
bytRegion(96) = 42
bytRegion(97) = 1
bytRegion(100) = 14
bytRegion(104) = 53
bytRegion(105) = 1
bytRegion(108) = 15
bytRegion(112) = 12
bytRegion(116) = 15
bytRegion(120) = 57
bytRegion(124) = 16
bytRegion(128) = 84
bytRegion(132) = 15
bytRegion(136) = 125
bytRegion(140) = 16
bytRegion(144) = 154
bytRegion(148) = 15
bytRegion(152) = 195
bytRegion(156) = 16
bytRegion(161) = 1
bytRegion(164) = 15
bytRegion(168) = 10
bytRegion(169) = 1
bytRegion(172) = 16
bytRegion(176) = 42
bytRegion(177) = 1
bytRegion(180) = 15
bytRegion(184) = 53
bytRegion(185) = 1
bytRegion(188) = 16
bytRegion(192) = 12
bytRegion(196) = 16
bytRegion(200) = 58
bytRegion(204) = 17
bytRegion(208) = 83
bytRegion(212) = 16
bytRegion(216) = 126
bytRegion(220) = 17
bytRegion(224) = 153
bytRegion(228) = 16
bytRegion(232) = 196
bytRegion(236) = 17
bytRegion(241) = 1
bytRegion(244) = 16
bytRegion(248) = 10
bytRegion(249) = 1
bytRegion(252) = 17
bytRegion(256) = 42
bytRegion(257) = 1
bytRegion(260) = 16
bytRegion(264) = 53
bytRegion(265) = 1
bytRegion(268) = 17
bytRegion(272) = 12
bytRegion(276) = 17
bytRegion(280) = 59
bytRegion(284) = 18
bytRegion(288) = 82
bytRegion(292) = 17
bytRegion(296) = 127
bytRegion(300) = 18
bytRegion(304) = 152
bytRegion(308) = 17
bytRegion(312) = 197
bytRegion(316) = 18
bytRegion(321) = 1
bytRegion(324) = 17
bytRegion(328) = 10
bytRegion(329) = 1
bytRegion(332) = 18
bytRegion(336) = 42
bytRegion(337) = 1
bytRegion(340) = 17
bytRegion(344) = 53
bytRegion(345) = 1
bytRegion(348) = 18
bytRegion(352) = 12
bytRegion(356) = 18
bytRegion(360) = 60
bytRegion(364) = 19
bytRegion(368) = 82
bytRegion(372) = 18
bytRegion(376) = 127
bytRegion(380) = 19
bytRegion(384) = 152
bytRegion(388) = 18
bytRegion(392) = 197
bytRegion(396) = 19
bytRegion(401) = 1
bytRegion(404) = 18
bytRegion(408) = 10
bytRegion(409) = 1
bytRegion(412) = 19
bytRegion(416) = 42
bytRegion(417) = 1
bytRegion(420) = 18
bytRegion(424) = 53
bytRegion(425) = 1
bytRegion(428) = 19
bytRegion(432) = 12
bytRegion(436) = 19
bytRegion(440) = 60
bytRegion(444) = 22
bytRegion(448) = 81
bytRegion(452) = 19
bytRegion(456) = 128
bytRegion(460) = 22
bytRegion(464) = 151
bytRegion(468) = 19
bytRegion(472) = 198
bytRegion(476) = 22
bytRegion(481) = 1
bytRegion(484) = 19
bytRegion(488) = 10
bytRegion(489) = 1
bytRegion(492) = 22
bytRegion(496) = 42
bytRegion(497) = 1
bytRegion(500) = 19
bytRegion(504) = 53
bytRegion(505) = 1
bytRegion(508) = 22
bytRegion(512) = 12
bytRegion(516) = 22
bytRegion(520) = 60
bytRegion(524) = 23
bytRegion(528) = 81
bytRegion(532) = 22
bytRegion(536) = 128
bytRegion(540) = 23
bytRegion(544) = 151
bytRegion(548) = 22
bytRegion(552) = 198
bytRegion(556) = 23
bytRegion(561) = 1
bytRegion(564) = 22
bytRegion(568) = 10
bytRegion(569) = 1
bytRegion(572) = 23
bytRegion(576) = 42
bytRegion(577) = 1
bytRegion(580) = 22
bytRegion(584) = 53
bytRegion(585) = 1
bytRegion(588) = 23
bytRegion(592) = 76
bytRegion(593) = 1
bytRegion(596) = 22
bytRegion(600) = 115
bytRegion(601) = 1
bytRegion(604) = 23
bytRegion(608) = 135
bytRegion(609) = 1
bytRegion(612) = 22
bytRegion(616) = 145
bytRegion(617) = 1
bytRegion(620) = 23
bytRegion(624) = 168
bytRegion(625) = 1
bytRegion(628) = 22
bytRegion(632) = 178
bytRegion(633) = 1
bytRegion(636) = 23
bytRegion(640) = 195
bytRegion(641) = 1
bytRegion(644) = 22
bytRegion(648) = 203
bytRegion(649) = 1
bytRegion(652) = 23
bytRegion(656) = 231
bytRegion(657) = 1
bytRegion(660) = 22
bytRegion(664) = 240
bytRegion(665) = 1
bytRegion(668) = 23
bytRegion(672) = 3
bytRegion(673) = 2
bytRegion(676) = 22
bytRegion(680) = 36
bytRegion(681) = 2
bytRegion(684) = 23
bytRegion(688) = 12
bytRegion(692) = 23
bytRegion(696) = 22
bytRegion(700) = 24
bytRegion(704) = 49
bytRegion(708) = 23
bytRegion(712) = 60
bytRegion(716) = 24
bytRegion(720) = 81
bytRegion(724) = 23
bytRegion(728) = 90
bytRegion(732) = 24
bytRegion(736) = 119
bytRegion(740) = 23
bytRegion(744) = 128
bytRegion(748) = 24
bytRegion(752) = 151
bytRegion(756) = 23
bytRegion(760) = 160
bytRegion(764) = 24
bytRegion(768) = 189
bytRegion(772) = 23
bytRegion(776) = 198
bytRegion(780) = 24
bytRegion(785) = 1
bytRegion(788) = 23
bytRegion(792) = 10
bytRegion(793) = 1
bytRegion(796) = 24
bytRegion(800) = 42
bytRegion(801) = 1
bytRegion(804) = 23
bytRegion(808) = 53
bytRegion(809) = 1
bytRegion(812) = 24
bytRegion(816) = 74
bytRegion(817) = 1
bytRegion(820) = 23
bytRegion(824) = 116
bytRegion(825) = 1
bytRegion(828) = 24
bytRegion(832) = 135
bytRegion(833) = 1
bytRegion(836) = 23
bytRegion(840) = 145
bytRegion(841) = 1
bytRegion(844) = 24
bytRegion(848) = 168
bytRegion(849) = 1
bytRegion(852) = 23
bytRegion(856) = 178
bytRegion(857) = 1
bytRegion(860) = 24
bytRegion(864) = 195
bytRegion(865) = 1
bytRegion(868) = 23
bytRegion(872) = 204
bytRegion(873) = 1
bytRegion(876) = 24
bytRegion(880) = 231
bytRegion(881) = 1
bytRegion(884) = 23
bytRegion(888) = 240
bytRegion(889) = 1
bytRegion(892) = 24
bytRegion(896) = 3
bytRegion(897) = 2
bytRegion(900) = 23
bytRegion(904) = 39
bytRegion(905) = 2
bytRegion(908) = 24
bytRegion(912) = 12
bytRegion(916) = 24
bytRegion(920) = 22
bytRegion(924) = 25
bytRegion(928) = 50
bytRegion(932) = 24
bytRegion(936) = 60
bytRegion(940) = 25
bytRegion(944) = 81
bytRegion(948) = 24
bytRegion(952) = 90
bytRegion(956) = 25
bytRegion(960) = 119
bytRegion(964) = 24
bytRegion(968) = 128
bytRegion(972) = 25
bytRegion(976) = 151
bytRegion(980) = 24
bytRegion(984) = 160
bytRegion(988) = 25
bytRegion(992) = 189
bytRegion(996) = 24
bytRegion(1000) = 198
bytRegion(1004) = 25
bytRegion(1009) = 1
bytRegion(1012) = 24
bytRegion(1016) = 10
bytRegion(1017) = 1
bytRegion(1020) = 25
bytRegion(1024) = 42
bytRegion(1025) = 1
bytRegion(1028) = 24
bytRegion(1032) = 53
bytRegion(1033) = 1
bytRegion(1036) = 25
bytRegion(1040) = 73
bytRegion(1041) = 1
bytRegion(1044) = 24
bytRegion(1048) = 117
bytRegion(1049) = 1
bytRegion(1052) = 25
bytRegion(1056) = 135
bytRegion(1057) = 1
bytRegion(1060) = 24
bytRegion(1064) = 145
bytRegion(1065) = 1
bytRegion(1068) = 25
bytRegion(1072) = 168
bytRegion(1073) = 1
bytRegion(1076) = 24
bytRegion(1080) = 178
bytRegion(1081) = 1
bytRegion(1084) = 25
bytRegion(1088) = 195
bytRegion(1089) = 1
bytRegion(1092) = 24
bytRegion(1096) = 206
bytRegion(1097) = 1
bytRegion(1100) = 25
bytRegion(1104) = 231
bytRegion(1105) = 1
bytRegion(1108) = 24
bytRegion(1112) = 240
bytRegion(1113) = 1
bytRegion(1116) = 25
bytRegion(1120) = 3
bytRegion(1121) = 2
bytRegion(1124) = 24
bytRegion(1128) = 42
bytRegion(1129) = 2
bytRegion(1132) = 25
bytRegion(1136) = 12
bytRegion(1140) = 25
bytRegion(1144) = 22
bytRegion(1148) = 26
bytRegion(1152) = 50
bytRegion(1156) = 25
bytRegion(1160) = 60
bytRegion(1164) = 26
bytRegion(1168) = 80
bytRegion(1172) = 25
bytRegion(1176) = 90
bytRegion(1180) = 26
bytRegion(1184) = 119
bytRegion(1188) = 25
bytRegion(1192) = 128
bytRegion(1196) = 26
bytRegion(1200) = 151
bytRegion(1204) = 25
bytRegion(1208) = 160
bytRegion(1212) = 26
bytRegion(1216) = 189
bytRegion(1220) = 25
bytRegion(1224) = 198
bytRegion(1228) = 26
bytRegion(1233) = 1
bytRegion(1236) = 25
bytRegion(1240) = 10
bytRegion(1241) = 1
bytRegion(1244) = 26
bytRegion(1248) = 42
bytRegion(1249) = 1
bytRegion(1252) = 25
bytRegion(1256) = 53
bytRegion(1257) = 1
bytRegion(1260) = 26
bytRegion(1264) = 73
bytRegion(1265) = 1
bytRegion(1268) = 25
bytRegion(1272) = 118
bytRegion(1273) = 1
bytRegion(1276) = 26
bytRegion(1280) = 135
bytRegion(1281) = 1
bytRegion(1284) = 25
bytRegion(1288) = 145
bytRegion(1289) = 1
bytRegion(1292) = 26
bytRegion(1296) = 168
bytRegion(1297) = 1
bytRegion(1300) = 25
bytRegion(1304) = 178
bytRegion(1305) = 1
bytRegion(1308) = 26
bytRegion(1312) = 195
bytRegion(1313) = 1
bytRegion(1316) = 25
bytRegion(1320) = 207
bytRegion(1321) = 1
bytRegion(1324) = 26
bytRegion(1328) = 231
bytRegion(1329) = 1
bytRegion(1332) = 25
bytRegion(1336) = 240
bytRegion(1337) = 1
bytRegion(1340) = 26
bytRegion(1344) = 3
bytRegion(1345) = 2
bytRegion(1348) = 25
bytRegion(1352) = 43
bytRegion(1353) = 2
bytRegion(1356) = 26
bytRegion(1360) = 12
bytRegion(1364) = 26
bytRegion(1368) = 22
bytRegion(1372) = 27
bytRegion(1376) = 50
bytRegion(1380) = 26
bytRegion(1384) = 60
bytRegion(1388) = 27
bytRegion(1392) = 80
bytRegion(1396) = 26
bytRegion(1400) = 90
bytRegion(1404) = 27
bytRegion(1408) = 119
bytRegion(1412) = 26
bytRegion(1416) = 128
bytRegion(1420) = 27
bytRegion(1424) = 151
bytRegion(1428) = 26
bytRegion(1432) = 160
bytRegion(1436) = 27
bytRegion(1440) = 189
bytRegion(1444) = 26
bytRegion(1448) = 198
bytRegion(1452) = 27
bytRegion(1457) = 1
bytRegion(1460) = 26
bytRegion(1464) = 10
bytRegion(1465) = 1
bytRegion(1468) = 27
bytRegion(1472) = 42
bytRegion(1473) = 1
bytRegion(1476) = 26
bytRegion(1480) = 53
bytRegion(1481) = 1
bytRegion(1484) = 27
bytRegion(1488) = 72
bytRegion(1489) = 1
bytRegion(1492) = 26
bytRegion(1496) = 118
bytRegion(1497) = 1
bytRegion(1500) = 27
bytRegion(1504) = 135
bytRegion(1505) = 1
bytRegion(1508) = 26
bytRegion(1512) = 145
bytRegion(1513) = 1
bytRegion(1516) = 27
bytRegion(1520) = 168
bytRegion(1521) = 1
bytRegion(1524) = 26
bytRegion(1528) = 178
bytRegion(1529) = 1
bytRegion(1532) = 27
bytRegion(1536) = 195
bytRegion(1537) = 1
bytRegion(1540) = 26
bytRegion(1544) = 208
bytRegion(1545) = 1
bytRegion(1548) = 27
bytRegion(1552) = 231
bytRegion(1553) = 1
bytRegion(1556) = 26
bytRegion(1560) = 240
bytRegion(1561) = 1
bytRegion(1564) = 27
bytRegion(1568) = 3
bytRegion(1569) = 2
bytRegion(1572) = 26
bytRegion(1576) = 44
bytRegion(1577) = 2
bytRegion(1580) = 27
bytRegion(1584) = 12
bytRegion(1588) = 27
bytRegion(1592) = 22
bytRegion(1596) = 28
bytRegion(1600) = 50
bytRegion(1604) = 27
bytRegion(1608) = 60
bytRegion(1612) = 28
bytRegion(1616) = 80
bytRegion(1620) = 27
bytRegion(1624) = 90
bytRegion(1628) = 28
bytRegion(1632) = 119
bytRegion(1636) = 27
bytRegion(1640) = 127
bytRegion(1644) = 28
bytRegion(1648) = 151
bytRegion(1652) = 27
bytRegion(1656) = 160
bytRegion(1660) = 28
bytRegion(1664) = 189
bytRegion(1668) = 27
bytRegion(1672) = 197
bytRegion(1676) = 28
bytRegion(1681) = 1
bytRegion(1684) = 27
bytRegion(1688) = 10
bytRegion(1689) = 1
bytRegion(1692) = 28
bytRegion(1696) = 42
bytRegion(1697) = 1
bytRegion(1700) = 27
bytRegion(1704) = 53
bytRegion(1705) = 1
bytRegion(1708) = 28
bytRegion(1712) = 72
bytRegion(1713) = 1
bytRegion(1716) = 27
bytRegion(1720) = 119
bytRegion(1721) = 1
bytRegion(1724) = 28
bytRegion(1728) = 135
bytRegion(1729) = 1
bytRegion(1732) = 27
bytRegion(1736) = 145
bytRegion(1737) = 1
bytRegion(1740) = 28
bytRegion(1744) = 168
bytRegion(1745) = 1
bytRegion(1748) = 27
bytRegion(1752) = 178
bytRegion(1753) = 1
bytRegion(1756) = 28
bytRegion(1760) = 195
bytRegion(1761) = 1
bytRegion(1764) = 27
bytRegion(1768) = 209
bytRegion(1769) = 1
bytRegion(1772) = 28
bytRegion(1776) = 231
bytRegion(1777) = 1
bytRegion(1780) = 27
bytRegion(1784) = 240
bytRegion(1785) = 1
bytRegion(1788) = 28
bytRegion(1792) = 3
bytRegion(1793) = 2
bytRegion(1796) = 27
bytRegion(1800) = 45
bytRegion(1801) = 2
bytRegion(1804) = 28
bytRegion(1808) = 12
bytRegion(1812) = 28
bytRegion(1816) = 22
bytRegion(1820) = 29
bytRegion(1824) = 50
bytRegion(1828) = 28
bytRegion(1832) = 60
bytRegion(1836) = 29
bytRegion(1840) = 80
bytRegion(1844) = 28
bytRegion(1848) = 90
bytRegion(1852) = 29
bytRegion(1856) = 119
bytRegion(1860) = 28
bytRegion(1864) = 124
bytRegion(1868) = 29
bytRegion(1872) = 151
bytRegion(1876) = 28
bytRegion(1880) = 160
bytRegion(1884) = 29
bytRegion(1888) = 189
bytRegion(1892) = 28
bytRegion(1896) = 194
bytRegion(1900) = 29
bytRegion(1905) = 1
bytRegion(1908) = 28
bytRegion(1912) = 10
bytRegion(1913) = 1
bytRegion(1916) = 29
bytRegion(1920) = 42
bytRegion(1921) = 1
bytRegion(1924) = 28
bytRegion(1928) = 53
bytRegion(1929) = 1
bytRegion(1932) = 29
bytRegion(1936) = 72
bytRegion(1937) = 1
bytRegion(1940) = 28
bytRegion(1944) = 119
bytRegion(1945) = 1
bytRegion(1948) = 29
bytRegion(1952) = 135
bytRegion(1953) = 1
bytRegion(1956) = 28
bytRegion(1960) = 145
bytRegion(1961) = 1
bytRegion(1964) = 29
bytRegion(1968) = 168
bytRegion(1969) = 1
bytRegion(1972) = 28
bytRegion(1976) = 178
bytRegion(1977) = 1
bytRegion(1980) = 29
bytRegion(1984) = 195
bytRegion(1985) = 1
bytRegion(1988) = 28
bytRegion(1992) = 210
bytRegion(1993) = 1
bytRegion(1996) = 29
bytRegion(2000) = 231
bytRegion(2001) = 1
bytRegion(2004) = 28
bytRegion(2008) = 240
bytRegion(2009) = 1
bytRegion(2012) = 29
bytRegion(2016) = 3
bytRegion(2017) = 2
bytRegion(2020) = 28
bytRegion(2024) = 46
bytRegion(2025) = 2
bytRegion(2028) = 29
bytRegion(2032) = 12
bytRegion(2036) = 29
bytRegion(2040) = 22
bytRegion(2044) = 30
bytRegion(2048) = 50
bytRegion(2052) = 29
bytRegion(2056) = 60
bytRegion(2060) = 30
bytRegion(2064) = 80
bytRegion(2068) = 29
bytRegion(2072) = 90
bytRegion(2076) = 30
bytRegion(2080) = 119
bytRegion(2084) = 29
bytRegion(2088) = 120
bytRegion(2092) = 30
bytRegion(2096) = 151
bytRegion(2100) = 29
bytRegion(2104) = 160
bytRegion(2108) = 30
bytRegion(2112) = 189
bytRegion(2116) = 29
bytRegion(2120) = 190
bytRegion(2124) = 30
bytRegion(2129) = 1
bytRegion(2132) = 29
bytRegion(2136) = 10
bytRegion(2137) = 1
bytRegion(2140) = 30
bytRegion(2144) = 42
bytRegion(2145) = 1
bytRegion(2148) = 29
bytRegion(2152) = 53
bytRegion(2153) = 1
bytRegion(2156) = 30
bytRegion(2160) = 72
bytRegion(2161) = 1
bytRegion(2164) = 29
bytRegion(2168) = 119
bytRegion(2169) = 1
bytRegion(2172) = 30
bytRegion(2176) = 135
bytRegion(2177) = 1
bytRegion(2180) = 29
bytRegion(2184) = 145
bytRegion(2185) = 1
bytRegion(2188) = 30
bytRegion(2192) = 168
bytRegion(2193) = 1
bytRegion(2196) = 29
bytRegion(2200) = 178
bytRegion(2201) = 1
bytRegion(2204) = 30
bytRegion(2208) = 195
bytRegion(2209) = 1
bytRegion(2212) = 29
bytRegion(2216) = 211
bytRegion(2217) = 1
bytRegion(2220) = 30
bytRegion(2224) = 231
bytRegion(2225) = 1
bytRegion(2228) = 29
bytRegion(2232) = 240
bytRegion(2233) = 1
bytRegion(2236) = 30
bytRegion(2240) = 3
bytRegion(2241) = 2
bytRegion(2244) = 29
bytRegion(2248) = 47
bytRegion(2249) = 2
bytRegion(2252) = 30
bytRegion(2256) = 12
bytRegion(2260) = 30
bytRegion(2264) = 22
bytRegion(2268) = 31
bytRegion(2272) = 50
bytRegion(2276) = 30
bytRegion(2280) = 60
bytRegion(2284) = 31
bytRegion(2288) = 80
bytRegion(2292) = 30
bytRegion(2296) = 90
bytRegion(2300) = 31
bytRegion(2304) = 114
bytRegion(2308) = 30
bytRegion(2312) = 66
bytRegion(2313) = 2
bytRegion(2316) = 31
bytRegion(2320) = 12
bytRegion(2324) = 31
bytRegion(2328) = 22
bytRegion(2332) = 32
bytRegion(2336) = 50
bytRegion(2340) = 31
bytRegion(2344) = 60
bytRegion(2348) = 32
bytRegion(2352) = 81
bytRegion(2356) = 31
bytRegion(2360) = 90
bytRegion(2364) = 32
bytRegion(2368) = 112
bytRegion(2372) = 31
bytRegion(2376) = 68
bytRegion(2377) = 2
bytRegion(2380) = 32
bytRegion(2384) = 12
bytRegion(2388) = 32
bytRegion(2392) = 22
bytRegion(2396) = 33
bytRegion(2400) = 50
bytRegion(2404) = 32
bytRegion(2408) = 60
bytRegion(2412) = 33
bytRegion(2416) = 81
bytRegion(2420) = 32
bytRegion(2424) = 69
bytRegion(2425) = 2
bytRegion(2428) = 33
bytRegion(2432) = 12
bytRegion(2436) = 33
bytRegion(2440) = 22
bytRegion(2444) = 34
bytRegion(2448) = 50
bytRegion(2452) = 33
bytRegion(2456) = 60
bytRegion(2460) = 34
bytRegion(2464) = 81
bytRegion(2468) = 33
bytRegion(2472) = 70
bytRegion(2473) = 2
bytRegion(2476) = 34
bytRegion(2480) = 12
bytRegion(2484) = 34
bytRegion(2488) = 22
bytRegion(2492) = 35
bytRegion(2496) = 49
bytRegion(2500) = 34
bytRegion(2504) = 60
bytRegion(2508) = 35
bytRegion(2512) = 81
bytRegion(2516) = 34
bytRegion(2520) = 70
bytRegion(2521) = 2
bytRegion(2524) = 35
bytRegion(2528) = 12
bytRegion(2532) = 35
bytRegion(2536) = 60
bytRegion(2540) = 37
bytRegion(2544) = 81
bytRegion(2548) = 35
bytRegion(2552) = 71
bytRegion(2553) = 2
bytRegion(2556) = 37
bytRegion(2560) = 12
bytRegion(2564) = 37
bytRegion(2568) = 60
bytRegion(2572) = 39
bytRegion(2576) = 82
bytRegion(2580) = 37
bytRegion(2584) = 71
bytRegion(2585) = 2
bytRegion(2588) = 39
bytRegion(2592) = 12
bytRegion(2596) = 39
bytRegion(2600) = 60
bytRegion(2604) = 40
bytRegion(2608) = 83
bytRegion(2612) = 39
bytRegion(2616) = 71
bytRegion(2617) = 2
bytRegion(2620) = 40
bytRegion(2624) = 12
bytRegion(2628) = 40
bytRegion(2632) = 59
bytRegion(2636) = 41
bytRegion(2640) = 85
bytRegion(2644) = 40
bytRegion(2648) = 71
bytRegion(2649) = 2
bytRegion(2652) = 41
bytRegion(2656) = 12
bytRegion(2660) = 41
bytRegion(2664) = 58
bytRegion(2668) = 42
bytRegion(2672) = 88
bytRegion(2676) = 41
bytRegion(2680) = 71
bytRegion(2681) = 2
bytRegion(2684) = 42
bytRegion(2688) = 12
bytRegion(2692) = 42
bytRegion(2696) = 57
bytRegion(2700) = 43
bytRegion(2704) = 109
bytRegion(2708) = 42
bytRegion(2712) = 71
bytRegion(2713) = 2
bytRegion(2716) = 43
bytRegion(2720) = 12
bytRegion(2724) = 43
bytRegion(2728) = 54
bytRegion(2732) = 44
bytRegion(2736) = 109
bytRegion(2740) = 43
bytRegion(2744) = 71
bytRegion(2745) = 2
bytRegion(2748) = 44
bytRegion(2752) = 12
bytRegion(2756) = 44
bytRegion(2760) = 22
bytRegion(2764) = 45
bytRegion(2768) = 34
bytRegion(2772) = 44
bytRegion(2776) = 48
bytRegion(2780) = 45
bytRegion(2784) = 109
bytRegion(2788) = 44
bytRegion(2792) = 71
bytRegion(2793) = 2
bytRegion(2796) = 45
bytRegion(2800) = 12
bytRegion(2804) = 45
bytRegion(2808) = 22
bytRegion(2812) = 46
bytRegion(2816) = 35
bytRegion(2820) = 45
bytRegion(2824) = 49
bytRegion(2828) = 46
bytRegion(2832) = 84
bytRegion(2836) = 45
bytRegion(2840) = 89
bytRegion(2844) = 46
bytRegion(2848) = 109
bytRegion(2852) = 45
bytRegion(2856) = 71
bytRegion(2857) = 2
bytRegion(2860) = 46
bytRegion(2864) = 12
bytRegion(2868) = 46
bytRegion(2872) = 22
bytRegion(2876) = 47
bytRegion(2880) = 36
bytRegion(2884) = 46
bytRegion(2888) = 50
bytRegion(2892) = 47
bytRegion(2896) = 80
bytRegion(2900) = 46
bytRegion(2904) = 89
bytRegion(2908) = 47
bytRegion(2912) = 109
bytRegion(2916) = 46
bytRegion(2920) = 71
bytRegion(2921) = 2
bytRegion(2924) = 47
bytRegion(2928) = 12
bytRegion(2932) = 47
bytRegion(2936) = 22
bytRegion(2940) = 48
bytRegion(2944) = 37
bytRegion(2948) = 47
bytRegion(2952) = 51
bytRegion(2956) = 48
bytRegion(2960) = 79
bytRegion(2964) = 47
bytRegion(2968) = 89
bytRegion(2972) = 48
bytRegion(2976) = 109
bytRegion(2980) = 47
bytRegion(2984) = 71
bytRegion(2985) = 2
bytRegion(2988) = 48
bytRegion(2992) = 12
bytRegion(2996) = 48
bytRegion(3000) = 22
bytRegion(3004) = 49
bytRegion(3008) = 38
bytRegion(3012) = 48
bytRegion(3016) = 53
bytRegion(3020) = 49
bytRegion(3024) = 79
bytRegion(3028) = 48
bytRegion(3032) = 89
bytRegion(3036) = 49
bytRegion(3040) = 109
bytRegion(3044) = 48
bytRegion(3048) = 71
bytRegion(3049) = 2
bytRegion(3052) = 49
bytRegion(3056) = 12
bytRegion(3060) = 49
bytRegion(3064) = 22
bytRegion(3068) = 50
bytRegion(3072) = 39
bytRegion(3076) = 49
bytRegion(3080) = 54
bytRegion(3084) = 50
bytRegion(3088) = 79
bytRegion(3092) = 49
bytRegion(3096) = 89
bytRegion(3100) = 50
bytRegion(3104) = 109
bytRegion(3108) = 49
bytRegion(3112) = 71
bytRegion(3113) = 2
bytRegion(3116) = 50
bytRegion(3120) = 12
bytRegion(3124) = 50
bytRegion(3128) = 22
bytRegion(3132) = 51
bytRegion(3136) = 40
bytRegion(3140) = 50
bytRegion(3144) = 55
bytRegion(3148) = 51
bytRegion(3152) = 79
bytRegion(3156) = 50
bytRegion(3160) = 89
bytRegion(3164) = 51
bytRegion(3168) = 109
bytRegion(3172) = 50
bytRegion(3176) = 71
bytRegion(3177) = 2
bytRegion(3180) = 51
bytRegion(3184) = 12
bytRegion(3188) = 51
bytRegion(3192) = 22
bytRegion(3196) = 52
bytRegion(3200) = 41
bytRegion(3204) = 51
bytRegion(3208) = 56
bytRegion(3212) = 52
bytRegion(3216) = 79
bytRegion(3220) = 51
bytRegion(3224) = 89
bytRegion(3228) = 52
bytRegion(3232) = 109
bytRegion(3236) = 51
bytRegion(3240) = 71
bytRegion(3241) = 2
bytRegion(3244) = 52
bytRegion(3248) = 12
bytRegion(3252) = 52
bytRegion(3256) = 22
bytRegion(3260) = 53
bytRegion(3264) = 42
bytRegion(3268) = 52
bytRegion(3272) = 57
bytRegion(3276) = 53
bytRegion(3280) = 79
bytRegion(3284) = 52
bytRegion(3288) = 71
bytRegion(3289) = 2
bytRegion(3292) = 53
bytRegion(3296) = 12
bytRegion(3300) = 53
bytRegion(3304) = 22
bytRegion(3308) = 54
bytRegion(3312) = 43
bytRegion(3316) = 53
bytRegion(3320) = 58
bytRegion(3324) = 54
bytRegion(3328) = 79
bytRegion(3332) = 53
bytRegion(3336) = 71
bytRegion(3337) = 2
bytRegion(3340) = 54
bytRegion(3344) = 12
bytRegion(3348) = 54
bytRegion(3352) = 22
bytRegion(3356) = 55
bytRegion(3360) = 44
bytRegion(3364) = 54
bytRegion(3368) = 60
bytRegion(3372) = 55
bytRegion(3376) = 79
bytRegion(3380) = 54
bytRegion(3384) = 71
bytRegion(3385) = 2
bytRegion(3388) = 55
bytRegion(3392) = 12
bytRegion(3396) = 55
bytRegion(3400) = 22
bytRegion(3404) = 56
bytRegion(3408) = 45
bytRegion(3412) = 55
bytRegion(3416) = 61
bytRegion(3420) = 56
bytRegion(3424) = 79
bytRegion(3428) = 55
bytRegion(3432) = 71
bytRegion(3433) = 2
bytRegion(3436) = 56
bytRegion(3440) = 12
bytRegion(3444) = 56
bytRegion(3448) = 22
bytRegion(3452) = 57
bytRegion(3456) = 46
bytRegion(3460) = 56
bytRegion(3464) = 62
bytRegion(3468) = 57
bytRegion(3472) = 80
bytRegion(3476) = 56
bytRegion(3480) = 71
bytRegion(3481) = 2
bytRegion(3484) = 57
bytRegion(3488) = 12
bytRegion(3492) = 57
bytRegion(3496) = 22
bytRegion(3500) = 58
bytRegion(3504) = 47
bytRegion(3508) = 57
bytRegion(3512) = 63
bytRegion(3516) = 58
bytRegion(3520) = 80
bytRegion(3524) = 57
bytRegion(3528) = 71
bytRegion(3529) = 2
bytRegion(3532) = 58
bytRegion(3536) = 12
bytRegion(3540) = 58
bytRegion(3544) = 22
bytRegion(3548) = 59
bytRegion(3552) = 48
bytRegion(3556) = 58
bytRegion(3560) = 64
bytRegion(3564) = 59
bytRegion(3568) = 81
bytRegion(3572) = 58
bytRegion(3576) = 71
bytRegion(3577) = 2
bytRegion(3580) = 59
bytRegion(3584) = 12
bytRegion(3588) = 59
bytRegion(3592) = 22
bytRegion(3596) = 60
bytRegion(3600) = 49
bytRegion(3604) = 59
bytRegion(3608) = 65
bytRegion(3612) = 60
bytRegion(3616) = 82
bytRegion(3620) = 59
bytRegion(3624) = 71
bytRegion(3625) = 2
bytRegion(3628) = 60
bytRegion(3632) = 12
bytRegion(3636) = 60
bytRegion(3640) = 22
bytRegion(3644) = 61
bytRegion(3648) = 50
bytRegion(3652) = 60
bytRegion(3656) = 67
bytRegion(3660) = 61
bytRegion(3664) = 84
bytRegion(3668) = 60
bytRegion(3672) = 71
bytRegion(3673) = 2
bytRegion(3676) = 61
bytRegion(3680) = 109
bytRegion(3684) = 61
bytRegion(3688) = 71
bytRegion(3689) = 2
bytRegion(3692) = 135
bytRegion(3696) = 100
bytRegion(3700) = 135
bytRegion(3704) = 103
bytRegion(3708) = 136
bytRegion(3712) = 107
bytRegion(3716) = 135
bytRegion(3720) = 71
bytRegion(3721) = 2
bytRegion(3724) = 136
bytRegion(3728) = 97
bytRegion(3732) = 136
bytRegion(3736) = 71
bytRegion(3737) = 2
bytRegion(3740) = 137
bytRegion(3744) = 89
bytRegion(3748) = 137
bytRegion(3752) = 71
bytRegion(3753) = 2
bytRegion(3756) = 138
bytRegion(3760) = 86
bytRegion(3764) = 138
bytRegion(3768) = 71
bytRegion(3769) = 2
bytRegion(3772) = 139
bytRegion(3776) = 83
bytRegion(3780) = 139
bytRegion(3784) = 71
bytRegion(3785) = 2
bytRegion(3788) = 140
bytRegion(3792) = 81
bytRegion(3796) = 140
bytRegion(3800) = 71
bytRegion(3801) = 2
bytRegion(3804) = 141
bytRegion(3808) = 79
bytRegion(3812) = 141
bytRegion(3816) = 71
bytRegion(3817) = 2
bytRegion(3820) = 142
bytRegion(3824) = 77
bytRegion(3828) = 142
bytRegion(3832) = 71
bytRegion(3833) = 2
bytRegion(3836) = 143
bytRegion(3840) = 75
bytRegion(3844) = 143
bytRegion(3848) = 71
bytRegion(3849) = 2
bytRegion(3852) = 144
bytRegion(3856) = 74
bytRegion(3860) = 144
bytRegion(3864) = 71
bytRegion(3865) = 2
bytRegion(3868) = 146
bytRegion(3872) = 73
bytRegion(3876) = 146
bytRegion(3880) = 71
bytRegion(3881) = 2
bytRegion(3884) = 147
bytRegion(3888) = 64
bytRegion(3892) = 147
bytRegion(3896) = 66
bytRegion(3900) = 148
bytRegion(3904) = 72
bytRegion(3908) = 147
bytRegion(3912) = 71
bytRegion(3913) = 2
bytRegion(3916) = 148
bytRegion(3920) = 70
bytRegion(3924) = 148
bytRegion(3928) = 71
bytRegion(3929) = 2
bytRegion(3932) = 149
bytRegion(3936) = 68
bytRegion(3940) = 149
bytRegion(3944) = 71
bytRegion(3945) = 2
bytRegion(3948) = 151
bytRegion(3952) = 67
bytRegion(3956) = 151
bytRegion(3960) = 71
bytRegion(3961) = 2
bytRegion(3964) = 152
bytRegion(3968) = 66
bytRegion(3972) = 152
bytRegion(3976) = 71
bytRegion(3977) = 2
bytRegion(3980) = 154
bytRegion(3984) = 64
bytRegion(3988) = 154
bytRegion(3992) = 71
bytRegion(3993) = 2
bytRegion(3996) = 157
bytRegion(4000) = 53
bytRegion(4004) = 157
bytRegion(4008) = 71
bytRegion(4009) = 2
bytRegion(4012) = 158
bytRegion(4016) = 51
bytRegion(4020) = 158
bytRegion(4024) = 71
bytRegion(4025) = 2
bytRegion(4028) = 159
bytRegion(4032) = 52
bytRegion(4036) = 159
bytRegion(4040) = 71
bytRegion(4041) = 2
bytRegion(4044) = 160
bytRegion(4048) = 50
bytRegion(4052) = 160
bytRegion(4056) = 71
bytRegion(4057) = 2
bytRegion(4060) = 163
bytRegion(4064) = 49
bytRegion(4068) = 163
bytRegion(4072) = 71
bytRegion(4073) = 2
bytRegion(4076) = 164
bytRegion(4080) = 48
bytRegion(4084) = 164
bytRegion(4088) = 71
bytRegion(4089) = 2
bytRegion(4092) = 169
bytRegion(4096) = 47
bytRegion(4100) = 169
bytRegion(4104) = 71
bytRegion(4105) = 2
bytRegion(4108) = 171
bytRegion(4112) = 46
bytRegion(4116) = 171
bytRegion(4120) = 71
bytRegion(4121) = 2
bytRegion(4124) = 173
bytRegion(4128) = 35
bytRegion(4132) = 173
bytRegion(4136) = 37
bytRegion(4140) = 175
bytRegion(4144) = 45
bytRegion(4148) = 173
bytRegion(4152) = 71
bytRegion(4153) = 2
bytRegion(4156) = 175
bytRegion(4160) = 45
bytRegion(4164) = 175
bytRegion(4168) = 71
bytRegion(4169) = 2
bytRegion(4172) = 176
bytRegion(4176) = 44
bytRegion(4180) = 176
bytRegion(4184) = 71
bytRegion(4185) = 2
bytRegion(4188) = 177
bytRegion(4192) = 39
bytRegion(4196) = 177
bytRegion(4200) = 41
bytRegion(4204) = 178
bytRegion(4208) = 43
bytRegion(4212) = 177
bytRegion(4216) = 71
bytRegion(4217) = 2
bytRegion(4220) = 178
bytRegion(4224) = 39
bytRegion(4228) = 178
bytRegion(4232) = 41
bytRegion(4236) = 179
bytRegion(4240) = 42
bytRegion(4244) = 178
bytRegion(4248) = 71
bytRegion(4249) = 2
bytRegion(4252) = 179
bytRegion(4256) = 27
bytRegion(4260) = 179
bytRegion(4264) = 29
bytRegion(4268) = 180
bytRegion(4272) = 41
bytRegion(4276) = 179
bytRegion(4280) = 71
bytRegion(4281) = 2
bytRegion(4284) = 180
bytRegion(4288) = 40
bytRegion(4292) = 180
bytRegion(4296) = 71
bytRegion(4297) = 2
bytRegion(4300) = 181
bytRegion(4304) = 38
bytRegion(4308) = 181
bytRegion(4312) = 71
bytRegion(4313) = 2
bytRegion(4316) = 183
bytRegion(4320) = 36
bytRegion(4324) = 183
bytRegion(4328) = 71
bytRegion(4329) = 2
bytRegion(4332) = 185
bytRegion(4336) = 35
bytRegion(4340) = 185
bytRegion(4344) = 71
bytRegion(4345) = 2
bytRegion(4348) = 186
bytRegion(4352) = 34
bytRegion(4356) = 186
bytRegion(4360) = 71
bytRegion(4361) = 2
bytRegion(4364) = 187
bytRegion(4368) = 9
bytRegion(4372) = 187
bytRegion(4376) = 11
bytRegion(4380) = 189
bytRegion(4384) = 33
bytRegion(4388) = 187
bytRegion(4392) = 71
bytRegion(4393) = 2
bytRegion(4396) = 189
bytRegion(4400) = 32
bytRegion(4404) = 189
bytRegion(4408) = 71
bytRegion(4409) = 2
bytRegion(4412) = 191
bytRegion(4416) = 31
bytRegion(4420) = 191
bytRegion(4424) = 71
bytRegion(4425) = 2
bytRegion(4428) = 193
bytRegion(4432) = 29
bytRegion(4436) = 193
bytRegion(4440) = 71
bytRegion(4441) = 2
bytRegion(4444) = 197
bytRegion(4448) = 28
bytRegion(4452) = 197
bytRegion(4456) = 71
bytRegion(4457) = 2
bytRegion(4460) = 198
bytRegion(4464) = 26
bytRegion(4468) = 198
bytRegion(4472) = 71
bytRegion(4473) = 2
bytRegion(4476) = 199
bytRegion(4480) = 24
bytRegion(4484) = 199
bytRegion(4488) = 71
bytRegion(4489) = 2
bytRegion(4492) = 200
bytRegion(4496) = 23
bytRegion(4500) = 200
bytRegion(4504) = 71
bytRegion(4505) = 2
bytRegion(4508) = 201
bytRegion(4512) = 22
bytRegion(4516) = 201
bytRegion(4520) = 71
bytRegion(4521) = 2
bytRegion(4524) = 204
bytRegion(4528) = 21
bytRegion(4532) = 204
bytRegion(4536) = 71
bytRegion(4537) = 2
bytRegion(4540) = 205
bytRegion(4544) = 20
bytRegion(4548) = 205
bytRegion(4552) = 71
bytRegion(4553) = 2
bytRegion(4556) = 207
bytRegion(4560) = 19
bytRegion(4564) = 207
bytRegion(4568) = 71
bytRegion(4569) = 2
bytRegion(4572) = 209
bytRegion(4576) = 18
bytRegion(4580) = 209
bytRegion(4584) = 71
bytRegion(4585) = 2
bytRegion(4588) = 210
bytRegion(4592) = 17
bytRegion(4596) = 210
bytRegion(4600) = 71
bytRegion(4601) = 2
bytRegion(4604) = 211
bytRegion(4608) = 16
bytRegion(4612) = 211
bytRegion(4616) = 71
bytRegion(4617) = 2
bytRegion(4620) = 213
bytRegion(4624) = 15
bytRegion(4628) = 213
bytRegion(4632) = 71
bytRegion(4633) = 2
bytRegion(4636) = 215
bytRegion(4640) = 14
bytRegion(4644) = 215
bytRegion(4648) = 71
bytRegion(4649) = 2
bytRegion(4652) = 217
bytRegion(4656) = 13
bytRegion(4660) = 217
bytRegion(4664) = 71
bytRegion(4665) = 2
bytRegion(4668) = 220
bytRegion(4672) = 12
bytRegion(4676) = 220
bytRegion(4680) = 71
bytRegion(4681) = 2
bytRegion(4684) = 223
bytRegion(4688) = 11
bytRegion(4692) = 223
bytRegion(4696) = 71
bytRegion(4697) = 2
bytRegion(4700) = 224
bytRegion(4704) = 10
bytRegion(4708) = 224
bytRegion(4712) = 71
bytRegion(4713) = 2
bytRegion(4716) = 227
bytRegion(4720) = 9
bytRegion(4724) = 227
bytRegion(4728) = 71
bytRegion(4729) = 2
bytRegion(4732) = 232
bytRegion(4736) = 10
bytRegion(4740) = 232
bytRegion(4744) = 71
bytRegion(4745) = 2
bytRegion(4748) = 237
bytRegion(4752) = 11
bytRegion(4756) = 237
bytRegion(4760) = 71
bytRegion(4761) = 2
bytRegion(4764) = 240
bytRegion(4768) = 12
bytRegion(4772) = 240
bytRegion(4776) = 71
bytRegion(4777) = 2
bytRegion(4780) = 245
bytRegion(4784) = 13
bytRegion(4788) = 245
bytRegion(4792) = 71
bytRegion(4793) = 2
bytRegion(4796) = 249
bytRegion(4800) = 14
bytRegion(4804) = 249
bytRegion(4808) = 71
bytRegion(4809) = 2
bytRegion(4812) = 251
bytRegion(4816) = 15
bytRegion(4820) = 251
bytRegion(4824) = 71
bytRegion(4825) = 2
bytRegion(4828) = 253
bytRegion(4832) = 16
bytRegion(4836) = 253
bytRegion(4840) = 71
bytRegion(4841) = 2
bytRegion(4844) = 1
bytRegion(4845) = 1
bytRegion(4848) = 17
bytRegion(4852) = 1
bytRegion(4853) = 1
bytRegion(4856) = 71
bytRegion(4857) = 2
bytRegion(4860) = 3
bytRegion(4861) = 1
bytRegion(4864) = 18
bytRegion(4868) = 3
bytRegion(4869) = 1
bytRegion(4872) = 71
bytRegion(4873) = 2
bytRegion(4876) = 6
bytRegion(4877) = 1
bytRegion(4880) = 19
bytRegion(4884) = 6
bytRegion(4885) = 1
bytRegion(4888) = 71
bytRegion(4889) = 2
bytRegion(4892) = 7
bytRegion(4893) = 1
bytRegion(4896) = 20
bytRegion(4900) = 7
bytRegion(4901) = 1
bytRegion(4904) = 71
bytRegion(4905) = 2
bytRegion(4908) = 10
bytRegion(4909) = 1
bytRegion(4912) = 21
bytRegion(4916) = 10
bytRegion(4917) = 1
bytRegion(4920) = 71
bytRegion(4921) = 2
bytRegion(4924) = 12
bytRegion(4925) = 1
bytRegion(4928) = 22
bytRegion(4932) = 12
bytRegion(4933) = 1
bytRegion(4936) = 71
bytRegion(4937) = 2
bytRegion(4940) = 14
bytRegion(4941) = 1
bytRegion(4944) = 23
bytRegion(4948) = 14
bytRegion(4949) = 1
bytRegion(4952) = 71
bytRegion(4953) = 2
bytRegion(4956) = 16
bytRegion(4957) = 1
bytRegion(4960) = 25
bytRegion(4964) = 16
bytRegion(4965) = 1
bytRegion(4968) = 71
bytRegion(4969) = 2
bytRegion(4972) = 17
bytRegion(4973) = 1
bytRegion(4976) = 26
bytRegion(4980) = 17
bytRegion(4981) = 1
bytRegion(4984) = 71
bytRegion(4985) = 2
bytRegion(4988) = 19
bytRegion(4989) = 1
bytRegion(4992) = 27
bytRegion(4996) = 19
bytRegion(4997) = 1
bytRegion(5000) = 71
bytRegion(5001) = 2
bytRegion(5004) = 21
bytRegion(5005) = 1
bytRegion(5008) = 28
bytRegion(5012) = 21
bytRegion(5013) = 1
bytRegion(5016) = 71
bytRegion(5017) = 2
bytRegion(5020) = 23
bytRegion(5021) = 1
bytRegion(5024) = 29
bytRegion(5028) = 23
bytRegion(5029) = 1
bytRegion(5032) = 71
bytRegion(5033) = 2
bytRegion(5036) = 24
bytRegion(5037) = 1
bytRegion(5040) = 30
bytRegion(5044) = 24
bytRegion(5045) = 1
bytRegion(5048) = 71
bytRegion(5049) = 2
bytRegion(5052) = 26
bytRegion(5053) = 1
bytRegion(5056) = 31
bytRegion(5060) = 26
bytRegion(5061) = 1
bytRegion(5064) = 71
bytRegion(5065) = 2
bytRegion(5068) = 42
bytRegion(5069) = 1
bytRegion(5072) = 33
bytRegion(5076) = 42
bytRegion(5077) = 1
bytRegion(5080) = 71
bytRegion(5081) = 2
bytRegion(5084) = 44
bytRegion(5085) = 1
bytRegion(5088) = 34
bytRegion(5092) = 44
bytRegion(5093) = 1
bytRegion(5096) = 71
bytRegion(5097) = 2
bytRegion(5100) = 45
bytRegion(5101) = 1
bytRegion(5104) = 37
bytRegion(5108) = 45
bytRegion(5109) = 1
bytRegion(5112) = 71
bytRegion(5113) = 2
bytRegion(5116) = 46
bytRegion(5117) = 1
bytRegion(5120) = 42
bytRegion(5124) = 46
bytRegion(5125) = 1
bytRegion(5128) = 71
bytRegion(5129) = 2
bytRegion(5132) = 47
bytRegion(5133) = 1
bytRegion(5136) = 48
bytRegion(5140) = 47
bytRegion(5141) = 1
bytRegion(5144) = 71
bytRegion(5145) = 2
bytRegion(5148) = 48
bytRegion(5149) = 1
bytRegion(5152) = 49
bytRegion(5156) = 48
bytRegion(5157) = 1
bytRegion(5160) = 71
bytRegion(5161) = 2
bytRegion(5164) = 52
bytRegion(5165) = 1
bytRegion(5168) = 48
bytRegion(5172) = 52
bytRegion(5173) = 1
bytRegion(5176) = 71
bytRegion(5177) = 2
bytRegion(5180) = 55
bytRegion(5181) = 1
bytRegion(5184) = 47
bytRegion(5188) = 55
bytRegion(5189) = 1
bytRegion(5192) = 71
bytRegion(5193) = 2
bytRegion(5196) = 57
bytRegion(5197) = 1
bytRegion(5200) = 48
bytRegion(5204) = 57
bytRegion(5205) = 1
bytRegion(5208) = 71
bytRegion(5209) = 2
bytRegion(5212) = 59
bytRegion(5213) = 1
bytRegion(5216) = 47
bytRegion(5220) = 59
bytRegion(5221) = 1
bytRegion(5224) = 71
bytRegion(5225) = 2
bytRegion(5228) = 62
bytRegion(5229) = 1
bytRegion(5232) = 46
bytRegion(5236) = 62
bytRegion(5237) = 1
bytRegion(5240) = 71
bytRegion(5241) = 2
bytRegion(5244) = 64
bytRegion(5245) = 1
bytRegion(5248) = 45
bytRegion(5252) = 64
bytRegion(5253) = 1
bytRegion(5256) = 70
bytRegion(5257) = 2
bytRegion(5260) = 66
bytRegion(5261) = 1
bytRegion(5264) = 21
bytRegion(5268) = 66
bytRegion(5269) = 1
bytRegion(5272) = 23
bytRegion(5276) = 67
bytRegion(5277) = 1
bytRegion(5280) = 44
bytRegion(5284) = 66
bytRegion(5285) = 1
bytRegion(5288) = 69
bytRegion(5289) = 2
bytRegion(5292) = 67
bytRegion(5293) = 1
bytRegion(5296) = 21
bytRegion(5300) = 67
bytRegion(5301) = 1
bytRegion(5304) = 22
bytRegion(5308) = 68
bytRegion(5309) = 1
bytRegion(5312) = 44
bytRegion(5316) = 67
bytRegion(5317) = 1
bytRegion(5320) = 68
bytRegion(5321) = 2
bytRegion(5324) = 68
bytRegion(5325) = 1
bytRegion(5328) = 44
bytRegion(5332) = 68
bytRegion(5333) = 1
bytRegion(5336) = 66
bytRegion(5337) = 2
bytRegion(5340) = 69
bytRegion(5341) = 1
bytRegion(5344) = 44
bytRegion(5348) = 69
bytRegion(5349) = 1
bytRegion(5352) = 199
bytRegion(5356) = 70
bytRegion(5357) = 1
bytRegion(5360) = 43
bytRegion(5364) = 70
bytRegion(5365) = 1
bytRegion(5368) = 199
bytRegion(5372) = 74
bytRegion(5373) = 1
bytRegion(5376) = 42
bytRegion(5380) = 74
bytRegion(5381) = 1
bytRegion(5384) = 199
bytRegion(5388) = 76
bytRegion(5389) = 1
bytRegion(5392) = 42
bytRegion(5396) = 76
bytRegion(5397) = 1
bytRegion(5400) = 200
bytRegion(5404) = 79
bytRegion(5405) = 1
bytRegion(5408) = 41
bytRegion(5412) = 79
bytRegion(5413) = 1
bytRegion(5416) = 200
bytRegion(5420) = 80
bytRegion(5421) = 1
bytRegion(5424) = 40
bytRegion(5428) = 80
bytRegion(5429) = 1
bytRegion(5432) = 200
bytRegion(5436) = 82
bytRegion(5437) = 1
bytRegion(5440) = 39
bytRegion(5444) = 82
bytRegion(5445) = 1
bytRegion(5448) = 200
bytRegion(5452) = 83
bytRegion(5453) = 1
bytRegion(5456) = 39
bytRegion(5460) = 83
bytRegion(5461) = 1
bytRegion(5464) = 201
bytRegion(5468) = 84
bytRegion(5469) = 1
bytRegion(5472) = 41
bytRegion(5476) = 84
bytRegion(5477) = 1
bytRegion(5480) = 201
bytRegion(5484) = 85
bytRegion(5485) = 1
bytRegion(5488) = 41
bytRegion(5492) = 85
bytRegion(5493) = 1
bytRegion(5496) = 202
bytRegion(5500) = 86
bytRegion(5501) = 1
bytRegion(5504) = 40
bytRegion(5508) = 86
bytRegion(5509) = 1
bytRegion(5512) = 202
bytRegion(5516) = 88
bytRegion(5517) = 1
bytRegion(5520) = 41
bytRegion(5524) = 88
bytRegion(5525) = 1
bytRegion(5528) = 203
bytRegion(5532) = 90
bytRegion(5533) = 1
bytRegion(5536) = 42
bytRegion(5540) = 90
bytRegion(5541) = 1
bytRegion(5544) = 203
bytRegion(5548) = 92
bytRegion(5549) = 1
End Sub
