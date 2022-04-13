Imports System.Data.OleDb

Public Class Form1
    Dim Index, CRC_Low, CRC_High, CRCTable(0 To 511) As Byte
    Dim Response_mod
    Dim address(39), silent(125), cek(39), J, jumlahsalah, opo, flag, slaveDEISV, DOn, ite, KL, opo1, save5, save6, save7, save8, jumlahsalah1 As Integer '= {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
    Dim Slave, slave1, poll, ID(8), var1, var2, var3, inc, cek1, count(38), n, m, rx, inc1, var4, var5, var6, quantitydei, startadd, startadd1, quantity1, iterasi, tes1, tes2, hari As Integer
    Dim providerTZ, dataFileTZ, connStringTZ, strg1, strg2, kp, kp1, status, Command_Head, Bcode, command1, BCC, command2, tes, start1, finish1 As String
    Dim var10(30), var11(30), digout(40) As String
    Dim time As TimeSpan

    Private Sub CRC_16(ByVal Data As String, Length As Integer)
        Dim i As Integer
        Dim Index As Byte
        On Error Resume Next
        CRC_Low = &HFF
        CRC_High = &HFF

        For i = 1 To Length
            Index = CRC_High Xor Asc(Mid(Data, i, 1))
            CRC_High = CRC_Low Xor CRCTable(Index)
            CRC_Low = CRCTable(Index + 256)
        Next i
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Slave = 100 'address awal untuk pembacaan data Pin UI (DEI)
        slave1 = 100 'address awal untuk pembacaan data pin DO (DEI)
        kp = "0000000000000000000000000000" 'menunjukan address berapa saja yang aktif dikemas dalam bilangan biner, dimana setiap bit menunjukkan satu address (DEI)
        kp1 = "000000000000000" 'menunjukan address berapa saja yang aktif dikemas dalam bilangan biner, dimana setiap bit menunjukkan satu address (TZ)
        poll = 0 'address awal TZ
        'inc = 210 'inc untuk iterasi memberikan efek kedip pada alarm (DEI)
        save6 = 1
        save5 = 0
        n = 0
        inc1 = 1
        flag = 0
        bacaON()
        bacaOFF()
        'slaveDEISV = 100
        opo1 = 1
        'opo = 0
        'range kiri 336
        'range kanan 335
        'spih =340
        'spic =338
        'DEI
        MSComm1.Settings = "9600,o,8,1"
        MSComm1.CommPort = 4
        MSComm1.PortOpen = True

        'tz
        MSComm2.Settings = "4800,n,8,1"
        MSComm2.CommPort = 5
        MSComm2.PortOpen = True

        'tk
        'MSComm3.Settings = "4800,n,8,2"
        'MSComm3.CommPort = 4
        'MSComm3.PortOpen = True

        ComboBox2.Items.Add("SEMUA RUANGAN")
        ComboBox2.Items.Add("POS 3")
        ComboBox2.Items.Add("PPIC")
        ComboBox2.Items.Add("2")
        ComboBox2.Items.Add("3")
        ComboBox2.Items.Add("Emulsi 2")
        ComboBox2.Items.Add("Emulsi 1")
        ComboBox2.Items.Add("POS 4")
        ComboBox2.Items.Add("INO")
        ComboBox2.Items.Add("Panen 4")
        ComboBox2.Items.Add("Panen 3")
        ComboBox2.Items.Add("Koridor")
        ComboBox2.Items.Add("INO 3")
        ComboBox2.Items.Add("PRA INO 2")
        ComboBox2.Items.Add("Passage")
        ComboBox2.Items.Add("PRA INO 1")
        ComboBox2.Items.Add("Alat Steril")
        ComboBox2.Items.Add("STERILISASI")
        ComboBox2.Items.Add("Persiapan")
        ComboBox2.Items.Add("INO 2")
        ComboBox2.Items.Add("INO 1")
        ComboBox2.Items.Add("POS INO 2")
        ComboBox2.Items.Add("POS INO 1")
        ComboBox2.Items.Add("Panen 1")
        ComboBox2.Items.Add("Panen 2")
        ComboBox2.Items.Add("Inaktif 1")
        ComboBox2.Items.Add("Inaktif 2")

        'startadd = 0
        'quantitydei = 8
        'startadd1 = 22
        'quantity1 = 8
        'startaddlabel.text = 0
        'quantitylabel.text = 8
        quantitylabel.Text = 4
        startaddlabel.Text = 0
        quantitylabel1.Text = 8
        startaddlabel1.Text = 22
        'baru.dei_startAdd1.Text = 22
        'baru.dei_quantity1.Text = 8

        CRCTable(0) = &H0
        CRCTable(1) = &HC1
        CRCTable(2) = &H81
        CRCTable(3) = &H40
        CRCTable(4) = &H1
        CRCTable(5) = &HC0
        CRCTable(6) = &H80
        CRCTable(7) = &H41
        CRCTable(8) = &H1
        CRCTable(9) = &HC0
        CRCTable(10) = &H80
        CRCTable(11) = &H41
        CRCTable(12) = &H0
        CRCTable(13) = &HC1
        CRCTable(14) = &H81
        CRCTable(15) = &H40
        CRCTable(16) = &H1
        CRCTable(17) = &HC0
        CRCTable(18) = &H80
        CRCTable(19) = &H41
        CRCTable(20) = &H0
        CRCTable(21) = &HC1
        CRCTable(22) = &H81
        CRCTable(23) = &H40
        CRCTable(24) = &H0
        CRCTable(25) = &HC1
        CRCTable(26) = &H81
        CRCTable(27) = &H40
        CRCTable(28) = &H1
        CRCTable(29) = &HC0
        CRCTable(30) = &H80
        CRCTable(31) = &H41
        CRCTable(32) = &H1
        CRCTable(33) = &HC0
        CRCTable(34) = &H80
        CRCTable(35) = &H41
        CRCTable(36) = &H0
        CRCTable(37) = &HC1
        CRCTable(38) = &H81
        CRCTable(39) = &H40
        CRCTable(40) = &H0
        CRCTable(41) = &HC1
        CRCTable(42) = &H81
        CRCTable(43) = &H40
        CRCTable(44) = &H1
        CRCTable(45) = &HC0
        CRCTable(46) = &H80
        CRCTable(47) = &H41
        CRCTable(48) = &H0
        CRCTable(49) = &HC1
        CRCTable(50) = &H81
        CRCTable(51) = &H40
        CRCTable(52) = &H1
        CRCTable(53) = &HC0
        CRCTable(54) = &H80
        CRCTable(55) = &H41
        CRCTable(56) = &H1
        CRCTable(57) = &HC0
        CRCTable(58) = &H80
        CRCTable(59) = &H41
        CRCTable(60) = &H0
        CRCTable(61) = &HC1
        CRCTable(62) = &H81
        CRCTable(63) = &H40
        CRCTable(64) = &H1
        CRCTable(65) = &HC0
        CRCTable(66) = &H80
        CRCTable(67) = &H41
        CRCTable(68) = &H0
        CRCTable(69) = &HC1
        CRCTable(70) = &H81
        CRCTable(71) = &H40
        CRCTable(72) = &H0
        CRCTable(73) = &HC1
        CRCTable(74) = &H81
        CRCTable(75) = &H40
        CRCTable(76) = &H1
        CRCTable(77) = &HC0
        CRCTable(78) = &H80
        CRCTable(79) = &H41
        CRCTable(80) = &H0
        CRCTable(81) = &HC1
        CRCTable(82) = &H81
        CRCTable(83) = &H40
        CRCTable(84) = &H1
        CRCTable(85) = &HC0
        CRCTable(86) = &H80
        CRCTable(87) = &H41
        CRCTable(88) = &H1
        CRCTable(89) = &HC0
        CRCTable(90) = &H80
        CRCTable(91) = &H41
        CRCTable(92) = &H0
        CRCTable(93) = &HC1
        CRCTable(94) = &H81
        CRCTable(95) = &H40
        CRCTable(96) = &H0
        CRCTable(97) = &HC1
        CRCTable(98) = &H81
        CRCTable(99) = &H40
        CRCTable(100) = &H1
        CRCTable(101) = &HC0
        CRCTable(102) = &H80
        CRCTable(103) = &H41
        CRCTable(104) = &H1
        CRCTable(105) = &HC0
        CRCTable(106) = &H80
        CRCTable(107) = &H41
        CRCTable(108) = &H0
        CRCTable(109) = &HC1
        CRCTable(110) = &H81
        CRCTable(111) = &H40
        CRCTable(112) = &H1
        CRCTable(113) = &HC0
        CRCTable(114) = &H80
        CRCTable(115) = &H41
        CRCTable(116) = &H0
        CRCTable(117) = &HC1
        CRCTable(118) = &H81
        CRCTable(119) = &H40
        CRCTable(120) = &H0
        CRCTable(121) = &HC1
        CRCTable(122) = &H81
        CRCTable(123) = &H40
        CRCTable(124) = &H1
        CRCTable(125) = &HC0
        CRCTable(126) = &H80
        CRCTable(127) = &H41
        CRCTable(128) = &H1
        CRCTable(129) = &HC0
        CRCTable(130) = &H80
        CRCTable(131) = &H41
        CRCTable(132) = &H0
        CRCTable(133) = &HC1
        CRCTable(134) = &H81
        CRCTable(135) = &H40
        CRCTable(136) = &H0
        CRCTable(137) = &HC1
        CRCTable(138) = &H81
        CRCTable(139) = &H40
        CRCTable(140) = &H1
        CRCTable(141) = &HC0
        CRCTable(142) = &H80
        CRCTable(143) = &H41
        CRCTable(144) = &H0
        CRCTable(145) = &HC1
        CRCTable(146) = &H81
        CRCTable(147) = &H40
        CRCTable(148) = &H1
        CRCTable(149) = &HC0
        CRCTable(150) = &H80
        CRCTable(151) = &H41
        CRCTable(152) = &H1
        CRCTable(153) = &HC0
        CRCTable(154) = &H80
        CRCTable(155) = &H41
        CRCTable(156) = &H0
        CRCTable(157) = &HC1
        CRCTable(158) = &H81
        CRCTable(159) = &H40
        CRCTable(160) = &H0
        CRCTable(161) = &HC1
        CRCTable(162) = &H81
        CRCTable(163) = &H40
        CRCTable(164) = &H1
        CRCTable(165) = &HC0
        CRCTable(166) = &H80
        CRCTable(167) = &H41
        CRCTable(168) = &H1
        CRCTable(169) = &HC0
        CRCTable(170) = &H80
        CRCTable(171) = &H41
        CRCTable(172) = &H0
        CRCTable(173) = &HC1
        CRCTable(174) = &H81
        CRCTable(175) = &H40
        CRCTable(176) = &H1
        CRCTable(177) = &HC0
        CRCTable(178) = &H80
        CRCTable(179) = &H41
        CRCTable(180) = &H0
        CRCTable(181) = &HC1
        CRCTable(182) = &H81
        CRCTable(183) = &H40
        CRCTable(184) = &H0
        CRCTable(185) = &HC1
        CRCTable(186) = &H81
        CRCTable(187) = &H40
        CRCTable(188) = &H1
        CRCTable(189) = &HC0
        CRCTable(190) = &H80
        CRCTable(191) = &H41
        CRCTable(192) = &H0
        CRCTable(193) = &HC1
        CRCTable(194) = &H81
        CRCTable(195) = &H40
        CRCTable(196) = &H1
        CRCTable(197) = &HC0
        CRCTable(198) = &H80
        CRCTable(199) = &H41
        CRCTable(200) = &H1
        CRCTable(201) = &HC0
        CRCTable(202) = &H80
        CRCTable(203) = &H41
        CRCTable(204) = &H0
        CRCTable(205) = &HC1
        CRCTable(206) = &H81
        CRCTable(207) = &H40
        CRCTable(208) = &H1
        CRCTable(209) = &HC0
        CRCTable(210) = &H80
        CRCTable(211) = &H41
        CRCTable(212) = &H0
        CRCTable(213) = &HC1
        CRCTable(214) = &H81
        CRCTable(215) = &H40
        CRCTable(216) = &H0
        CRCTable(217) = &HC1
        CRCTable(218) = &H81
        CRCTable(219) = &H40
        CRCTable(220) = &H1
        CRCTable(221) = &HC0
        CRCTable(222) = &H80
        CRCTable(223) = &H41
        CRCTable(224) = &H1
        CRCTable(225) = &HC0
        CRCTable(226) = &H80
        CRCTable(227) = &H41
        CRCTable(228) = &H0
        CRCTable(229) = &HC1
        CRCTable(230) = &H81
        CRCTable(231) = &H40
        CRCTable(232) = &H0
        CRCTable(233) = &HC1
        CRCTable(234) = &H81
        CRCTable(235) = &H40
        CRCTable(236) = &H1
        CRCTable(237) = &HC0
        CRCTable(238) = &H80
        CRCTable(239) = &H41
        CRCTable(240) = &H0
        CRCTable(241) = &HC1
        CRCTable(242) = &H81
        CRCTable(243) = &H40
        CRCTable(244) = &H1
        CRCTable(245) = &HC0
        CRCTable(246) = &H80
        CRCTable(247) = &H41
        CRCTable(248) = &H1
        CRCTable(249) = &HC0
        CRCTable(250) = &H80
        CRCTable(251) = &H41
        CRCTable(252) = &H0
        CRCTable(253) = &HC1
        CRCTable(254) = &H81
        CRCTable(255) = &H40
        CRCTable(256) = &H0
        CRCTable(257) = &HC0
        CRCTable(258) = &HC1
        CRCTable(259) = &H1
        CRCTable(260) = &HC3
        CRCTable(261) = &H3
        CRCTable(262) = &H2
        CRCTable(263) = &HC2
        CRCTable(264) = &HC6
        CRCTable(265) = &H6
        CRCTable(266) = &H7
        CRCTable(267) = &HC7
        CRCTable(268) = &H5
        CRCTable(269) = &HC5
        CRCTable(270) = &HC4
        CRCTable(271) = &H4
        CRCTable(272) = &HCC
        CRCTable(273) = &HC
        CRCTable(274) = &HD
        CRCTable(275) = &HCD
        CRCTable(276) = &HF
        CRCTable(277) = &HCF
        CRCTable(278) = &HCE
        CRCTable(279) = &HE
        CRCTable(280) = &HA
        CRCTable(281) = &HCA
        CRCTable(282) = &HCB
        CRCTable(283) = &HB
        CRCTable(284) = &HC9
        CRCTable(285) = &H9
        CRCTable(286) = &H8
        CRCTable(287) = &HC8
        CRCTable(288) = &HD8
        CRCTable(289) = &H18
        CRCTable(290) = &H19
        CRCTable(291) = &HD9
        CRCTable(292) = &H1B
        CRCTable(293) = &HDB
        CRCTable(294) = &HDA
        CRCTable(295) = &H1A
        CRCTable(296) = &H1E
        CRCTable(297) = &HDE
        CRCTable(298) = &HDF
        CRCTable(299) = &H1F
        CRCTable(300) = &HDD
        CRCTable(301) = &H1D
        CRCTable(302) = &H1C
        CRCTable(303) = &HDC
        CRCTable(304) = &H14
        CRCTable(305) = &HD4
        CRCTable(306) = &HD5
        CRCTable(307) = &H15
        CRCTable(308) = &HD7
        CRCTable(309) = &H17
        CRCTable(310) = &H16
        CRCTable(311) = &HD6
        CRCTable(312) = &HD2
        CRCTable(313) = &H12
        CRCTable(314) = &H13
        CRCTable(315) = &HD3
        CRCTable(316) = &H11
        CRCTable(317) = &HD1
        CRCTable(318) = &HD0
        CRCTable(319) = &H10
        CRCTable(320) = &HF0
        CRCTable(321) = &H30
        CRCTable(322) = &H31
        CRCTable(323) = &HF1
        CRCTable(324) = &H33
        CRCTable(325) = &HF3
        CRCTable(326) = &HF2
        CRCTable(327) = &H32
        CRCTable(328) = &H36
        CRCTable(329) = &HF6
        CRCTable(330) = &HF7
        CRCTable(331) = &H37
        CRCTable(332) = &HF5
        CRCTable(333) = &H35
        CRCTable(334) = &H34
        CRCTable(335) = &HF4
        CRCTable(336) = &H3C
        CRCTable(337) = &HFC
        CRCTable(338) = &HFD
        CRCTable(339) = &H3D
        CRCTable(340) = &HFF
        CRCTable(341) = &H3F
        CRCTable(342) = &H3E
        CRCTable(343) = &HFE
        CRCTable(344) = &HFA
        CRCTable(345) = &H3A
        CRCTable(346) = &H3B
        CRCTable(347) = &HFB
        CRCTable(348) = &H39
        CRCTable(349) = &HF9
        CRCTable(350) = &HF8
        CRCTable(351) = &H38
        CRCTable(352) = &H28
        CRCTable(353) = &HE8
        CRCTable(354) = &HE9
        CRCTable(355) = &H29
        CRCTable(356) = &HEB
        CRCTable(357) = &H2B
        CRCTable(358) = &H2A
        CRCTable(359) = &HEA
        CRCTable(360) = &HEE
        CRCTable(361) = &H2E
        CRCTable(362) = &H2F
        CRCTable(363) = &HEF
        CRCTable(364) = &H2D
        CRCTable(365) = &HED
        CRCTable(366) = &HEC
        CRCTable(367) = &H2C
        CRCTable(368) = &HE4
        CRCTable(369) = &H24
        CRCTable(370) = &H25
        CRCTable(371) = &HE5
        CRCTable(372) = &H27
        CRCTable(373) = &HE7
        CRCTable(374) = &HE6
        CRCTable(375) = &H26
        CRCTable(376) = &H22
        CRCTable(377) = &HE2
        CRCTable(378) = &HE3
        CRCTable(379) = &H23
        CRCTable(380) = &HE1
        CRCTable(381) = &H21
        CRCTable(382) = &H20
        CRCTable(383) = &HE0
        CRCTable(384) = &HA0
        CRCTable(385) = &H60
        CRCTable(386) = &H61
        CRCTable(387) = &HA1
        CRCTable(388) = &H63
        CRCTable(389) = &HA3
        CRCTable(390) = &HA2
        CRCTable(391) = &H62
        CRCTable(392) = &H66
        CRCTable(393) = &HA6
        CRCTable(394) = &HA7
        CRCTable(395) = &H67
        CRCTable(396) = &HA5
        CRCTable(397) = &H65
        CRCTable(398) = &H64
        CRCTable(399) = &HA4
        CRCTable(400) = &H6C
        CRCTable(401) = &HAC
        CRCTable(402) = &HAD
        CRCTable(403) = &H6D
        CRCTable(404) = &HAF
        CRCTable(405) = &H6F
        CRCTable(406) = &H6E
        CRCTable(407) = &HAE
        CRCTable(408) = &HAA
        CRCTable(409) = &H6A
        CRCTable(410) = &H6B
        CRCTable(411) = &HAB
        CRCTable(412) = &H69
        CRCTable(413) = &HA9
        CRCTable(414) = &HA8
        CRCTable(415) = &H68
        CRCTable(416) = &H78
        CRCTable(417) = &HB8
        CRCTable(418) = &HB9
        CRCTable(419) = &H79
        CRCTable(420) = &HBB
        CRCTable(421) = &H7B
        CRCTable(422) = &H7A
        CRCTable(423) = &HBA
        CRCTable(424) = &HBE
        CRCTable(425) = &H7E
        CRCTable(426) = &H7F
        CRCTable(427) = &HBF
        CRCTable(428) = &H7D
        CRCTable(429) = &HBD
        CRCTable(430) = &HBC
        CRCTable(431) = &H7C
        CRCTable(432) = &HB4
        CRCTable(433) = &H74
        CRCTable(434) = &H75
        CRCTable(435) = &HB5
        CRCTable(436) = &H77
        CRCTable(437) = &HB7
        CRCTable(438) = &HB6
        CRCTable(439) = &H76
        CRCTable(440) = &H72
        CRCTable(441) = &HB2
        CRCTable(442) = &HB3
        CRCTable(443) = &H73
        CRCTable(444) = &HB1
        CRCTable(445) = &H71
        CRCTable(446) = &H70
        CRCTable(447) = &HB0
        CRCTable(448) = &H50
        CRCTable(449) = &H90
        CRCTable(450) = &H91
        CRCTable(451) = &H51
        CRCTable(452) = &H93
        CRCTable(453) = &H53
        CRCTable(454) = &H52
        CRCTable(455) = &H92
        CRCTable(456) = &H96
        CRCTable(457) = &H56
        CRCTable(458) = &H57
        CRCTable(459) = &H97
        CRCTable(460) = &H55
        CRCTable(461) = &H95
        CRCTable(462) = &H94
        CRCTable(463) = &H54
        CRCTable(464) = &H9C
        CRCTable(465) = &H5C
        CRCTable(466) = &H5D
        CRCTable(467) = &H9D
        CRCTable(468) = &H5F
        CRCTable(469) = &H9F
        CRCTable(470) = &H9E
        CRCTable(471) = &H5E
        CRCTable(472) = &H5A
        CRCTable(473) = &H9A
        CRCTable(474) = &H9B
        CRCTable(475) = &H5B
        CRCTable(476) = &H99
        CRCTable(477) = &H59
        CRCTable(478) = &H58
        CRCTable(479) = &H98
        CRCTable(480) = &H88
        CRCTable(481) = &H48
        CRCTable(482) = &H49
        CRCTable(483) = &H89
        CRCTable(484) = &H4B
        CRCTable(485) = &H8B
        CRCTable(486) = &H8A
        CRCTable(487) = &H4A
        CRCTable(488) = &H4E
        CRCTable(489) = &H8E
        CRCTable(490) = &H8F
        CRCTable(491) = &H4F
        CRCTable(492) = &H8D
        CRCTable(493) = &H4D
        CRCTable(494) = &H4C
        CRCTable(495) = &H8C
        CRCTable(496) = &H44
        CRCTable(497) = &H84
        CRCTable(498) = &H85
        CRCTable(499) = &H45
        CRCTable(500) = &H87
        CRCTable(501) = &H47
        CRCTable(502) = &H46
        CRCTable(503) = &H86
        CRCTable(504) = &H82
        CRCTable(505) = &H42
        CRCTable(506) = &H43
        CRCTable(507) = &H83
        CRCTable(508) = &H41
        CRCTable(509) = &H81
        CRCTable(510) = &H80
        CRCTable(511) = &H40
    End Sub
    Private Sub DEI_timer_Tick_1(sender As Object, e As EventArgs) Handles DEI_timer.Tick
        Dim data As String
        Dim t1 As Long
        Dim k, i, save1, save2, save3, save4 As Integer
        save1 = 0
        save2 = 0
        save3 = 1
        save4 = 1
        If Slave = 126 Then Slave = 100
        'Slave = 108
        On Error Resume Next
        t1 = (CLng(197) * 256)
        data = Chr(Slave) + Chr(3) + Chr(Val(startaddlabel.Text) \ 256) + Chr((startaddlabel.Text) Mod 256) + Chr(0) + Chr(Val(quantitylabel.Text))
        CRC_16(data, 6)
        data = data + Chr(CRC_High) + Chr(CRC_Low)
        Dim PauseTime, Start, Finish
        baru.dei_response.Text = ""
        MSComm1.InputLen = 0
        baru.dei_request.Text = Val(Asc(Mid(data, 1, 3))) & Val(Asc(Mid(data, 2, 1))) & Val(Asc(Mid(data, 3, 1))) _
        & (Val(Asc(Mid(data, 4, 1))) + (Val(Asc(Mid(data, 5, 1))))) _
        & (Val(Asc(Mid(data, 6, 1))) + (Val(Asc(Mid(data, 7, 1))))) _
        & (Mid(data, 8, 1)) & (Mid(data, 9, 1))
        'baru.dei_request.Text = Hex(Asc(Mid(data, 1, 3))) & Hex(Asc(Mid(data, 2, 1))) & Hex(Asc(Mid(data, 3, 1))) _
        '& (Hex(Asc(Mid(data, 4, 1))) + (Hex(Asc(Mid(data, 5, 1))))) _
        '& (Hex(Asc(Mid(data, 6, 1))) + (Hex(Asc(Mid(data, 7, 1))))) _
        '& (Mid(data, 8, 1)) & (Mid(data, 9, 1))
        MSComm1.Output = data

        Do While MSComm1.OutBufferCount > 0
        Loop
        PauseTime = 0.1  ' Set duration.
        Start = Microsoft.VisualBasic.Timer    ' Set start time.
        Do While (Microsoft.VisualBasic.Timer < Start + PauseTime) And (MSComm1.InBufferCount < Val(quantitylabel.Text) * 2 + 5)
            Application.DoEvents()
        Loop

        baru.dei_inBuffer.Text = (MSComm1.InBufferCount)
        'quantitylabel.text = quantitydei
        'startaddlabel.text = startadd
        Response_mod = MSComm1.Input

        baru.dei_response.Text = Val(Asc(Mid(Response_mod, 1, 1))) _
        & Val(Asc(Mid(Response_mod, 2, 1))) _
        & Val(Asc(Mid(Response_mod, 3, 1))) _
        & Val(Asc(Mid(Response_mod, 4, 1))) * CLng(256) + (Val(Asc(Mid(Response_mod, 5, 1)))) _
        & Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + (Val(Asc(Mid(Response_mod, 7, 1)))) _
        & Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + (Val(Asc(Mid(Response_mod, 9, 1)))) _
        & Val(Asc(Mid(Response_mod, 10, 1))) * CLng(256) + (Val(Asc(Mid(Response_mod, 11, 1)))) _
        & (Mid(Response_mod, 11, 1)) & (Mid(Response_mod, 12, 1))

        TextBox1.Text = Hex(Asc(Mid(Response_mod, 1, 1))) _
         & " " _
         & Hex(Asc(Mid(Response_mod, 2, 1))) _
         & " " _
         & Hex(Asc(Mid(Response_mod, 3, 1))) _
         & " " _
         & Hex(Asc(Mid(Response_mod, 4, 1))) _
         & " " _
         & Hex(Asc(Mid(Response_mod, 5, 1))) _
         & " " _
         & Hex(Asc(Mid(Response_mod, 6, 1))) _
         & " " _
         & Hex(Asc(Mid(Response_mod, 7, 1))) _
         & " " _
         & Hex(Asc(Mid(Response_mod, 8, 1))) _
         & " " _
         & Hex(Asc(Mid(Response_mod, 9, 1))) _
         & " " _
        & Hex(Asc(Mid(Response_mod, 10, 1))) _
         & " " _
        & Hex(Asc(Mid(Response_mod, 11, 1))) _
         & " " _
        & Hex(Asc(Mid(Response_mod, 12, 1))) _
         & " " _
        & Hex(Asc(Mid(Response_mod, 13, 1)))

        Finish = Microsoft.VisualBasic.Timer    ' Set end time.

        Dim J As Integer = 0

        baru.dei_list.Items.Clear()

        CRC_16(Response_mod, (Val(quantitylabel.Text) * 2 + 3))
        baru.dei_crc_hi.Text = Val(Asc(CRC_High))
        baru.dei_crc_lo.Text = Val(Asc(CRC_Low))
        'NOTE!!!
        'code di bawah akan mendeteksi slave yang membaca nilai dibawah 0, jika di bawah 0 maka slave akan disimpan ke variable array address(n), n menunjukkan urutan array.
        ' apabila slave baru pertama kali membaca nilai negatif maka nilai slave akan disimpan ke array address(n) yang baru, tidak akan pernah disimpan bertumpukkan dengan address lain
        ' misal waktu pertama kali run program, slave 100 memiliki nilai negatif, maka 100 akan disimpan di address(1). dan apabila slave 100 salah lagi pada siklus berikutnya maka masih disimpan di address(1)
        ' apabila ada address lain yang bermasalah misal slave 110 maka, n akan ditambah dengan 1. dan 110 akan disimpan di address(2), kemudian apabila siklus berikutnya slave 100 bermasalah lagi maka program akan mendeteksi 100 = address(1) dan hitungan n kembali ke urutan 1 (n=1) 

        'If Response_mod <> "" Then
        '    If Val(Asc(Mid(Response_mod, 1, 1))) = Slave Then
        '        If Val(Asc(Mid(Response_mod, 2, 1))) = 3 Then
        '            If Hex(Asc(Mid(Response_mod, 3, 1))) = Hex(Val(quantitylabel.Text) * 2) Then
        '                If Val(Asc(Mid(Response_mod, 5, 1))) = 20 Then
        '                    If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 > Val(28) Then 'untuk mengetahui nilai suhu yang terbaca minus apa tidak
        '                        'tes1 = 0 'visible tombol true dan memainkan alarm
        '                        If address(n) <> Val(Asc(Mid(Response_mod, 1, 1))) Then 'jika array address sekarang tidak sama dengan slave sekarang maka      
        '                            Do
        '                                For n = 1 To 39 Step 1 'ngecek apakah slave sekarang sudah pernah salah
        '                                    If address(n) = Val(Asc(Mid(Response_mod, 1, 1))) Then ' jika sudah pernah salah, maka loop akan berhenti di n yang mengandung slave sekarang
        '                                        n = n ' nilai n telah dirubah menjadi nilai n yang sesuai dengan slave sekarang
        '                                        'silent(Val(Asc(Mid(Response_mod, 1, 1)))) = 1 ' untuk memberikan efek kedip
        '                                        address(n) = Val(Asc(Mid(Response_mod, 1, 1))) ' array address(n) diisi dengan slave sekarang
        '                                        'tes1 = 0 'visible tombol true
        '                                        save1 = 1 'kedipOn
        '                                        Exit Do 'membuat loop berhenti apabila kondisi "if" terpenuhi
        '                                    End If
        '                                Next
        '                                ' apabila belum pernah salah / belum pernah mengisi array
        '                                For n = 1 To 39 Step 1 'ngecek apakah slave sekarang sudah pernah salah
        '                                    If address(n) = 0 Then ' jika sudah pernah salah, maka loop akan berhenti di n yang mengandung slave sekarang
        '                                        n = n ' nilai n telah dirubah menjadi nilai n yang sesuai dengan slave sekarang
        '                                        'silent(Val(Asc(Mid(Response_mod, 1, 1)))) = 1 ' untuk memberikan efek kedip
        '                                        address(n) = Val(Asc(Mid(Response_mod, 1, 1))) ' array address(n) diisi dengan slave sekarang
        '                                        'tes1 = 0 'visible tombol true
        '                                        save3 = 0 'alarmOn
        '                                        save1 = 1 'KedipOn
        '                                        save5 = 0
        '                                        ListBox1.Items.Add(KonversiRuang(Asc(Mid(Response_mod, 1, 1))) + " : " + tanggal.Text + " : " + "Temperature too high")
        '                                        DbaseAlarmOn(CInt(Slave), CStr(tanggal.Text), CDbl((Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10), CStr("Temperature too high"))
        '                                        Exit Do 'membuat loop berhenti apabila kondisi "if" terpenuhi
        '                                    End If
        '                                Next
        '                            Loop While False
        '                        Else ' apabila address(n) memiliki nilai yang sama dengan slave sekarang
        '                            address(n) = Val(Asc(Mid(Response_mod, 1, 1))) 'mengisi nilai address(n) dengan slave sekarang
        '                            save1 = 1 'kedip 'Val(silent(Val(Asc(Mid(Response_mod, 1, 1)))))
        '                        End If
        '                    Else ' kondisi nilai yang terbaca positif
        '                        Do
        '                            For n = 1 To 39 Step 1 'ngecek apakah slave sekarang sudah pernah mengisi array atau sudah pernah salah
        '                                If address(n) = Val(Asc(Mid(Response_mod, 1, 1))) Then
        '                                    n = n
        '                                    If Hex(Asc(Mid(Response_mod, 8, 2))) < "D7" Then
        '                                        ListBox2.Items.Add(KonversiRuang(Asc(Mid(Response_mod, 1, 1))) + " : " + tanggal.Text + " : " + "Temperature Pass")
        '                                        DbaseAlarmOff(CInt(Slave), CStr(tanggal.Text), CDbl((Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10), CStr("Temperature Pass"))
        '                                        save1 = 0 'kedipOff
        '                                        save3 = 1 'AlarmOff   
        '                                        'address(n) = 0
        '                                    End If
        '                                    Exit Do
        '                                End If
        '                            Next
        '                            'di luar loop for
        '                            n = 0 'mengembalikan hitungan nilai n ke default=0, apabila tidak ada code ini maka nilai n =39
        '                        Loop While False
        '                        'save1 = Val(silent(Val(Asc(Mid(Response_mod, 1, 1)))))
        '                        'save3 = Val(tes1)
        '                    End If


        '                    If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then 'untuk mengetahui nilai tekanan yang terbaca minus apa tidak
        '                        'tes1 = 0 'visible tombol true dan memainkan alarm
        '                        If address(n) <> Val(Asc(Mid(Response_mod, 1, 1))) Then 'jika array address sekarang tidak sama dengan slave sekarang maka                              
        '                            Do
        '                                For n = 1 To 39 Step 1 'ngecek apakah slave sekarang sudah pernah salah
        '                                    If address(n) = Val(Asc(Mid(Response_mod, 1, 1))) Then ' jika sudah pernah salah, maka loop akan berhenti di n yang mengandung slave sekarang
        '                                        n = n ' nilai n telah dirubah menjadi nilai n yang sesuai dengan slave sekarang
        '                                        'silent(Val(Asc(Mid(Response_mod, 1, 1)))) = 1 ' untuk memberikan efek kedip
        '                                        address(n) = Val(Asc(Mid(Response_mod, 1, 1))) ' array address(n) diisi dengan slave sekarang
        '                                        'tes1 = 0 'visible tombol true
        '                                        save2 = 1 'KedipOn
        '                                        Exit Do 'membuat loop berhenti apabila kondisi "if" terpenuhi
        '                                    End If
        '                                Next
        '                                ' apabila belum pernah salah / belum pernah mengisi array
        '                                For n = 1 To 39 Step 1 'ngecek apakah slave sekarang sudah pernah salah
        '                                    If address(n) = 0 Then ' jika sudah pernah salah, maka loop akan berhenti di n yang mengandung slave sekarang
        '                                        n = n ' nilai n telah dirubah menjadi nilai n yang sesuai dengan slave sekarang
        '                                        'silent(Val(Asc(Mid(Response_mod, 1, 1)))) = 1 ' untuk memberikan efek kedip
        '                                        address(n) = Val(Asc(Mid(Response_mod, 1, 1))) ' array address(n) diisi dengan slave sekarang
        '                                        'tes1 = 0 'visible tombol true
        '                                        ListBox1.Items.Add((Asc(Mid(Response_mod, 1, 1))) + " : " + tanggal.Text + " : " + "Pressure too low")
        '                                        DbaseAlarmOn(CInt(Slave), CStr(tanggal.Text), CDbl(Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10), CStr("Pressure too low"))
        '                                        save2 = 1 'KedipOn
        '                                        save4 = 0 'AlarmOn
        '                                        save5 = 0
        '                                        Exit Do 'membuat loop berhenti apabila kondisi "if" terpenuhi
        '                                    End If
        '                                Next
        '                            Loop While False
        '                            'save2 = Val(silent(Val(Asc(Mid(Response_mod, 1, 1)))))
        '                            'save4 = Val(tes1)
        '                        Else ' apabila address(n) memiliki nilai yang sama dengan slave sekarang
        '                            address(n) = Val(Asc(Mid(Response_mod, 1, 1))) 'mengisi nilai address(n) dengan slave sekarang
        '                            'silent(Val(Asc(Mid(Response_mod, 1, 1)))) = 1 ' untuk memberikan efek kedip
        '                            save2 = 1 'KedipOn'Val(silent(Val(Asc(Mid(Response_mod, 1, 1)))))
        '                            'save4 = 0 'Val(tes1)
        '                        End If
        '                    Else ' kondisi nilai yang terbaca positif
        '                        Do
        '                            For n = 1 To 39 Step 1 'ngecek apakah slave sekarang sudah pernah mengisi array atau sudah pernah salah
        '                                If address(n) = Val(Asc(Mid(Response_mod, 1, 1))) Then
        '                                    n = n
        '                                    If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
        '                                        ListBox2.Items.Add(KonversiRuang(Asc(Mid(Response_mod, 1, 1))) + " : " + tanggal.Text + " : " + "Pressure Pass")
        '                                        DbaseAlarmOff(CInt(Slave), CStr(tanggal.Text), CDbl((Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10), CStr("Pressure Pass"))
        '                                        save2 = 0 'KedipOff
        '                                        save4 = 1 'AlarmOff
        '                                        address(n) = 0
        '                                    End If
        '                                    Exit Do
        '                                End If
        '                            Next
        '                            'di luar loop for
        '                            n = 0 'mengembalikan hitungan nilai n ke default=0, apabila tidak ada code ini maka nilai n =39
        '                        Loop While False
        '                    End If
        '                End If
        '            End If
        '        End If
        '    End If
        'End If
        'baru.TextBox2.Text = address(n)
        'silent(Val(Asc(Mid(Response_mod, 1, 1)))) = Val(save1) + Val(save2)
        'tes1 = (Val(save3) * Val(save4)) + Val(save5)
        'If tes1 > 0 Then tes1 = 1

        'If tes1 = 1 Then
        '    ListBox2.Items.Add(KonversiRuang(Asc(Mid(Response_mod, 1, 1))) + " : " + tanggal.Text)
        '    DbaseAlarmOff(CInt(Slave), CStr(tanggal.Text))
        'End If
        'If silent(Val(Asc(Mid(Response_mod, 1, 1)))) = 0 Then
        'code di bawah hampir sama dengan code TK series di VB6
        For i = 4 To Val(quantitylabel.Text) * 2 + 3 Step 2
            baru.dei_list.Items.Add("UI" + Str(k) + ":" + Str((Asc(Mid(Response_mod, i, 1)) * CLng(256) + Asc(Mid(Response_mod, i + 1, 1)))))
            k = k + 1
            If Response_mod <> "" Then
                If Val(Asc(Mid(Response_mod, 2, 1))) = 3 Then
                    If Hex(Asc(Mid(Response_mod, 3, 1))) = Hex(Val(quantitylabel.Text) * 2) Then
                        If Val(Asc(Mid(Response_mod, 5, 1))) = 20 Then
                            'If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then DbaseAlarmOn(CInt(Slave), CStr(tanggal.Text), CDbl(0), CStr("Fail UI1"))
                            'If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then DbaseAlarmOn(CInt(Slave), CStr(tanggal.Text), CDbl(0), CStr("Fail UI2"))
                            Select Case Slave 'menyelesksi respon yang benar dari DEI, dilihat dari address, function modbus, quantity. untuk menyatakan kondisi minus maka digunakan batas D7 dimana itu titik pertemuan antara nilai + dan -
                                Case 100 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "64" Then
                                        Mid(kp, 28, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_100.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg100.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_100.Text = "Fail UI1"
                                                deg100.Visible = False
                                                Mid(kp, 28, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_100.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg100.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_100.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg100.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> 0 Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_100.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa100.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL 
                                                Pa_100.Text = "Fail UI2"
                                                pa100.Visible = False
                                                Mid(kp, 28, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nllainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_100.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa100.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 101 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "65" Then
                                        Mid(kp, 27, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_101.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg101.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_101.Text = "Fail UI1"
                                                deg101.Visible = False
                                                Mid(kp, 27, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nllainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_101.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg101.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_101.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg101.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_101.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa101.Visible = True
                                                If tes1 = 1 Then
                                                    Button5.Visible = False
                                                    My.Computer.Audio.Stop()
                                                ElseIf tes1 = 0 Then
                                                    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                    Button5.Visible = True
                                                End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL 
                                                Pa_101.Text = "Fail UI2"
                                                pa101.Visible = False
                                                Mid(kp, 27, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nllainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_101.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa101.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 102 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "66" Then
                                        Mid(kp, 26, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> CDbl(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_102.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                flow102.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_102.Text = "Fail UI1"
                                                flow102.Visible = False
                                                Mid(kp, 26, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_102.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                flow102.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_102.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                flow102.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> CDbl(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_102.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                persen102.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_102.Text = "Fail UI2"
                                                persen102.Visible = False
                                                Mid(kp, 26, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_102.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                persen102.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 103 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "67" Then
                                        Mid(kp, 25, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_103.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                flow103.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_103.Text = "Fail UI1"
                                                flow103.Visible = False
                                                Mid(kp, 25, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_103.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                flow103.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_103.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                flow103.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_103.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                persen103.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_103.Text = "Fail UI2"
                                                persen103.Visible = False
                                                Mid(kp, 25, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_103.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 1
                                                persen103.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 104 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "68" Then
                                        Mid(kp, 24, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca  
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_104.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg104.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_104.Text = "Fail UI1"
                                                deg104.Visible = False
                                                Mid(kp, 24, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_104.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg104.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_104.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg104.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_104.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa104.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_104.Text = "Fail UI2"
                                                pa104.Visible = False
                                                Mid(kp, 24, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_104.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa104.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 105 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "69" Then
                                        Mid(kp, 23, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_105.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg105.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_105.Text = "Fail UI1"
                                                deg105.Visible = False
                                                Mid(kp, 23, 1) = "0" 'untuk menunjukan bahwa pada alamat ini masih salah nilainya
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_105.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg105.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_105.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg105.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_105.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa105.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_105.Text = "Fail UI2"
                                                pa105.Visible = False
                                                Mid(kp, 23, 1) = "0" 'untuk menunjukan bahwa pada alamat ini masih salah nilainya
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_105.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa105.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 106 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "6A" Then
                                        Mid(kp, 22, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_106.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg106.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_106.Text = "Fail UI1"
                                                deg106.Visible = False
                                                Mid(kp, 22, 1) = "0" 'untuk menunjukan bahwa pada alamat ini masih salah nilainya
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_106.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg106.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_106.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg106.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_106.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa106.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_106.Text = "Fail UI2"
                                                pa106.Visible = False
                                                Mid(kp, 21, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_106.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa106.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 107 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "6B" Then
                                        Mid(kp, 21, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_107.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg107.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_107.Text = "Fail UI1"
                                                deg107.Visible = False
                                                Mid(kp, 21, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_107.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg107.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_107.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg107.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_107.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa107.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                pa107.Visible = False
                                                Mid(kp, 21, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_107.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa107.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 108 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "6C" Then
                                        Mid(kp, 20, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_108.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg108.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_108.Text = "Fail UI1"
                                                deg108.Visible = False
                                                Mid(kp, 20, 1) = "0" 'untuk menunjukan bahwa pada alamat ini masih salah nilainya
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_108.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg108.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_108.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg108.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_108.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa108.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_108.Text = "Fail UI2"
                                                pa108.Visible = False
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_108.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa108.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 109 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "6D" Then
                                        Mid(kp, 19, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_109.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg109.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_109.Text = "Fail UI1"
                                                deg109.Visible = False
                                                Mid(kp, 19, 1) = "0" 'untuk menunjukan bahwa pada alamat ini masih salah nilainya
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_109.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg109.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_109.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg109.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_109.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa109.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_109.Text = "Fail UI2"
                                                pa109.Visible = False
                                                Mid(kp, 19, 1) = "0" 'untuk menunjukan bahwa pada alamat ini masih salah nilainya
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_109.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa109.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 110 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "6E" Then
                                        Mid(kp, 18, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_110.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg110.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_110.Text = "Fail UI1"
                                                deg110.Visible = False
                                                Mid(kp, 18, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_110.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg110.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_110.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg110.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_110.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa110.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_110.Text = "Fail UI2"
                                                pa110.Visible = False
                                                Mid(kp, 18, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_110.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa110.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 111 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "6F" Then
                                        Mid(kp, 17, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_111.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg111.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_111.Text = "Fail UI1"
                                                deg111.Visible = False
                                                Mid(kp, 17, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_111.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg111.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_111.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg111.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_111.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa111.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_111.Text = "Fail UI2"
                                                pa111.Visible = False
                                                Mid(kp, 17, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_111.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa111.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 112 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "70" Then
                                        Mid(kp, 16, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_112.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg112.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_112.Text = "Fail UI1"
                                                deg112.Visible = False
                                                Mid(kp, 16, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_112.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg112.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_112.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg112.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_112.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa112.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_112.Text = "Fail UI2"
                                                pa112.Visible = False
                                                Mid(kp, 16, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_112.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa112.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 113 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "71" Then
                                        Mid(kp, 15, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_113.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg113.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_113.Text = "Fail UI1"
                                                deg113.Visible = False
                                                Mid(kp, 15, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_113.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg113.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_113.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg113.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_113.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa113.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_113.Text = "Fail UI2"
                                                pa113.Visible = False
                                                Mid(kp, 15, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_113.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa113.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 114 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "72" Then
                                        Mid(kp, 14, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_114.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg114.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_114.Text = "Fail UI1"
                                                deg114.Visible = False
                                                Mid(kp, 14, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_114.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg114.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_112.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg112.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_114.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa114.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_114.Text = "Fail UI2"
                                                pa114.Visible = False
                                                Mid(kp, 14, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_114.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa114.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 115 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "73" Then
                                        Mid(kp, 13, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_115.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg115.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_115.Text = "Fail UI1"
                                                deg115.Visible = False
                                                Mid(kp, 13, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_115.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg115.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_115.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg115.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_115.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa115.Visible = True
                                                My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                Button5.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_115.Text = "Fail UI2"
                                                pa115.Visible = False
                                                Mid(kp, 13, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_115.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa115.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 116 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "74" Then
                                        Mid(kp, 12, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_116.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg116.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_116.Text = "Fail UI1"
                                                deg116.Visible = False
                                                Mid(kp, 12, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_116.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg116.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_116.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg116.Visible = True

                                                'If tes1 = 1 Then
                                                '        Button5.Visible = False
                                                '        My.Computer.Audio.Stop()
                                                '    ElseIf tes1 = 0 Then
                                                '        My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '        Button5.Visible = True
                                                '    End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_116.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa116.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_116.Text = "Fail UI2"
                                                pa116.Visible = False
                                                Mid(kp, 12, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_116.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa116.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 117 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "75" Then
                                        Mid(kp, 11, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_117.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg117.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_117.Text = "Fail UI1"
                                                deg117.Visible = False
                                                Mid(kp, 11, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_117.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg117.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_117.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg117.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_117.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa117.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_117.Text = "Fail UI2"
                                                pa117.Visible = False
                                                Mid(kp, 11, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_117.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa117.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 118 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "76" Then
                                        Mid(kp, 10, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_118.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg118.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_118.Text = "Fail UI1"
                                                deg118.Visible = False
                                                Mid(kp, 10, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_118.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg118.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_118.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg118.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_118.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa118.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_118.Text = "Fail UI2"
                                                pa118.Visible = False
                                                Mid(kp, 10, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_118.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa118.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 119 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "77" Then
                                        Mid(kp, 9, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_119.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg119.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_119.Text = "Fail UI1"
                                                deg119.Visible = False
                                                Mid(kp, 9, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_119.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg119.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_118.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg118.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_119.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa119.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_119.Text = "Fail UI2"
                                                deg119.Visible = False
                                                Mid(kp, 9, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_119.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa119.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 120 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "78" Then
                                        Mid(kp, 8, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_120.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg120.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_120.Text = "Fail UI1"
                                                deg120.Visible = False
                                                Mid(kp, 8, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_120.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg120.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_120.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg120.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_120.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa120.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_120.Text = "Fail UI2"
                                                pa120.Visible = False
                                                Mid(kp, 8, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_120.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa120.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 121 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "79" Then
                                        Mid(kp, 7, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_121.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg121.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_121.Text = "Fail UI1"
                                                deg121.Visible = False
                                                Mid(kp, 7, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_121.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg121.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_121.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg121.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_121.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa121.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_121.Text = "Fail UI2"
                                                pa121.Visible = False
                                                Mid(kp, 7, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_121.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa121.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 122 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "7A" Then
                                        Mid(kp, 6, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_122.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg122.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_122.Text = "Fail UI1"
                                                deg122.Visible = False
                                                Mid(kp, 6, 1) = "0" 'untuk menunjukan bahwa pada alamat ini masih salah nilainya
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_122.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg122.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_122.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg122.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_122.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa122.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_122.Text = "Fail UI2"
                                                pa122.Visible = False
                                                Mid(kp, 6, 1) = "0" 'untuk menunjukan bahwa pada alamat ini masih salah nilainya
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_122.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa122.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 123 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "7B" Then
                                        Mid(kp, 5, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_123.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg123.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_123.Text = "Fail UI1"
                                                deg123.Visible = False
                                                Mid(kp, 5, 1) = "0" 'untuk menunjukan bahwa pada alamat ini masih salah nilainya
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_123.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg123.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_123.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg123.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_123.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa123.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_123.Text = "Fail UI2"
                                                pa123.Visible = False
                                                Mid(kp, 5, 1) = "0" 'untuk menunjukan bahwa pada alamat ini masih salah nilainya
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_123.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa123.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 124 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "7C" Then
                                        Mid(kp, 4, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_124.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg124.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_124.Text = "Fail UI1"
                                                deg124.Visible = False
                                                Mid(kp, 4, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_124.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg124.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_123.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg123.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_124.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa124.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_124.Text = "Fail UI2"
                                                pa124.Visible = False
                                                Mid(kp, 4, 1) = "0" 'untuk menunjukan bahwa pada alamat ini nilainya masih salah
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_124.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa124.Visible = True
                                            End If
                                        End If
                                    End If
                                Case 125 : If Hex(Asc(Mid(Response_mod, 1, 1))) = "7D" Then
                                        Mid(kp, 3, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                                        'membaca suhu UI1
                                        If (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 6, 2))) > "D7" Then
                                                slave_125.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 6, 1))) * 256) + Val(Asc(Mid(Response_mod, 7, 1)))))) / 10
                                                deg125.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                slave_125.Text = "Fail UI1"
                                                deg125.Visible = False
                                                Mid(kp, 3, 1) = "0" 'untuk menunjukan bahwa pada alamat ini masih salah nilainya
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 <= Val(28) Then
                                                slave_125.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg125.Visible = True
                                            ElseIf (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10 < Val(100) Then
                                                slave_125.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
                                                deg125.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            End If
                                        End If
                                        'membaca tekanan UI2
                                        If (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 <> Val(0) Then
                                            If Hex(Asc(Mid(Response_mod, 8, 2))) > "D7" Then
                                                Pa_125.Text = Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10
                                                pa125.Visible = True
                                                'If tes1 = 1 Then
                                                '    Button5.Visible = False
                                                '    My.Computer.Audio.Stop()
                                                'ElseIf tes1 = 0 Then
                                                '    My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                                                '    Button5.Visible = True
                                                'End If
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 = Val(3276.8) Then 'apabila di display DEI menampilkan FAIL
                                                Pa_125.Text = "Fail UI2"
                                                pa125.Visible = False
                                                Mid(kp, 3, 1) = "0" 'untuk menunjukan bahwa pada alamat ini masih salah nilainya
                                            ElseIf (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10 < Val(100) Then
                                                Pa_125.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
                                                pa125.Visible = True
                                            End If
                                        End If
                                    End If
                            End Select
                        End If
                    End If
                End If
            End If
            J = J + 1
        Next i
        save7 = jumlahsalah

        alamatsalah() ' untuk menghitung jumlah slave address DEI dan TZ yang bermasalah
        baru.TextBox4.Text = jumlahsalah 'jumlahsalah merupakan variable output dari private sub alamatsalah()
        If jumlahsalah > save7 Then
            save6 = 1
            ListBox1.Items.Add(KonversiRuang((Asc(Mid(Response_mod, 1, 1)))) + " : " + tanggal.Text)
            'DbaseAlarmOn(CInt(Slave), CStr(tanggal.Text), CDbl(Val("-" & (65536 - ((Val(Asc(Mid(Response_mod, 8, 1))) * 256) + Val(Asc(Mid(Response_mod, 9, 1)))))) / 10), CStr("Pressure too low"))
        End If
        tes1 = save5 * save6
        'If tes1 > 0 Then tes1 = 1
        If tes1 > 0 Then
            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
            Button5.Visible = True
        ElseIf tes1 = 0 Then
            'My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
            'Button5.Visible = True
            Button5.Visible = False
        End If
        'mendeteksi offline 
        'apabila response dari slave address salah selama 10 menit maka akan ditulis offline pada label.text 
        '10 menit setelah dihitung-hitung dengan asumsi DEI_timer 1 detik dan DigitalOutput_timer 1 detik, jadi dibutuhkan 12 kali loop dalam 10 menit untuk mengakses slave address yang sama
        If Response_mod <> "" Then
            If Val(Asc(Mid(Response_mod, 2, 1))) = 3 Then
                checktrue(Slave) 'apabila telah memberikan response yang benar, kondisi ini akan terpenuhi apabila memiliki salah satu dari 2 kemungkinan, 1. memang memiliki response yang selalu benar, 2. response benar tapi pernah tidak menjawab dan tidak sampai 10 menit 
            Else checkfalse(Slave) ' apabila response tidak function 3
            End If
        Else checkfalse(Slave) 'apabila tidak ada response maka counter akan menghitung
        End If

        baru.slaveAdd.Text = Slave
        Slave = Slave + 1
        baru.TextBox3.Text = n
        Waktu.Enabled = True 'mengaktifkan timer waktu, apabila tidak ada code ini maka akan error pada timer waktu, karena pada timer waktu juga membaca response dan juga timer waktu 100 ms, sedangkan dei_timer 1000 ms. 
        DEI_timer.Enabled = False
        TimerPV.Enabled = True
        DigitalOutput.Enabled = True
    End Sub
    Private Sub checkfalse(offlineadd As Integer)
        ' code dibawah ini ntuk menghitung berapa kali slave address tidak meresponse dengan benar , setiap slave address DEI maupun TZ. setiap slave address memiliki counter sendiri, jadi memungkinkan untuk digunakan counter multitasking
        Select Case offlineadd
            Case 1 : If count(26) = 13 Then
                    functiontulisoffline(1)
                    count(26) = 0
                Else count(26) = count(26) + 1
                End If
            Case 2 : If count(27) = 13 Then
                    functiontulisoffline(2)
                    count(27) = 0
                Else
                    count(27) = count(27) + 1
                End If
            Case 3 : If count(28) = 13 Then
                    functiontulisoffline(3)
                    count(28) = 0
                Else count(28) = count(28) + 1
                End If
            Case 4 : If count(29) = 13 Then
                    functiontulisoffline(4)
                    count(29) = 0
                Else count(29) = count(29) + 1
                End If
            Case 5 : If count(30) = 13 Then
                    functiontulisoffline(5)
                    count(30) = 0
                Else
                    count(30) = count(30) + 1
                End If
            Case 6 : If count(31) = 13 Then
                    functiontulisoffline(6)
                    count(31) = 0
                Else count(31) = count(31) + 1
                End If
            Case 7 : If count(32) = 13 Then
                    functiontulisoffline(7)
                    count(32) = 0
                Else count(32) = count(32) + 1
                End If
            Case 8 : If count(33) = 13 Then
                    functiontulisoffline(8)
                    count(33) = 0
                Else count(33) = count(33) + 1
                End If
            Case 9 : If count(34) = 13 Then
                    functiontulisoffline(9)
                    count(34) = 0
                Else count(34) = count(34) + 1
                End If
            Case 10 : If count(35) = 13 Then
                    functiontulisoffline(10)
                    count(35) = 0
                Else count(35) = count(35) + 1
                End If
            Case 11 : If count(36) = 13 Then
                    functiontulisoffline(11)
                    count(36) = 0
                Else count(36) = count(36) + 1
                End If
            Case 12 : If count(37) = 13 Then
                    functiontulisoffline(12)
                    count(37) = 0
                Else count(37) = count(37) + 1
                End If
            Case 15 : If count(38) = 13 Then
                    functiontulisoffline(15)
                    count(38) = 0
                Else count(38) = count(38) + 1
                End If
            Case 100 : If count(0) = 13 Then
                    functiontulisoffline(100)
                    count(0) = 0
                Else count(0) = count(0) + 1
                End If
            Case 101 : If count(1) = 13 Then
                    functiontulisoffline(101)
                    count(1) = 0
                Else count(1) = count(1) + 1
                End If
            Case 102 : If count(2) = 13 Then
                    functiontulisoffline(102)
                    count(2) = 0
                Else count(2) = count(2) + 1
                End If
            Case 103 : If count(3) = 13 Then
                    functiontulisoffline(103)
                    count(3) = 0
                Else count(3) = count(3) + 1
                End If
            Case 104 : If count(4) = 13 Then
                    functiontulisoffline(104)
                    count(4) = 0
                Else count(4) = count(4) + 1
                End If
            Case 105 : If count(5) = 13 Then
                    functiontulisoffline(105)
                    count(5) = 0
                Else count(5) = count(5) + 1
                End If
            Case 106 : If count(6) = 13 Then
                    functiontulisoffline(106)
                    count(6) = 0
                Else count(6) = count(6) + 1
                End If
            Case 107 : If count(7) = 13 Then
                    functiontulisoffline(107)
                    count(7) = 0
                Else count(7) = count(7) + 1
                End If
            Case 108 : If count(8) = 13 Then
                    functiontulisoffline(108)
                    count(8) = 0
                Else count(8) = count(8) + 1
                End If
            Case 109 : If count(9) = 13 Then
                    functiontulisoffline(109)
                    count(9) = 0
                Else count(9) = count(9) + 1
                End If
            Case 110 : If count(10) = 13 Then
                    functiontulisoffline(110)
                    count(10) = 0
                Else count(10) = count(10) + 1
                End If
            Case 111 : If count(11) = 13 Then
                    functiontulisoffline(111)
                    count(11) = 0
                Else count(11) = count(11) + 1
                End If
            Case 112 : If count(12) = 13 Then
                    functiontulisoffline(112)
                    count(12) = 0
                Else count(12) = count(12) + 1
                End If
            Case 113 : If count(13) = 13 Then
                    functiontulisoffline(113)
                    count(13) = 0
                Else count(13) = count(13) + 1
                End If
            Case 114 : If count(14) = 13 Then
                    functiontulisoffline(114)
                    count(14) = 0
                Else count(14) = count(14) + 1
                End If
            Case 115 : If count(15) = 13 Then
                    functiontulisoffline(115)
                    count(15) = 0
                Else count(15) = count(15) + 1
                End If
            Case 116 : If count(16) = 13 Then
                    functiontulisoffline(116)
                    count(16) = 0
                Else count(16) = count(16) + 1
                End If
            Case 117 : If count(17) = 13 Then
                    functiontulisoffline(117)
                    count(17) = 0
                Else count(17) = count(17) + 1
                End If
            Case 118 : If count(18) = 13 Then
                    functiontulisoffline(118)
                    count(18) = 0
                Else count(18) = count(18) + 1
                End If
            Case 119 : If count(19) = 13 Then
                    functiontulisoffline(119)
                    count(19) = 0
                Else count(119) = count(19) + 1
                End If
            Case 120 : If count(20) = 13 Then
                    functiontulisoffline(120)
                    count(20) = 0
                Else count(20) = count(20) + 1
                End If
            Case 121 : If count(21) = 13 Then
                    functiontulisoffline(121)
                    count(21) = 0
                Else count(21) = count(21) + 1
                End If
            Case 122 : If count(22) = 13 Then
                    functiontulisoffline(122)
                    count(22) = 0
                Else count(22) = count(22) + 1
                End If
            Case 123 : If count(23) = 13 Then
                    functiontulisoffline(123)
                    count(23) = 0
                Else count(23) = count(23) + 1
                End If
            Case 124 : If count(24) = 13 Then
                    functiontulisoffline(124)
                    count(24) = 0
                Else count(24) = count(24) + 1
                End If
            Case 125 : If count(25) = 13 Then
                    functiontulisoffline(125)
                    count(25) = 0
                Else count(25) = count(25) + 1
                End If
        End Select
    End Sub
    Private Sub checktrue(onlineadd As Integer)
        'code di bawah untuk mengmebalikan counter ke nilai 0, karena telah meresponse dengan benar.
        Select Case onlineadd
            Case 1 : count(26) = 0
            Case 2 : count(27) = 0
            Case 3 : count(28) = 0
            Case 4 : count(29) = 0
            Case 5 : count(30) = 0
            Case 6 : count(31) = 0
            Case 7 : count(32) = 0
            Case 8 : count(33) = 0
            Case 9 : count(34) = 0
            Case 10 : count(35) = 0
            Case 11 : count(36) = 0
            Case 12 : count(37) = 0
            Case 15 : count(38) = 0
            Case 100 : count(0) = 0
            Case 101 : count(1) = 0
            Case 102 : count(2) = 0
            Case 103 : count(3) = 0
            Case 104 : count(4) = 0
            Case 105 : count(5) = 0
            Case 106 : count(6) = 0
            Case 107 : count(7) = 0
            Case 108 : count(8) = 0
            Case 109 : count(9) = 0
            Case 110 : count(10) = 0
            Case 111 : count(11) = 0
            Case 112 : count(12) = 0
            Case 113 : count(13) = 0
            Case 114 : count(14) = 0
            Case 115 : count(15) = 0
            Case 116 : count(16) = 0
            Case 117 : count(17) = 0
            Case 118 : count(18) = 0
            Case 119 : count(19) = 0
            Case 120 : count(20) = 0
            Case 121 : count(21) = 0
            Case 122 : count(22) = 0
            Case 123 : count(23) = 0
            Case 124 : count(24) = 0
            Case 125 : count(25) = 0
        End Select
    End Sub
    Private Sub functiontulisoffline(ByVal alamat As Integer)
        'untuk menulis offline pada label.text sesuai slave address yang diinputkan
        Select Case alamat
            Case 1 : BIOCOLD032.Text = "OFFLINE"
                SV032.Text = "OFFLINE"
            Case 2 : BIOCOLD037A.Text = "OFFLINE"
                SV037A.Text = "OFFLINE"
            Case 3 : BIOCOLD038A.Text = "OFFLINE"
                SV038A.Text = "OFFLINE"
            Case 4 : BIOCOLD038B.Text = "OFFLINE"
                SV038B.Text = "OFFLINE"
            Case 5 : BIOCOLD039.Text = "OFFLINE"
                SV039.Text = "OFFLINE"
            Case 6 : BIOCOLD043.Text = "OFFLINE"
                SV043.Text = "OFFLINE"
            Case 7 : BIOCOLD040.Text = "OFFLINE"
                SV040.Text = "OFFLINE"
            Case 8 : BIOCOLD041A.Text = "OFFLINE"
                SV041A.Text = "OFFLINE"
            Case 9 : BIOCOLD041A.Text = "OFFLINE"
                SV041A.Text = "OFFLINE"
            Case 10 : BIOCOLD042A.Text = "OFFLINE"
                SV042A.Text = "OFFLINE"
            Case 11 : BIOCOLD042B.Text = "OFFLINE"
                SV042B.Text = "OFFLINE"
            Case 12 : BIOCOLD044.Text = "OFFLINE"
                SV044.Text = "OFFLINE"
            Case 100 : slave_100.Text = "Offline"
                Pa_100.Text = "Offline"
            Case 101 : slave_101.Text = "Offline"
                Pa_101.Text = "Offline"
            Case 102 : slave_102.Text = "Offline"
                Pa_102.Text = "Offline"
            Case 103 : slave_103.Text = "Offline"
                Pa_103.Text = "Offline"
            Case 104 : slave_104.Text = "Offline"
                Pa_104.Text = "Offline"
            Case 105 : slave_105.Text = "Offline"
                Pa_105.Text = "Offline"
            Case 106 : slave_106.Text = "Offline"
                Pa_106.Text = "Offline"
            Case 107 : slave_107.Text = "Offline"
                Pa_107.Text = "Offline"
            Case 108 : slave_108.Text = "Offline"
                Pa_108.Text = "Offline"
            Case 109 : slave_109.Text = "Offline"
                Pa_109.Text = "Offline"
            Case 110 : slave_110.Text = "Offline"
                Pa_110.Text = "Offline"
            Case 111 : slave_111.Text = "Offline"
                Pa_111.Text = "Offline"
            Case 112 : slave_112.Text = "Offline"
                Pa_112.Text = "Offline"
            Case 113 : slave_113.Text = "Offline"
                Pa_113.Text = "Offline"
            Case 114 : slave_114.Text = "Offline"
                Pa_114.Text = "Offline"
            Case 115 : slave_115.Text = "Offline"
                Pa_115.Text = "Offline"
            Case 116 : slave_116.Text = "Offline"
                Pa_116.Text = "Offline"
            Case 117 : slave_117.Text = "Offline"
                Pa_117.Text = "Offline"
            Case 118 : slave_118.Text = "Offline"
                Pa_118.Text = "Offline"
            Case 119 : slave_119.Text = "Offline"
                Pa_119.Text = "Offline"
            Case 120 : slave_120.Text = "Offline"
                Pa_120.Text = "Offline"
            Case 121 : slave_121.Text = "Offline"
                Pa_121.Text = "Offline"
            Case 122 : slave_122.Text = "Offline"
                Pa_122.Text = "Offline"
            Case 123 : slave_123.Text = "Offline"
                Pa_123.Text = "Offline"
            Case 124 : slave_124.Text = "Offline"
                Pa_124.Text = "Offline"
            Case 125 : slave_125.Text = "Offline"
                Pa_125.Text = "Offline"
        End Select
    End Sub
    Private Sub dei_dbase_Tick(sender As Object, e As EventArgs) Handles dei_dbase.Tick
        'TZDbase.Enabled = False
        Dim provider, dataFile As String
        Dim myConnection As OleDbConnection = New OleDbConnection
        provider = "Provider=Microsoft.ACE.OleDB.12.0;Data Source="
        dataFile = "E:\PENS\try.mdb"
        myConnection.ConnectionString = provider & dataFile
        myConnection.Open()
        Dim str As String
        'Digunakan untuk memasukan data yang terbaca ke database sesuai dengan field yang ada
        str = "insert Into Table1([d/t],[Suhu POS 3],[Suhu PPIC],[2],[3],[Suhu Emulsi 2],[Suhu Emulsi 1],[Suhu POS 4],[Suhu INO],[Suhu Panen 4],[Suhu Panen 3],[Suhu Koridor],[Suhu INO 3],[Suhu Pra INO 2],[Suhu Passage],[Suhu Pra INO 1],[Suhu Alat Steril],[Suhu STERILISASI],[Suhu Persiapan],[Suhu INO 2],[Suhu INO 1],[Suhu POS INO 2],[Suhu POS INO 1],[Suhu Panen 1],[Suhu Panen 2],[Suhu Inaktif 1],[Suhu Inaktif2],[Tekanan POS 3],[Tekanan PPIC],[press102],[press103],[Tekanan Emulsi 2],[Tekanan Emulsi 1],[Tekanan POS 4],[Tekanan INO],[Tekanan Panen 4],[Tekanan Panen 3],[Tekanan Koridor],[Tekanan INO 3],[Tekanan Pra INO 2],[Tekanan Passage],[Tekanan Pra INO 1],[Tekanan Alat Steril],[Tekanan STERILISASI],[Tekanan Persiapan],[Tekanan INO 2],[Tekanan INO 1],[Tekanan POS INO 2],[Tekanan POS INO 1],[Tekanan Panen 1],[Tekanan Panen 2],[Tekanan Inaktif 1],[Tekanan Inaktif 2]) Values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
        Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
        cmd.Parameters.Add(New OleDbParameter("d/t", CDate(tanggal.Text)))
        'record data ke database DEI, apabila ada alamat yang belum ada nilainya maka, akan disimpan sebagai nilai 0 di database. hal ini telah diantisipasi sebelumnya dengan adanya variable kp yang memastikan semua alamat yang diinginkan memiliki nilai terlebih dahulu meskipun niilainya 0
        If slave_100.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu POS 3", CDbl(0.0001)))
        ElseIf slave_100.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu POS 3", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu POS 3", CDbl(slave_100.Text)))
        End If

        If slave_101.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu PPIC", CDbl(0.0001)))
        ElseIf slave_101.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu PPIC", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu PPIC", CDbl(slave_101.Text)))
        End If

        If slave_102.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("2", CDbl(0.0001)))
        ElseIf slave_102.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("2", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("2", CDbl(slave_102.Text)))
        End If

        If slave_103.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("3", CDbl(0.0001)))
        ElseIf slave_103.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("3", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("3", CDbl(slave_103.Text)))
        End If

        If slave_104.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Emulsi 2", CDbl(0.0001)))
        ElseIf slave_104.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Emulsi 2", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu Emulsi 2", CDbl(slave_104.Text)))
        End If

        If slave_105.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Emulsi 1", CDbl(0.0001)))
        ElseIf slave_105.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Emulsi 1", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu Emulsi 1", CDbl(slave_105.Text)))
        End If

        If slave_106.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu POS 4", CDbl(0.0001)))
        ElseIf slave_106.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu POS 4", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu POS 4", CDbl(slave_106.Text)))
        End If

        If slave_107.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu INO", CDbl(0.0001)))
        ElseIf slave_107.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu INO", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu INO", CDbl(slave_107.Text)))
        End If

        If slave_108.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Panen 4", CDbl(0.0001)))
        ElseIf slave_108.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Panen 4", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu Panen 4", CDbl(slave_108.Text)))
        End If

        If slave_109.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Panen 3", CDbl(0.0001)))
        ElseIf slave_109.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Panen 3", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu Panen 3", CDbl(slave_109.Text)))
        End If

        If slave_110.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Koridor", CDbl(0.0001)))
        ElseIf slave_110.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Koridor", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu Koridor", CDbl(slave_110.Text)))
        End If

        If slave_111.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu INO 3", CDbl(0.0001)))
        ElseIf slave_111.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu INO 3", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu INO 3", CDbl(slave_111.Text)))
        End If

        If slave_112.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu PRA INO 2", CDbl(0.0001)))
        ElseIf slave_112.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu PRA INO 2", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu PRA INO 2", CDbl(slave_112.Text)))
        End If

        If slave_113.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Passage", CDbl(0.0001)))
        ElseIf slave_113.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Passage", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu Passage", CDbl(slave_113.Text)))
        End If

        If slave_114.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu PRA INO 1", CDbl(0.0001)))
        ElseIf slave_114.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu PRA INO 1", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu PRA INO 1", CDbl(slave_114.Text)))
        End If

        If slave_115.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Alat Steril", CDbl(0.0001)))
        ElseIf slave_115.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Alat Steril", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu Alat Steril", CDbl(slave_115.Text)))
        End If

        If slave_116.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu STERILISASI", CDbl(0.0001)))
        ElseIf slave_116.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu STERILISASI", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu STERILISASI", CDbl(slave_116.Text)))
        End If

        If slave_117.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Persiapan", CDbl(0.0001)))
        ElseIf slave_117.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Persiapan", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu Persiapan", CDbl(slave_117.Text)))
        End If

        If slave_118.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu INO 2", CDbl(0.0001)))
        ElseIf slave_118.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu INO 2", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu INO 2", CDbl(slave_118.Text)))
        End If

        If slave_119.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu INO 1", CDbl(0.0001)))
        ElseIf slave_119.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu INO 1", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu INO 1", CDbl(slave_119.Text)))
        End If

        If slave_120.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu POS INO 2", CDbl(0.0001)))
        ElseIf slave_120.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu POS INO 2", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu POS INO 2", CDbl(slave_120.Text)))
        End If

        If slave_121.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu POS INO 1", CDbl(0.0001)))
        ElseIf slave_121.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu POS INO 1", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu POS INO 1", CDbl(slave_121.Text)))
        End If

        If slave_122.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Panen 1", CDbl(0.0001)))
        ElseIf slave_122.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Panen 1", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu Panen 1", CDbl(slave_122.Text)))
        End If

        If slave_123.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Panen 2", CDbl(0.0001)))
        ElseIf slave_123.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Panen 2", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu Panen 2", CDbl(slave_123.Text)))
        End If

        If slave_124.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Inaktif 1", CDbl(0.0002)))
        ElseIf slave_124.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Inaktif 1", CDbl(0.0002)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu Inaktif 1", CDbl(slave_124.Text)))
        End If

        If slave_125.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Inaktif2", 0))
        ElseIf slave_125.Text = "Fail UI1" Then
            cmd.Parameters.Add(New OleDbParameter("Suhu Inaktif2", 0))
        Else
            cmd.Parameters.Add(New OleDbParameter("Suhu Inaktif2", CDbl(slave_125.Text)))
        End If

        If Pa_100.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan POS 3", CDbl(0.0001)))
        ElseIf Pa_100.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan POS 3", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan POS 3", CDbl(Pa_100.Text)))
        End If

        If Pa_101.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan PPIC", CDbl(0.0001)))
        ElseIf Pa_101.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan PPIC", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan PPIC", CDbl(Pa_101.Text)))
        End If

        If Pa_102.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("press102", CDbl(0.0001)))
        ElseIf Pa_102.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("press102", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("press102", CDbl(Pa_102.Text)))
        End If

        If Pa_103.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("press103", CDbl(0.0001)))
        ElseIf Pa_103.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("press103", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("press103", CDbl(Pa_103.Text)))
        End If

        If Pa_104.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Emulsi 2", CDbl(0.0001)))
        ElseIf Pa_104.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Emulsi 2", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan Emulsi 2", CDbl(Pa_104.Text)))
        End If

        If Pa_105.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Emulsi 1", CDbl(0.0001)))
        ElseIf Pa_105.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Emulsi 1", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan Emulsi 1", CDbl(Pa_105.Text)))
        End If

        If Pa_106.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan POS 4", CDbl(0.0001)))
        ElseIf Pa_106.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan POS 4", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan POS 4", CDbl(Pa_106.Text)))
        End If

        If Pa_107.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan INO", CDbl(0.0001)))
        ElseIf Pa_107.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan INO", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan INO", CDbl(Pa_107.Text)))
        End If

        If Pa_108.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Panen 4", CDbl(0.0001)))
        ElseIf Pa_108.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Panen 4", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan Panen 4", CDbl(Pa_108.Text)))
        End If

        If Pa_109.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Panen 3", CDbl(0.0001)))
        ElseIf Pa_109.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Panen 3", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan Panen 3", CDbl(Pa_109.Text)))
        End If

        If Pa_110.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Koridor", CDbl(0.0001)))
        ElseIf Pa_110.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Koridor", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan Koridor", CDbl(Pa_110.Text)))
        End If

        If Pa_111.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan INO 3", CDbl(0.0001)))
        ElseIf Pa_111.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan INO 3", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan INO 3", CDbl(Pa_111.Text)))
        End If

        If Pa_112.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Pra INO 2", CDbl(0.0001)))
        ElseIf Pa_112.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Pra INO 2", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan Pra INO 2", CDbl(Pa_112.Text)))
        End If

        If Pa_113.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Passage", CDbl(0.0001)))
        ElseIf Pa_113.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Passage", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan Passage", CDbl(Pa_113.Text)))
        End If

        If Pa_114.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Pra INO 1", CDbl(0.0001)))
        ElseIf Pa_114.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Pra INO 1", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan Pra INO 1", CDbl(Pa_114.Text)))
        End If

        If Pa_115.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Alat Steril", CDbl(0.0001)))
        ElseIf Pa_115.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Alat Steril", CDbl(0.0003)))
        Else cmd.Parameters.Add(New OleDbParameter("Tekanan Alat Steril", CDbl(Pa_115.Text)))
        End If

        If Pa_116.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan STERILISASI", CDbl(0.0001)))
        ElseIf Pa_116.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan STERILISASI", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan STERILISASI", CDbl(Pa_116.Text)))
        End If

        If Pa_117.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Persiapan", CDbl(0.0001)))
        ElseIf Pa_117.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Persiapan", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan Persiapan", CDbl(Pa_117.Text)))
        End If

        If Pa_118.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan INO 2", CDbl(0.0001)))
        ElseIf Pa_118.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan INO 2", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan INO 2", CDbl(Pa_118.Text)))
        End If

        If Pa_119.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan INO 1", CDbl(0.0001)))
        ElseIf Pa_119.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan INO 1", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan INO 1", CDbl(Pa_119.Text)))
        End If

        If Pa_120.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan POS INO 2", CDbl(0.0001)))
        ElseIf Pa_120.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan POS INO 2", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan POS INO 2", CDbl(Pa_120.Text)))
        End If

        If Pa_121.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan POS INO 1", CDbl(0.0001)))
        ElseIf Pa_121.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan POS INO 1", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan POS INO 1", CDbl(Pa_121.Text)))
        End If

        If Pa_122.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Panen 1", CDbl(0.0001)))
        ElseIf Pa_122.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Panen 1", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan Panen 1", CDbl(Pa_122.Text)))
        End If

        If Pa_123.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Panen 2", CDbl(0.0001)))
        ElseIf Pa_123.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Panen 2", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan Panen 2", CDbl(Pa_123.Text)))
        End If

        If Pa_124.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Inaktif 1", CDbl(0.0001)))
        ElseIf Pa_124.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Inaktif 1", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan Inaktif 1", CDbl(Pa_124.Text)))
        End If

        If Pa_125.Text = "Offline" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Inaktif 2", CDbl(0.0001)))
        ElseIf Pa_125.Text = "Fail UI2" Then
            cmd.Parameters.Add(New OleDbParameter("Tekanan Inaktif 2", CDbl(0.0003)))
        Else
            cmd.Parameters.Add(New OleDbParameter("Tekanan Inaktif 2", CDbl(Pa_125.Text)))
        End If

        Label28.Text = DateAndTime.Now 'menunjukkan last update data

        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        'TZDbase.Enabled = True
        'dei_dbase.Enabled = False
    End Sub
    Private Sub DigitalOutput_Tick(sender As Object, e As EventArgs) Handles DigitalOutput.Tick
        Dim data As String
        On Error Resume Next
        Dim k As Integer
        'If nodenum.Text = 111 Then nodenum.Text = 110
        If slave1 = 126 Then slave1 = 100
        'slave1 = 105
        baru.slaveAdd1.Text = slave1
        data = Chr(slave1) + Chr(3) + Chr(Val(startaddlabel1.Text) \ 256) + Chr(Val(startaddlabel1.Text) Mod 256) + Chr(0) + Chr(Val(quantitylabel1.Text))
        CRC_16(data, 6)
        data = data + Chr(CRC_High) + Chr(CRC_Low)
        Dim PauseTime, Start, Finish
        baru.dei_response1.Text = ""
        MSComm1.InputLen = 0

        baru.dei_request1.Text = Val(Asc(Mid(data, 1, 3))) & Val(Asc(Mid(data, 2, 1))) & Val(Asc(Mid(data, 3, 1))) _
        & (Val(Asc(Mid(data, 4, 1))) + (Val(Asc(Mid(data, 5, 1))))) _
        & (Val(Asc(Mid(data, 6, 1))) + (Val(Asc(Mid(data, 7, 1))))) _
        & (Mid(data, 8, 1)) & (Mid(data, 9, 1))


        MSComm1.Output = data

        Do While MSComm1.OutBufferCount > 0
        Loop
        PauseTime = 0.1  ' Set duration.
        Start = Microsoft.VisualBasic.Timer    ' Set start time.
        Do While (Microsoft.VisualBasic.Timer < Start + PauseTime) And (MSComm1.InBufferCount < Val(quantitylabel1.Text) * 2 + 5)
            Application.DoEvents()
        Loop

        baru.dei_inBuffer1.Text = (MSComm1.InBufferCount)

        Response_mod = MSComm1.Input

        baru.dei_response1.Text = Val(Asc(Mid(Response_mod, 1, 1))) _
        & Val(Asc(Mid(Response_mod, 2, 1))) _
        & Val(Asc(Mid(Response_mod, 3, 1))) _
        & Val(Asc(Mid(Response_mod, 4, 1))) * CLng(256) + (Val(Asc(Mid(Response_mod, 5, 1)))) _
        & Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + (Val(Asc(Mid(Response_mod, 7, 1)))) _
        & Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + (Val(Asc(Mid(Response_mod, 9, 1)))) _
        & Val(Asc(Mid(Response_mod, 10, 1))) * CLng(256) + (Val(Asc(Mid(Response_mod, 11, 1)))) _
        & (Mid(Response_mod, 11, 1)) & (Mid(Response_mod, 12, 1))


        Finish = Microsoft.VisualBasic.Timer    ' Set end time.

        Dim J As Integer = 0

        baru.dei_list1.Items.Clear()
        CRC_16(Response_mod, (Val(quantitylabel1.Text) * 2 + 3))
        baru.dei_crc_hi1.Text = Val(Asc(CRC_High))
        baru.dei_crc_lo1.Text = Val(Asc(CRC_Low))

        'karena menggunakan array untuk menyimpan hasil pembacaan maka diperlukan men"0"kan array setiap siklus baru 
        If Response_mod <> "" Then
            If Val(Asc(Mid(Response_mod, 2, 1))) = 3 Then
                k = 0
                For i = 4 To Val(quantitylabel1.Text) * 2 + 3 Step 2
                    k = k + 1
                    ID(k) = 0
                Next i
                k = 0
                For i = 4 To Val(quantitylabel1.Text) * 2 + 3 Step 2
                    k = k + 1
                    ID(k) = Asc(Mid(Response_mod, i, 1)) * CLng(256) + Asc(Mid(Response_mod, i + 1, 1))
                    If ID(k) = 1 Then status = "ON" Else status = "OFF"
                    baru.dei_list1.Items.Add("ID" + Str(slave1) + "DO" + Str(k) + " : " + status)
                Next i
                Select Case Val(Asc(Mid(Response_mod, 1, 1)))
                    Case 100 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label107.BackColor = Color.LightGreen
                        Else label107.BackColor = Color.Red
                        End If
                    Case 101 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label113.BackColor = Color.LightGreen
                        Else Label113.BackColor = Color.Red
                        End If
                    Case 104 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label112.BackColor = Color.LightGreen
                        Else Label112.BackColor = Color.Red
                        End If
                    Case 105 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label111.BackColor = Color.LightGreen
                        Else Label111.BackColor = Color.Red
                        End If
                    Case 106 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label110.BackColor = Color.LightGreen
                        Else Label110.BackColor = Color.Red
                        End If
                    Case 107 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label109.BackColor = Color.LightGreen
                        Else Label109.BackColor = Color.Red
                        End If
                    Case 108 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label108.BackColor = Color.LightGreen
                        Else Label108.BackColor = Color.Red
                        End If
                    Case 109 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label105.BackColor = Color.LightGreen
                        Else Label105.BackColor = Color.Red
                        End If
                    Case 110 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label98.BackColor = Color.LightGreen
                        Else Label98.BackColor = Color.Red
                        End If
                    Case 111 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label106.BackColor = Color.LightGreen
                        Else Label106.BackColor = Color.Red
                        End If
                    Case 112 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label104.BackColor = Color.LightGreen
                        Else Label104.BackColor = Color.Red
                        End If
                    Case 113 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label101.BackColor = Color.LightGreen
                        Else Label101.BackColor = Color.Red
                        End If
                    Case 114 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label103.BackColor = Color.LightGreen
                        Else Label103.BackColor = Color.Red
                        End If
                    Case 115 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label102.BackColor = Color.LightGreen
                        Else Label102.BackColor = Color.Red
                        End If
                    Case 116 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label100.BackColor = Color.LightGreen
                        Else Label100.BackColor = Color.Red
                        End If
                    Case 117 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label99.BackColor = Color.LightGreen
                        Else Label99.BackColor = Color.Red
                        End If
                    Case 118 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label94.BackColor = Color.LightGreen
                        Else Label94.BackColor = Color.Red
                        End If
                    Case 119 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label90.BackColor = Color.LightGreen
                        Else Label90.BackColor = Color.Red
                        End If
                    Case 120 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label95.BackColor = Color.LightGreen
                        Else Label95.BackColor = Color.Red
                        End If
                    Case 121 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label91.BackColor = Color.LightGreen
                        Else Label91.BackColor = Color.Red
                        End If
                    Case 122 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label92.BackColor = Color.LightGreen
                        Else Label92.BackColor = Color.Red
                        End If
                    Case 123 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label96.BackColor = Color.LightGreen
                        Else Label96.BackColor = Color.Red
                        End If
                    Case 124 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label93.BackColor = Color.LightGreen
                        Else Label93.BackColor = Color.Red
                        End If
                    Case 125 : If Val(ID(1)) * Val(ID(2)) = 1 Then
                            Label97.BackColor = Color.LightGreen
                        Else Label97.BackColor = Color.Red
                        End If
                End Select
            Else baru.dei_list1.Items.Clear()
            End If
        End If



        slave1 = slave1 + 1
        DigitalOutput.Enabled = False
        DEI_timer.Enabled = True
    End Sub
    Private Sub Waktu_Tick(sender As Object, e As EventArgs) Handles Waktu.Tick
        Dim var(30), var1(30) As Integer
        TextBox3.Text = DateAndTime.Now
        tanggal.Text = DateAndTime.Now
        now.Text = DateAndTime.Now
        inc = inc + 1
        If n = 39 Then n = 0
        commqualityTZ.status.Text = BinToDec(kp1)
        baru.TextBox1.Text = BinToDec(kp)
        If jumlahsalah = 0 Then
            If jumlahsalah1 = 0 Then
                Button5.Visible = False
                My.Computer.Audio.Stop()
            End If
        End If
        If dei_dbase.Enabled = False Then 'mengaktifkan timer record data DEI ke database
            If baru.TextBox1.Text >= "67085327" Then dei_dbase.Enabled = True 'timer akan aktif jika semua alamat telah memiliki nilai. angka 67107831 didapatkan dari konversi HEX variable Kp ke Decimal
        End If
        If TZDbase.Enabled = False Then 'mengaktifkan timer record data TZ ke database
            If commqualityTZ.status.Text >= "19580" Then TZDbase.Enabled = True 'timer akan aktif jika semua alamat telah memiliki nilai. angka 20478 didapatkan dari konversi HEX variable Kp1 ke Decimal
        End If

        'If flag = 1 Then
        '    bacaON()
        '    bacaOFF()
        '    Flag = 0
        'End If
        'm = 1
        'For ite = 1 To 30 Step 1
        '    If Mid(tanggal.Text, 12, 8) = var10(ite) + ":00" Then
        '        rx = 1
        '        digout(m) = ite
        '        m = m + 1
        '        DEISV.Enabled = True
        '    End If
        '    'If rx = 1 Then
        '    '    If ite = 30 Then DEISV.Enabled = True
        '    'End If
        'Next

        'For ite = 1 To 30 Step 1
        '    If Mid(tanggal.Text, 12, 8) = var11(ite) + ":00" Then
        '        rx = 2
        '        digout(m) = ite
        '        m = m + 1
        '        DEISV.Enabled = True
        '    End If
        '    'If rx = 2 Then
        '    '    If ite = 30 Then DEISV.Enabled = True
        '    'End If
        'Next
        'code dibawah untuk memasukkan  slave address yang membaca nilai dibawah 0 ke tulisruang

        'For m = 1 To m = 39 Step 1
        '    cek(m) = address(m)
        'Next
        tulisruang(100)
        tulisruang(101)
        tulisruang(102)
        tulisruang(103)
        tulisruang(104)
        tulisruang(105)
        tulisruang(106)
        tulisruang(107)
        tulisruang(108)
        tulisruang(109)
        tulisruang(110)
        tulisruang(111)
        tulisruang(112)
        tulisruang(113)
        tulisruang(114)
        tulisruang(115)
        tulisruang(116)
        tulisruang(117)
        tulisruang(118)
        tulisruang(119)
        tulisruang(120)
        tulisruang(121)
        tulisruang(122)
        tulisruang(123)
        tulisruang(124)
        tulisruang(125)
        tulisruang(1)
        tulisruang(2)
        tulisruang(3)
        tulisruang(4)
        tulisruang(5)
        tulisruang(6)
        tulisruang(7)
        tulisruang(8)
        tulisruang(9)
        tulisruang(10)
        tulisruang(11)
        tulisruang(12)
        tulisruang(15)

        If inc = 2 Then inc = 0

        If DateAndTime.Now.Day <> 1 Then
            hari = 0
        End If
        If DateAndTime.Now.Day = 1 Then
            If hari <> 1 Then
                backup()
            End If
        End If
    End Sub
    Private Sub backup()
        System.IO.File.Copy("e:\pens\try.mdb", String.Format("e:\backup data\backup ruangan\{0:dd-MM-yyyy hh.mm}.mdb", DateAndTime.Now))
        System.IO.File.Copy("e:\pens\biocolddb.mdb", String.Format("e:\backup data\backup cold storage\{0:dd-MM-yyyy hh.mm}.mdb", DateAndTime.Now))
        hapus()
        hari = 1
    End Sub
    Private Sub hapus()
        Dim provider, dataFile, providertz, dataFiletz As String
        'hapus database hvac
        Dim myConnection As OleDbConnection = New OleDbConnection
        provider = "Provider=Microsoft.ACE.OleDB.12.0;Data Source="
        dataFile = "E:\PENS\try.mdb"
        myConnection.ConnectionString = provider & dataFile
        myConnection.Open()
        Dim str As String
        str = "DELETE FROM Table1"
        Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
        cmd.ExecuteNonQuery()
        myConnection.Close()

        'hapus database ruang simpan
        Dim myConnectiontz As OleDbConnection = New OleDbConnection
        providertz = "Provider=Microsoft.ACE.OleDB.12.0;Data Source="
        dataFiletz = "E:\PENS\BiocoldDB.mdb"
        myConnection.ConnectionString = providertz & dataFiletz
        myConnection.Open()
        Dim str1 As String
        str1 = "DELETE FROM Biocold"
        Dim cmd1 As OleDbCommand = New OleDbCommand(str, myConnectiontz)
        'cmd1.ExecuteNonQuery()
       ' myConnection.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        save6 = 0
        My.Computer.Audio.Stop()
        Button5.Visible = False
    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        tes2 = 1
        My.Computer.Audio.Stop()
        Button16.Visible = False

    End Sub
    Private Sub tulisruang(alamat As Integer)
        'apabila silent(n) menunjukkan apakah slave address membaca nilai di bawah 0 atau sebaliknya 
        Select Case alamat
            'Case 1 : If silent(1) = 1 Then
            '  If inc Mod 2 = 0 Then Label47.BackColor = Color.Red Else Label47.BackColor = Color.White
            'ElseIf silent(1) = 0 Then
            ' Label47.BackColor = Color.Transparent
            'End If
            Case 2 : If silent(2) = 1 Then
                    If inc Mod 2 = 0 Then Label46.BackColor = Color.Red Else Label46.BackColor = Color.White
                ElseIf silent(2) = 0 Then
                    Label46.BackColor = Color.Transparent
                End If
            '  Case 3 : If silent(3) = 1 Then
            'If inc Mod 2 = 0 Then Label44.BackColor = Color.Red Else Label44.BackColor = Color.White
            '  ElseIf silent(3) = 0 Then
            'Label44.BackColor = Color.Transparent
            '   End If

            Case 4 : If silent(4) = 1 Then
                    If inc Mod 2 = 0 Then Label42.BackColor = Color.Red Else Label42.BackColor = Color.White
                ElseIf silent(4) = 0 Then
                    Label42.BackColor = Color.Transparent
                End If
            Case 5 : If silent(5) = 1 Then
                    If inc Mod 2 = 0 Then Label30.BackColor = Color.Red Else Label30.BackColor = Color.White
                ElseIf silent(5) = 0 Then
                    Label30.BackColor = Color.Transparent
                End If

            Case 6 : If silent(6) = 1 Then
                    If inc Mod 2 = 0 Then Label31.BackColor = Color.Red Else Label31.BackColor = Color.White
                ElseIf silent(6) = 0 Then
                    Label31.BackColor = Color.Transparent
                End If

            Case 7 : If silent(7) = 1 Then
                    If inc Mod 2 = 0 Then Label33.BackColor = Color.Red Else Label33.BackColor = Color.White
                ElseIf silent(7) = 0 Then
                    Label33.BackColor = Color.Transparent
                End If

            ' Case 8 : If silent(8) = 1 Then
            ' If inc Mod 2 = 0 Then Label18.BackColor = Color.Red Else Label18.BackColor = Color.White
            '  ElseIf silent(8) = 0 Then
            'Label18.BackColor = Color.Transparent
            ' End If

            Case 9 : If silent(9) = 1 Then
                    If inc Mod 2 = 0 Then Label36.BackColor = Color.Red Else Label36.BackColor = Color.White
                ElseIf silent(9) = 0 Then
                    Label36.BackColor = Color.Transparent
                End If

            ' Case 10 : If silent(10) = 1 Then
            'If inc Mod 2 = 0 Then Label40.BackColor = Color.Red Else Label40.BackColor = Color.White
            '    ElseIf silent(10) = 0 Then
            ' Label40.BackColor = Color.Transparent
            '   End If

            Case 11 : If silent(11) = 1 Then
                    If inc Mod 2 = 0 Then Label38.BackColor = Color.Red Else Label38.BackColor = Color.White
                ElseIf silent(11) = 0 Then
                    Label38.BackColor = Color.Transparent
                End If

            Case 12 : If silent(12) = 1 Then
                    If inc Mod 2 = 0 Then Label45.BackColor = Color.Red Else Label45.BackColor = Color.White
                ElseIf silent(12) = 0 Then
                    Label45.BackColor = Color.Transparent
                End If

            Case 15 : If silent(15) = 1 Then
                    If inc Mod 2 = 0 Then Label43.BackColor = Color.Red Else Label43.BackColor = Color.White
                ElseIf silent(15) = 0 Then
                    Label43.BackColor = Color.Transparent
                End If

            Case 100 : If silent(100) = 1 Then
                    If inc Mod 2 = 0 Then Label32.BackColor = Color.Red Else Label32.BackColor = Color.White
                Else
                    Label32.BackColor = Color.Transparent
                End If

            Case 101 : If silent(101) = 1 Then
                    If inc Mod 2 = 0 Then Label25.BackColor = Color.Red Else Label25.BackColor = Color.White
                Else
                    Label25.BackColor = Color.Transparent
                End If

            Case 102 : If silent(102) = 1 Then
                    If inc Mod 2 = 0 Then Label53.BackColor = Color.Red Else Label53.BackColor = Color.White
                Else
                    Label53.BackColor = Color.Transparent
                End If

            Case 103 : If silent(103) = 1 Then
                    If inc Mod 2 = 0 Then Label52.BackColor = Color.Red Else Label52.BackColor = Color.White
                Else
                    Label52.BackColor = Color.Transparent
                End If

            Case 104 : If silent(104) = 1 Then
                    If inc Mod 2 = 0 Then Label11.BackColor = Color.Red Else Label11.BackColor = Color.White
                Else
                    Label11.BackColor = Color.Transparent
                End If

            Case 105 : If silent(105) = 1 Then
                    If inc Mod 2 = 0 Then Label12.BackColor = Color.Red Else Label12.BackColor = Color.White
                Else
                    Label12.BackColor = Color.Transparent
                End If
            Case 106 : If silent(106) = 1 Then
                    If inc Mod 2 = 0 Then Label22.BackColor = Color.Red Else Label22.BackColor = Color.White
                Else
                    Label22.BackColor = Color.Transparent
                End If
            Case 107 : If silent(107) = 1 Then
                    If inc Mod 2 = 0 Then Label19.BackColor = Color.Red Else Label19.BackColor = Color.White
                Else
                    Label19.BackColor = Color.Transparent
                End If
            Case 108 : If silent(108) = 1 Then
                    If inc Mod 2 = 0 Then Label20.BackColor = Color.Red Else Label20.BackColor = Color.White
                Else
                    Label20.BackColor = Color.Transparent
                End If

            Case 109 : If silent(109) = 1 Then
                    If inc Mod 2 = 0 Then Label21.BackColor = Color.Red Else Label21.BackColor = Color.White
                Else
                    Label21.BackColor = Color.Transparent
                End If

            Case 110 : If silent(110) = 1 Then
                    If inc Mod 2 = 0 Then Label9.BackColor = Color.Red Else Label9.BackColor = Color.White
                Else
                    Label9.BackColor = Color.Transparent
                End If

            Case 111 : If silent(111) = 1 Then
                    If inc Mod 2 = 0 Then Label23.BackColor = Color.Red Else Label23.BackColor = Color.White
                Else
                    Label23.BackColor = Color.Transparent
                End If
            Case 112 : If silent(112) = 1 Then
                    If inc Mod 2 = 0 Then Label14.BackColor = Color.Red Else Label14.BackColor = Color.White
                Else
                    Label14.BackColor = Color.Transparent
                End If

            Case 113 : If silent(113) = 1 Then
                    If inc Mod 2 = 0 Then Label10.BackColor = Color.Red Else Label10.BackColor = Color.White
                Else
                    Label10.BackColor = Color.Transparent
                End If
            Case 114 : If silent(114) = 1 Then
                    If inc Mod 2 = 0 Then Label15.BackColor = Color.Red Else Label15.BackColor = Color.White
                Else
                    Label15.BackColor = Color.Transparent
                End If
            Case 115 : If silent(115) = 1 Then
                    If inc Mod 2 = 0 Then Label16.BackColor = Color.Red Else Label16.BackColor = Color.White
                Else
                    Label16.BackColor = Color.Transparent
                End If
            Case 116 : If silent(116) = 1 Then
                    If inc Mod 2 = 0 Then Label26.BackColor = Color.Red Else Label26.BackColor = Color.White
                Else
                    Label26.BackColor = Color.Transparent
                End If

            Case 117 : If silent(117) = 1 Then
                    If inc Mod 2 = 0 Then Label24.BackColor = Color.Red Else Label24.BackColor = Color.White
                Else
                    Label24.BackColor = Color.Transparent
                End If

            Case 118 : If silent(118) = 1 Then
                    If inc Mod 2 = 0 Then Label2.BackColor = Color.Red Else Label2.BackColor = Color.White
                Else
                    Label2.BackColor = Color.Transparent
                End If

            Case 119 : If silent(119) = 1 Then
                    If inc Mod 2 = 0 Then Label1.BackColor = Color.Red Else Label1.BackColor = Color.White
                Else
                    Label1.BackColor = Color.Transparent
                End If

            Case 120 : If silent(120) = 1 Then
                    If inc Mod 2 = 0 Then Label3.BackColor = Color.Red Else Label3.BackColor = Color.White
                Else
                    Label3.BackColor = Color.Transparent
                End If

            Case 121 : If silent(121) = 1 Then
                    If inc Mod 2 = 0 Then Label4.BackColor = Color.Red Else Label4.BackColor = Color.White
                Else
                    Label4.BackColor = Color.Transparent
                End If

            Case 122 : If silent(122) = 1 Then
                    If inc Mod 2 = 0 Then Label8.BackColor = Color.Red Else Label8.BackColor = Color.White
                Else
                    Label8.BackColor = Color.Transparent
                End If

            Case 123 : If silent(123) = 1 Then
                    If inc Mod 2 = 0 Then Label7.BackColor = Color.Red Else Label7.BackColor = Color.White
                Else
                    Label7.BackColor = Color.Transparent
                End If

            Case 124 : If silent(124) = 1 Then
                    If inc Mod 2 = 0 Then Label6.BackColor = Color.Red Else Label6.BackColor = Color.White
                Else
                    Label6.BackColor = Color.Transparent
                End If

            Case 125 : If silent(125) = 1 Then
                    If inc Mod 2 = 0 Then Label5.BackColor = Color.Red Else Label5.BackColor = Color.White
                Else
                    Label5.BackColor = Color.Transparent
                End If

        End Select
    End Sub
    Private Sub alamatsalah1()
        jumlahsalah1 = 0 'awal siklus selalu bernilai 0
        If BIOCOLD032.Text <> "OFFLINE" Then
            If Val(CDbl(BIOCOLD032.Text)) < Val(2) Then
                jumlahsalah1 = jumlahsalah1 + 1
            ElseIf Val(CDbl(BIOCOLD032.Text)) > Val(8) Then
                jumlahsalah1 = jumlahsalah1 + 1
            End If
        End If
        If BIOCOLD037A.Text <> "OFFLINE" Then
            If Val(CDbl(BIOCOLD037A.Text)) < Val(2) Then
                jumlahsalah1 = jumlahsalah1 + 1
            ElseIf Val(CDbl(BIOCOLD037A.Text)) > Val(8) Then
                jumlahsalah1 = jumlahsalah1 + 1
            End If
        End If
        If BIOCOLD037B.Text <> "OFFLINE" Then
            If Val(CDbl(BIOCOLD037B.Text)) < Val(2) Then
                jumlahsalah1 = jumlahsalah1 + 1
            ElseIf Val(CDbl(BIOCOLD037B.Text)) > Val(8) Then
                jumlahsalah1 = jumlahsalah1 + 1
            End If
        End If
        If BIOCOLD038A.Text <> "OFFLINE" Then
            If Val(CDbl(BIOCOLD038A.Text)) < Val(2) Then
                jumlahsalah1 = jumlahsalah1 + 1
            ElseIf Val(CDbl(BIOCOLD038A.Text)) > Val(8) Then
                jumlahsalah1 = jumlahsalah1 + 1
            End If
        End If
        If BIOCOLD038B.Text <> "OFFLINE" Then
            If Val(CDbl(BIOCOLD038B.Text)) < Val(2) Then
                jumlahsalah1 = jumlahsalah1 + 1
            ElseIf Val(CDbl(BIOCOLD038B.Text)) > Val(8) Then
                jumlahsalah1 = jumlahsalah1 + 1
            End If
        End If
        If BIOCOLD039.Text <> "OFFLINE" Then
            If Val(CDbl(BIOCOLD039.Text)) < Val(2) Then
                jumlahsalah1 = jumlahsalah1 + 1
            ElseIf Val(CDbl(BIOCOLD039.Text)) > Val(8) Then
                jumlahsalah1 = jumlahsalah1 + 1
            End If
        End If
        If BIOCOLD040.Text <> "OFFLINE" Then
            If Val(CDbl(BIOCOLD040.Text)) < Val(2) Then
                jumlahsalah1 = jumlahsalah1 + 1
            ElseIf Val(CDbl(BIOCOLD040.Text)) > Val(8) Then
                jumlahsalah1 = jumlahsalah1 + 1
            End If
        End If
        If BIOCOLD041A.Text <> "OFFLINE" Then
            If Val(CDbl(BIOCOLD041A.Text)) < Val(2) Then
                jumlahsalah1 = jumlahsalah1 + 1
            ElseIf Val(CDbl(BIOCOLD041A.Text)) > Val(8) Then
                jumlahsalah1 = jumlahsalah1 + 1
            End If
        End If
        If BIOCOLD041B.Text <> "OFFLINE" Then
            If Val(CDbl(BIOCOLD041B.Text)) < Val(2) Then
                jumlahsalah1 = jumlahsalah1 + 1
            ElseIf Val(CDbl(BIOCOLD041B.Text)) > Val(8) Then
                jumlahsalah1 = jumlahsalah1 + 1
            End If
        End If
        If BIOCOLD042A.Text <> "OFFLINE" Then
            If Val(CDbl(BIOCOLD042A.Text)) < Val(2) Then
                jumlahsalah1 = jumlahsalah1 + 1
            ElseIf Val(CDbl(BIOCOLD042A.Text)) > Val(8) Then
                jumlahsalah1 = jumlahsalah1 + 1
            End If
        End If
        If BIOCOLD042B.Text <> "OFFLINE" Then
            If Val(CDbl(BIOCOLD042B.Text)) < Val(2) Then
                jumlahsalah1 = jumlahsalah1 + 1
            ElseIf Val(CDbl(BIOCOLD042B.Text)) > Val(8) Then
                jumlahsalah1 = jumlahsalah1 + 1
            End If
        End If
        If BIOCOLD043.Text <> "OFFLINE" Then
            If Val(CDbl(BIOCOLD043.Text)) > Val(0) Then
                jumlahsalah1 = jumlahsalah1 + 1
            ElseIf Val(CDbl(BIOCOLD043.Text)) < Val(-23) Then
                jumlahsalah1 = jumlahsalah1 + 1
            End If
        End If
        If BIOCOLD044.Text <> "OFFLINE" Then
            If Val(CDbl(BIOCOLD044.Text)) < Val(15) Then
                jumlahsalah1 = jumlahsalah1 + 1
            ElseIf Val(CDbl(BIOCOLD044.Text)) > Val(20) Then
                jumlahsalah1 = jumlahsalah1 + 1
            End If
        End If
    End Sub
    Private Sub alamatsalah()
        Dim logic(52) As Integer
        'code di bawah untuk scan nilai setiap label.text apakah nilainya di bawah/di atas batas atau tidak. apabila di bawah 0 jumlahsalah bertambah satu
        jumlahsalah = 0 'awal siklus selalu bernilai 0
        'If BIOCOLD032.Text <> "OFFLINE" Then
        '    If Val(CDbl(BIOCOLD032.Text)) < Val(2) Then
        '        jumlahsalah = jumlahsalah + 1
        '    ElseIf Val(CDbl(BIOCOLD032.Text)) > Val(8) Then
        '        jumlahsalah = jumlahsalah + 1
        '    End If
        'End If
        'If BIOCOLD037A.Text <> "OFFLINE" Then
        '    If Val(CDbl(BIOCOLD037A.Text)) < Val(2) Then
        '        jumlahsalah = jumlahsalah + 1
        '    ElseIf Val(CDbl(BIOCOLD037A.Text)) > Val(8) Then
        '        jumlahsalah = jumlahsalah + 1
        '    End If
        'End If
        'If BIOCOLD037B.Text <> "OFFLINE" Then
        '    If Val(CDbl(BIOCOLD037B.Text)) < Val(2) Then
        '        jumlahsalah = jumlahsalah + 1
        '    ElseIf Val(CDbl(BIOCOLD037B.Text)) > Val(8) Then
        '        jumlahsalah = jumlahsalah + 1
        '    End If
        'End If
        'If BIOCOLD038A.Text <> "OFFLINE" Then
        '    If Val(CDbl(BIOCOLD038A.Text)) < Val(2) Then
        '        jumlahsalah = jumlahsalah + 1
        '    ElseIf Val(CDbl(BIOCOLD038A.Text)) > Val(8) Then
        '        jumlahsalah = jumlahsalah + 1
        '    End If
        'End If
        'If BIOCOLD038B.Text <> "OFFLINE" Then
        '    If Val(CDbl(BIOCOLD038B.Text)) < Val(2) Then
        '        jumlahsalah = jumlahsalah + 1
        '    ElseIf Val(CDbl(BIOCOLD038B.Text)) > Val(8) Then
        '        jumlahsalah = jumlahsalah + 1
        '    End If
        'End If
        'If BIOCOLD039.Text <> "OFFLINE" Then
        '    If Val(CDbl(BIOCOLD039.Text)) < Val(2) Then
        '        jumlahsalah = jumlahsalah + 1
        '    ElseIf Val(CDbl(BIOCOLD039.Text)) > Val(8) Then
        '        jumlahsalah = jumlahsalah + 1
        '    End If
        'End If
        'If BIOCOLD040.Text <> "OFFLINE" Then
        '    If Val(CDbl(BIOCOLD040.Text)) < Val(2) Then
        '        jumlahsalah = jumlahsalah + 1
        '    ElseIf Val(CDbl(BIOCOLD040.Text)) > Val(8) Then
        '        jumlahsalah = jumlahsalah + 1
        '    End If
        'End If
        'If BIOCOLD041A.Text <> "OFFLINE" Then
        '    If Val(CDbl(BIOCOLD041A.Text)) < Val(2) Then
        '        jumlahsalah = jumlahsalah + 1
        '    ElseIf Val(CDbl(BIOCOLD041A.Text)) > Val(8) Then
        '        jumlahsalah = jumlahsalah + 1
        '    End If
        'End If
        'If BIOCOLD041B.Text <> "OFFLINE" Then
        '    If Val(CDbl(BIOCOLD041B.Text)) < Val(2) Then
        '        jumlahsalah = jumlahsalah + 1
        '    ElseIf Val(CDbl(BIOCOLD041B.Text)) > Val(8) Then
        '        jumlahsalah = jumlahsalah + 1
        '    End If
        'End If
        'If BIOCOLD042A.Text <> "OFFLINE" Then
        '    If Val(CDbl(BIOCOLD042A.Text)) < Val(2) Then
        '        jumlahsalah = jumlahsalah + 1
        '    ElseIf Val(CDbl(BIOCOLD042A.Text)) > Val(8) Then
        '        jumlahsalah = jumlahsalah + 1
        '    End If
        'End If
        'If BIOCOLD042B.Text <> "OFFLINE" Then
        '    If Val(CDbl(BIOCOLD042B.Text)) < Val(2) Then
        '        jumlahsalah = jumlahsalah + 1
        '    ElseIf Val(CDbl(BIOCOLD042B.Text)) > Val(8) Then
        '        jumlahsalah = jumlahsalah + 1
        '    End If
        'End If
        'If BIOCOLD043.Text <> "OFFLINE" Then
        '    If Val(CDbl(BIOCOLD043.Text)) > Val(0) Then
        '        jumlahsalah = jumlahsalah + 1
        '    ElseIf Val(CDbl(BIOCOLD043.Text)) < Val(-23) Then
        '        jumlahsalah = jumlahsalah + 1
        '    End If
        'End If
        'If BIOCOLD044.Text <> "OFFLINE" Then
        '    If Val(CDbl(BIOCOLD044.Text)) < Val(15) Then
        '        jumlahsalah = jumlahsalah + 1
        '    ElseIf Val(CDbl(BIOCOLD044.Text)) > Val(20) Then
        '        jumlahsalah = jumlahsalah + 1
        '    End If
        'End If

        'Fail UI1
        If slave_100.Text <> "Offline" Then
            If slave_100.Text <> "Fail UI1" Then
                If Val(CDbl(slave_100.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(1) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_100.Text)) < Val(27) Then
                    logic(1) = 0
                End If
            End If
        End If
        If slave_101.Text <> "Offline" Then
            If slave_101.Text <> "Fail UI1" Then
                If Val(CDbl(slave_101.Text)) > Val(27) Then
                    logic(2) = 1
                    jumlahsalah = jumlahsalah + 1
                    save5 = 1
                ElseIf Val(CDbl(slave_101.Text)) < Val(27) Then
                    logic(2) = 0
                End If
            End If
        End If
        If slave_102.Text <> "Offline" Then
            If slave_102.Text <> "Fail UI1" Then
                If Val(CDbl(slave_102.Text)) > Val(CDbl(0.54)) Then
                    logic(3) = 1
                    jumlahsalah = jumlahsalah + 1
                    save5 = 1
                ElseIf Val(CDbl(slave_102.Text)) < Val(CDbl(0.36)) Then
                    logic(3) = 1
                    jumlahsalah = jumlahsalah + 1
                    save5 = 1
                Else
                    logic(3) = 0
                End If
            End If
        End If
        If slave_103.Text <> "Offline" Then
            If slave_103.Text <> "Fail UI1" Then
                If Val(CDbl(slave_103.Text)) > Val(CDbl(0.54)) Then
                    logic(4) = 1
                    jumlahsalah = jumlahsalah + 1
                    save5 = 1
                ElseIf Val(CDbl(slave_103.Text)) < Val(CDbl(0.36)) Then
                    logic(4) = 1
                    jumlahsalah = jumlahsalah + 1
                    save5 = 1
                Else
                    logic(4) = 0
                End If
            End If
        End If
        If slave_104.Text <> "Offline" Then
            If slave_104.Text <> "Fail UI1" Then
                If Val(CDbl(slave_104.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(5) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_104.Text)) < Val(27) Then
                    logic(5) = 0
                End If
            End If
        End If
        If slave_105.Text <> "Offline" Then
            If slave_105.Text <> "Fail UI1" Then
                If Val(CDbl(slave_105.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(6) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_105.Text)) < Val(27) Then
                    logic(6) = 0
                End If
            End If
        End If
        If slave_106.Text <> "Offline" Then
            If slave_106.Text <> "Fail UI1" Then
                If Val(CDbl(slave_106.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(7) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_106.Text)) < Val(27) Then
                    logic(7) = 0
                End If
            End If
        End If
        If slave_107.Text <> "Offline" Then
            If slave_107.Text <> "Fail UI1" Then
                If Val(CDbl(slave_107.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(8) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_107.Text)) < Val(27) Then
                    logic(8) = 0
                End If
            End If
        End If
        If slave_108.Text <> "Offline" Then
            If slave_108.Text <> "Fail UI1" Then
                If Val(CDbl(slave_108.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(9) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_108.Text)) <= Val(27) Then
                    logic(9) = 0
                End If
            End If
        End If
        If slave_109.Text <> "Offline" Then
            If slave_109.Text <> "Fail UI1" Then
                If Val(CDbl(slave_109.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(10) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_109.Text)) <= Val(27) Then
                    logic(10) = 0
                End If
            End If
        End If
        If slave_110.Text <> "Offline" Then
            If slave_110.Text <> "Fail UI1" Then
                If Val(CDbl(slave_110.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(11) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_110.Text)) <= Val(27) Then
                    logic(11) = 0
                End If
            End If
        End If
        If slave_111.Text <> "Offline" Then
            If slave_111.Text <> "Fail UI1" Then
                If Val(CDbl(slave_111.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(12) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_111.Text)) <= Val(27) Then
                    logic(12) = 0
                End If
            End If
        End If
        If slave_112.Text <> "Offline" Then
            If slave_112.Text <> "Fail UI1" Then
                If Val(CDbl(slave_112.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(13) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_112.Text)) <= Val(27) Then
                    logic(13) = 0
                End If
            End If
        End If
        If slave_113.Text <> "Offline" Then
            If slave_113.Text <> "Fail UI1" Then
                If Val(CDbl(slave_113.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(14) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_113.Text)) <= Val(27) Then
                    logic(14) = 0
                End If
            End If
        End If
        If slave_114.Text <> "Offline" Then
            If slave_114.Text <> "Fail UI1" Then
                If Val(CDbl(slave_114.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(15) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_114.Text)) <= Val(27) Then
                    logic(15) = 0
                End If
            End If
        End If
        If slave_115.Text <> "Offline" Then
            If slave_115.Text <> "Fail UI1" Then
                If Val(CDbl(slave_115.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(16) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_115.Text)) <= Val(27) Then
                    logic(16) = 0
                End If
            End If
        End If
        If slave_116.Text <> "Offline" Then
            If slave_116.Text <> "Fail UI1" Then
                If Val(CDbl(slave_116.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(17) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_116.Text)) <= Val(27) Then
                    logic(17) = 0
                End If
            End If
        End If
        If slave_117.Text <> "Offline" Then
            If slave_117.Text <> "Fail UI1" Then
                If Val(CDbl(slave_117.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(18) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_117.Text)) <= Val(27) Then
                    logic(18) = 0
                End If
            End If
        End If
        If slave_118.Text <> "Offline" Then
            If slave_118.Text <> "Fail UI1" Then
                If Val(CDbl(slave_118.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(19) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_118.Text)) <= Val(27) Then
                    logic(19) = 0
                End If
            End If
        End If
        If slave_119.Text <> "Offline" Then
            If slave_119.Text <> "Fail UI1" Then
                If Val(CDbl(slave_119.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(20) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_119.Text)) <= Val(27) Then
                    logic(20) = 0
                End If
            End If
        End If
        If slave_120.Text <> "Offline" Then
            If slave_120.Text <> "Fail UI1" Then
                If Val(CDbl(slave_120.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(21) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_120.Text)) <= Val(27) Then
                    logic(21) = 0
                End If
            End If
        End If
        If slave_121.Text <> "Offline" Then
            If slave_121.Text <> "Fail UI1" Then
                If Val(CDbl(slave_121.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(22) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_121.Text)) <= Val(27) Then
                    logic(22) = 0
                End If
            End If
        End If
        If slave_122.Text <> "Offline" Then
            If slave_122.Text <> "Fail UI1" Then
                If Val(CDbl(slave_122.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(23) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_122.Text)) <= Val(27) Then
                    logic(23) = 0
                End If
            End If
        End If
        If slave_123.Text <> "Offline" Then
            If slave_123.Text <> "Fail UI1" Then
                If Val(CDbl(slave_123.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(24) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_123.Text)) <= Val(27) Then
                    logic(24) = 0
                End If
            End If
        End If
        If slave_124.Text <> "Offline" Then
            If slave_124.Text <> "Fail UI1" Then
                If Val(CDbl(slave_124.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(25) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_124.Text)) <= Val(27) Then
                    logic(25) = 0
                End If
            End If
        End If
        If slave_125.Text <> "Offline" Then
            If slave_125.Text <> "Fail UI1" Then
                If Val(CDbl(slave_125.Text)) > Val(27) Then
                    jumlahsalah = jumlahsalah + 1
                    logic(26) = 1
                    save5 = 1
                ElseIf Val(CDbl(slave_125.Text)) <= Val(27) Then
                    logic(26) = 0
                End If
            End If
        End If

        'Fail UI2
        If Pa_100.Text <> "Offline" Then
            If Pa_100.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_100.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(27) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_100.Text)) > 0 Then
                    logic(27) = 0
                End If
            End If
        End If
        If Pa_101.Text <> "Offline" Then
            If Pa_101.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_101.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(28) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_101.Text)) > 0 Then
                    logic(28) = 0
                End If
            End If
        End If
        If Pa_102.Text <> "Offline" Then
            If Pa_102.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_102.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(29) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_102.Text)) > 0 Then
                    logic(29) = 0
                End If
            End If
        End If
        If Pa_103.Text <> "Offline" Then
            If Pa_103.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_103.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(30) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_103.Text)) > 0 Then
                    logic(30) = 0
                End If
            End If
        End If
        If Pa_104.Text <> "Offline" Then
            If Pa_104.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_104.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(31) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_104.Text)) > 0 Then
                    logic(31) = 0
                End If
            End If
        End If
        If Pa_105.Text <> "Offline" Then
            If Pa_105.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_105.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(32) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_105.Text)) > 0 Then
                    logic(32) = 0
                End If
            End If
        End If
        If Pa_106.Text <> "Offline" Then
            If Pa_106.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_106.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(33) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_106.Text)) > 0 Then
                    logic(33) = 0
                End If
            End If
        End If
        If Pa_107.Text <> "Offline" Then
            If Pa_107.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_107.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(34) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_107.Text)) > 0 Then
                    logic(34) = 0
                End If
            End If
        End If
        If Pa_108.Text <> "Offline" Then
            If Pa_108.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_108.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(35) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_108.Text)) > 0 Then
                    logic(35) = 0
                End If
            End If
        End If
        If Pa_109.Text <> "Offline" Then
            If Pa_109.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_109.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(36) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_109.Text)) > 0 Then
                    logic(36) = 0
                End If
            End If
        End If
        If Pa_110.Text <> "Offline" Then
            If Pa_110.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_110.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(37) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_110.Text)) > 0 Then
                    logic(37) = 0
                End If
            End If
        End If
        If Pa_111.Text <> "Offline" Then
            If Pa_111.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_111.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(38) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_111.Text)) > 0 Then
                    logic(38) = 0
                End If
            End If
        End If
        If Pa_112.Text <> "Offline" Then
            If Pa_112.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_112.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(39) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_112.Text)) > 0 Then
                    logic(39) = 0
                End If
            End If
        End If
        If Pa_113.Text <> "Offline" Then
            If Pa_113.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_113.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(40) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_113.Text)) > 0 Then
                    logic(40) = 0
                End If
            End If
        End If
        If Pa_114.Text <> "Offline" Then
            If Pa_114.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_114.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(41) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_114.Text)) > 0 Then
                    logic(41) = 0
                End If
            End If
        End If
        If Pa_115.Text <> "Offline" Then
            If Pa_115.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_115.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(42) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_115.Text)) > 0 Then
                    logic(42) = 0
                End If
            End If
        End If
        If Pa_116.Text <> "Offline" Then
            If Pa_116.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_116.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(43) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_116.Text)) > 0 Then
                    logic(43) = 0
                End If
            End If
        End If
        If Pa_117.Text <> "Offline" Then
            If Pa_117.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_117.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(44) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_117.Text)) > 0 Then
                    logic(44) = 0
                End If
            End If
        End If
        If Pa_118.Text <> "Offline" Then
            If Pa_118.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_118.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(45) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_118.Text)) > 0 Then
                    logic(45) = 0
                End If
            End If
        End If
        If Pa_119.Text <> "Offline" Then
            If Pa_119.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_119.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(46) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_119.Text)) > 0 Then
                    logic(46) = 0
                End If
            End If
        End If
        If Pa_120.Text <> "Offline" Then
            If Pa_120.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_120.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(47) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_120.Text)) > 0 Then
                    logic(47) = 0
                End If
            End If
        End If
        If Pa_121.Text <> "Offline" Then
            If Pa_121.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_121.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(48) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_121.Text)) > 0 Then
                    logic(48) = 0
                End If
            End If
        End If
        If Pa_122.Text <> "Offline" Then
            If Pa_122.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_122.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(49) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_122.Text)) > 0 Then
                    logic(49) = 0
                End If
            End If
        End If
        If Pa_123.Text <> "Offline" Then
            If Pa_123.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_123.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(50) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_123.Text)) > 0 Then
                    logic(50) = 0
                End If
            End If
        End If
        If Pa_124.Text <> "Offline" Then
            If Pa_124.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_124.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(51) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_124.Text)) > 0 Then
                    logic(51) = 0
                End If
            End If
        End If
        If Pa_125.Text <> "Offline" Then
            If Pa_125.Text <> "Fail UI2" Then
                If Val(CDbl(Pa_125.Text)) < 0 Then
                    jumlahsalah = jumlahsalah + 1
                    logic(52) = 1
                    save5 = 1
                ElseIf Val(CDbl(Pa_125.Text)) > 0 Then
                    logic(52) = 0
                End If
            End If
        End If
        silent(100) = Val(logic(1)) + Val(logic(27))
        silent(101) = Val(logic(2)) + Val(logic(28))
        silent(102) = Val(logic(3)) + Val(logic(29))
        silent(103) = Val(logic(4)) + Val(logic(30))
        silent(104) = Val(logic(5)) + Val(logic(31))
        silent(105) = Val(logic(6)) + Val(logic(32))
        silent(106) = Val(logic(7)) + Val(logic(33))
        silent(107) = Val(logic(8)) + Val(logic(34))
        silent(108) = Val(logic(9)) + Val(logic(35))
        silent(109) = Val(logic(10)) + Val(logic(36))
        silent(110) = Val(logic(11)) + Val(logic(37))
        silent(111) = Val(logic(12)) + Val(logic(38))
        silent(112) = Val(logic(13)) + Val(logic(39))
        silent(113) = Val(logic(14)) + Val(logic(40))
        silent(114) = Val(logic(15)) + Val(logic(41))
        silent(115) = Val(logic(16)) + Val(logic(42))
        silent(116) = Val(logic(17)) + Val(logic(43))
        silent(117) = Val(logic(18)) + Val(logic(44))
        silent(118) = Val(logic(19)) + Val(logic(45))
        silent(119) = Val(logic(20)) + Val(logic(46))
        silent(120) = Val(logic(21)) + Val(logic(47))
        silent(121) = Val(logic(22)) + Val(logic(48))
        silent(122) = Val(logic(23)) + Val(logic(49))
        silent(123) = Val(logic(24)) + Val(logic(50))
        silent(124) = Val(logic(25)) + Val(logic(51))
        silent(125) = Val(logic(26)) + Val(logic(52))
    End Sub
    Function BinToDec(ByVal Bin As String)
        'mengubah biner ke decimal
        Dim dec As String = Nothing
        Dim length As Integer = Len(Bin)
        Dim temp As Integer = Nothing
        Dim x As Integer = Nothing

        For x = 1 To length
            temp = Val(Mid(Bin, length, 1))
            length = length - 1
            If temp <> "0" Then
                dec += (2 ^ (x - 1))
            End If
        Next
        Return dec
    End Function
    Private Sub TimerPV_Tick(sender As Object, e As EventArgs) Handles TimerPV.Tick
        Dim PauseTime, Start As Double
        Dim p As Integer = 0
        If MSComm2.PortOpen = True Then
            poll = poll + 1

            Command_Head = "P"
            If Command_Head = "P" Then
                Select Case Val(poll)
                    Case 1, 10, 23, 32 : Bcode = "j"
                    Case 2, 13, 20, 31 : Bcode = "i"
                    Case 3, 12, 21, 30 : Bcode = "h"
                    Case 4, 15, 26 : Bcode = "o"
                    Case 5, 14, 27 : Bcode = "n"
                    Case 6, 17, 24 : Bcode = "m"
                    Case 7, 16, 25 : Bcode = "l"
                    Case 8, 19 : Bcode = "c"
                    Case 9, 18 : Bcode = "b"
                    Case 11, 22 : Bcode = "k"
                    Case 28 : Bcode = "a"
                    Case 29 : Bcode = "`"
                End Select
            End If
            'code di bawah digunakan untuk memisahkan digit pertama dan kedua dari poll yang lebih dari 2 digit, karena jika tidak dipisah maka poll yang dikirim bersama command akan terbaca berbeda
            If poll >= 1 Then
                var1 = var1 + 1
                If var1 > 9 Then
                    var2 = var2 + 1
                    var1 = 0
                End If
            End If

            If poll > 9 Then
                commqualityTZ.commandPV.Text = Chr(2) + Chr(Asc(var2)) + Chr(Asc(var1)) + "RX" + Command_Head + "0" +
                Chr(3) + Bcode
            Else
                commqualityTZ.commandPV.Text = Chr(2) + "0" + Chr(Asc(var1)) + "RX" + Command_Head + "0" +
                Chr(3) + Bcode
            End If

            MSComm2.Output = commqualityTZ.commandPV.Text
            Do While MSComm2.OutBufferCount > 0
            Loop
            PauseTime = 1
            Start = Microsoft.VisualBasic.DateAndTime.Timer
            Do While (Microsoft.VisualBasic.DateAndTime.Timer < Start + PauseTime)

            Loop
            commqualityTZ.responsePV.Text = commqualityTZ.responsePV.Text + MSComm2.Input
            'program di bawah hampir sama dengan DEI, perbedaannya adalah beberapa address memiliki batas atas / bawah yang bebeda. dan ada address yang hanya memilii batas atas saja
            If commqualityTZ.responsePV.Text <> "" Then
                If poll = 6 Then 'untuk menyeleksi ruang 043 karena memiliki batas atas yang berbeda yaitu, > 0 derajat
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(0) Then
                        If address(n) <> poll Then 'apabila poll sekarang belum pernah salah, maka mengisi array baru
                            ' p = n + 1
                            Do
                                For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah salah
                                    If address(n) = poll Then
                                        n = n
                                        silent(poll) = 1
                                        address(n) = poll
                                        'tes2 = 0 'visible tombol true
                                        Exit Do
                                    End If
                                Next
                                ' mengisi array yang baru                               
                                For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah salah
                                    If address(n) = 0 Then
                                        n = n
                                        silent(poll) = 1
                                        address(n) = poll
                                        tes2 = 0 'visible tombol true
                                        ListBox1.Items.Add(KonversiRuang(poll) + " : " + tanggal.Text + " : " + "Temperature too high")
                                        DbaseAlarmOn(CInt(poll), CStr(tanggal.Text), CDbl(Val(Mid(commqualityTZ.responsePV.Text, 9, 5))), CStr("Temperature too high"))
                                        Exit Do
                                    End If
                                Next
                            Loop While False
                        Else ' address(n) sesuai dengan poll sekarang
                            address(n) = poll
                            silent(poll) = 1
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(-23) Then
                        If address(n) <> poll Then 'apabila poll sekarang belum pernah salah, maka mengisi array baru
                            p = n + 1
                            Do
                                For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah salah
                                    If address(n) = poll Then
                                        n = n
                                        silent(poll) = 1
                                        address(n) = poll
                                        'tes2 = 0 'visible tombol true
                                        Exit Do
                                    End If
                                Next
                                For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah salah
                                    If address(n) = 0 Then
                                        n = n
                                        silent(poll) = 1
                                        address(n) = poll
                                        tes2 = 0 'visible tombol true
                                        ListBox1.Items.Add(KonversiRuang(poll) + " : " + tanggal.Text + " : " + "Temperature too Low")
                                        DbaseAlarmOn(CInt(poll), CStr(tanggal.Text), CDbl(Val(Mid(commqualityTZ.responsePV.Text, 9, 5))), CStr("Temperature too Low"))
                                        Exit Do
                                    End If
                                Next
                            Loop While False
                        Else ' address(n) sesuai dengan poll sekarang
                            address(n) = poll
                            silent(poll) = 1
                        End If
                    Else
                        Do
                            For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah memiliki kamar atau sudah pernah salah
                                If address(n) = poll Then
                                    n = n
                                    address(n) = 0
                                    silent(poll) = 0 'mematikan kedip
                                    tes2 = 1 'menghilangkan tombol 'audio stop
                                    ListBox2.Items.Add(KonversiRuang(poll) + " : " + tanggal.Text)
                                    DbaseAlarmOff(CInt(poll), CStr(tanggal.Text), CDbl(Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10), CStr("Temperature Pass"))
                                    Exit Do
                                End If
                            Next
                            n = 0
                        Loop While False
                    End If
                ElseIf poll = 12 Then ' poll 12 memiliki range yang berbeda dengan yang lain
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(20) Then
                        'tes2 = 0 'visible tombol true 'audio play
                        If address(n) <> poll Then 'apabila poll sekarang belum pernah salah, maka mengisi array baru
                            p = n + 1
                            Do
                                For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah salah
                                    If address(n) = poll Then
                                        n = n
                                        silent(poll) = 1
                                        address(n) = poll
                                        'tes2 = 0 'visible tombol true
                                        Exit Do
                                    End If
                                Next
                                For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah salah
                                    If address(n) = 0 Then
                                        n = n
                                        silent(poll) = 1
                                        address(n) = poll
                                        tes2 = 0 'visible tombol true
                                        ListBox1.Items.Add(KonversiRuang(poll) + " : " + tanggal.Text + " : " + "Temperature too high")
                                        DbaseAlarmOn(CInt(poll), CStr(tanggal.Text), CDbl(Val(Mid(commqualityTZ.responsePV.Text, 9, 5))), CStr("Temperature too high"))
                                        Exit Do
                                    End If
                                Next
                            Loop While False
                        Else ' address(n) sesuai dengan poll sekarang
                            address(n) = poll
                            silent(poll) = 1
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(15) Then
                        'tes2 = 0 'visible tombol true 'audio play
                        If address(n) <> poll Then 'apabila poll sekarang belum pernah salah, maka mengisi array baru
                            p = n + 1
                            Do
                                For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah salah
                                    If address(n) = poll Then
                                        n = n
                                        silent(poll) = 1
                                        address(n) = poll
                                        tes2 = 0 'visible tombol true
                                        Exit Do
                                    End If
                                Next
                                For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah salah
                                    If address(n) = 0 Then
                                        n = n
                                        silent(poll) = 1
                                        address(n) = poll
                                        tes2 = 0 'visible tombol true
                                        ListBox1.Items.Add(KonversiRuang(poll) + " : " + tanggal.Text + " : " + "Temperature too Low")
                                        DbaseAlarmOn(CInt(poll), CStr(tanggal.Text), CDbl(Val(Mid(commqualityTZ.responsePV.Text, 9, 5))), CStr("Temperature too Low"))
                                        Exit Do
                                    End If
                                Next
                            Loop While False
                        Else ' address(n) sesuai dengan poll sekarang
                            address(n) = poll
                            silent(poll) = 1
                        End If
                    Else
                        Do
                            For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah memiliki kamar atau sudah pernah salah
                                If address(n) = poll Then
                                    n = n
                                    address(n) = 0
                                    silent(poll) = 0 'mematikan kedip
                                    tes2 = 1 'menghilangkan tombol 'audio stop
                                    ListBox2.Items.Add(KonversiRuang(poll) + " : " + tanggal.Text)
                                    DbaseAlarmOff(CInt(poll), CStr(tanggal.Text), CDbl(Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10), CStr("Temperature Pass"))
                                    Exit Do
                                End If
                            Next

                            n = 0
                        Loop While False
                    End If
                Else 'selain poll 6 dan 12
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 <= Val(2) Then
                        'tes2 = 0 'visible tombol true 'audio play
                        If address(n) <> poll Then 'masalah kamar
                            p = n + 1
                            Do
                                For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah memiliki kamar
                                    If address(n) = poll Then
                                        n = n
                                        silent(poll) = 1
                                        address(n) = poll
                                        'tes2 = 0 'visible tombol true
                                        Exit Do
                                    End If
                                Next
                                For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah salah
                                    If address(n) = 0 Then
                                        n = n
                                        silent(poll) = 1
                                        address(n) = poll
                                        tes2 = 0 'visible tombol true
                                        ListBox1.Items.Add(KonversiRuang(poll) + " : " + tanggal.Text + " : " + "Temperature too Low")
                                        DbaseAlarmOn(CInt(poll), CStr(tanggal.Text), CDbl(Val(Mid(commqualityTZ.responsePV.Text, 9, 5))), CStr("Temperature too Low"))
                                        Exit Do
                                    End If
                                Next
                            Loop While False
                        Else ' address(n) sesuai dengan poll sekarang
                            address(n) = poll
                            silent(poll) = 1
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(8) Then
                        ' tes2 = 0 'visible tombol true 'audio play
                        If address(n) <> poll Then 'masalah kamar
                            p = n + 1
                            Do
                                For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah memiliki kamar
                                    If address(n) = poll Then
                                        n = n
                                        silent(poll) = 1
                                        address(n) = poll
                                        ' tes2 = 0 'visible tombol true
                                        Exit Do
                                    End If
                                Next
                                For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah salah
                                    If address(n) = 0 Then
                                        n = n
                                        silent(poll) = 1
                                        address(n) = poll
                                        tes2 = 0 'visible tombol true
                                        ListBox1.Items.Add(KonversiRuang(poll) + " : " + tanggal.Text + " : " + "Temperature too high")
                                        DbaseAlarmOn(CInt(poll), CStr(tanggal.Text), CDbl(Val(Mid(commqualityTZ.responsePV.Text, 9, 5))), CStr("Temperature too high"))
                                        Exit Do
                                    End If
                                Next
                            Loop While False
                        Else ' address(n) sesuai dengan poll sekarang
                            address(n) = poll
                            silent(poll) = 1
                        End If
                    Else
                        Do
                            For n = 1 To 39 Step 1 'ngecek apakah poll sekarang sudah pernah memiliki kamar atau sudah pernah salah
                                If address(n) = poll Then
                                    n = n
                                    address(n) = 0
                                    silent(poll) = 0 'mematikan kedip
                                    tes2 = 1 'menghilangkan tombol 'audio stop
                                    DbaseAlarmOff(CInt(poll), CStr(tanggal.Text), CDbl(Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10), CStr("Temperature Pass"))
                                    'DbaseAlarmOff(CInt(poll), CStr(tanggal.Text))
                                    Exit Do
                                End If
                            Next
                            n = 0
                        Loop While False
                    End If
                End If
            End If


            Select Case Mid(commqualityTZ.responsePV.Text, 3, 5)
                Case "15RDP" : BIOCOLD032.Text = Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(2) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(8) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    End If
                    : If BIOCOLD032.Text <> "OFFLINE" Then mid(kp1, 1, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                Case "01RDP" : BIOCOLD037A.Text = Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(2) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(8) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    End If
                    : If BIOCOLD037A.Text <> "OFFLINE" Then mid(kp1, 14, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                Case "02RDP" : BIOCOLD037B.Text = Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(2) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(8) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    End If
                    : If BIOCOLD037B.Text <> "OFFLINE" Then mid(kp1, 15, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                Case "03RDP" : BIOCOLD038A.Text = Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(2) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(8) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    End If
                    : If BIOCOLD038A.Text <> "OFFLINE" Then mid(kp1, 13, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                Case "04RDP" : BIOCOLD038B.Text = Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(2) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(8) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    End If
                    : If BIOCOLD038B.Text <> "OFFLINE" Then mid(kp1, 12, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                Case "05RDP" : BIOCOLD039.Text = Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(2) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(8) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    End If
                    : If BIOCOLD039.Text <> "OFFLINE" Then mid(kp1, 11, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                Case "07RDP" : BIOCOLD040.Text = Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(2) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(8) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    End If
                    : If BIOCOLD040.Text <> "OFFLINE" Then mid(kp1, 9, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                Case "08RDP" : BIOCOLD041A.Text = Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(2) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(8) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    End If
                    : If BIOCOLD041A.Text <> "OFFLINE" Then mid(kp1, 8, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                Case "09RDP" : BIOCOLD041B.Text = Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(2) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(8) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    End If
                    : If BIOCOLD041B.Text <> "OFFLINE" Then mid(kp1, 7, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                Case "10RDP" : BIOCOLD042A.Text = Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(2) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(8) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    End If
                    : If BIOCOLD042A.Text <> "OFFLINE" Then mid(kp1, 6, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                Case "11RDP" : BIOCOLD042B.Text = Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(2) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(8) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    End If
                    : If BIOCOLD042B.Text <> "OFFLINE" Then mid(kp1, 5, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                Case "06RDP" : BIOCOLD043.Text = Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(0) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(-23) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    End If
                    : If BIOCOLD043.Text <> "OFFLINE" Then mid(kp1, 10, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
                Case "12RDP" : BIOCOLD044.Text = Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10
                    If Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 < Val(15) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    ElseIf Val(Mid(commqualityTZ.responsePV.Text, 9, 5)) / 10 > Val(20) Then
                        If tes2 = 1 Then
                            Button16.Visible = False
                            My.Computer.Audio.Stop()
                        ElseIf tes2 = 0 Then
                            My.Computer.Audio.Play(My.Resources.Warning_Sounds_Free_Sound_Effects_Warning_Sound_C, AudioPlayMode.BackgroundLoop)
                            Button16.Visible = True
                        End If
                    End If
                    : If BIOCOLD044.Text <> "OFFLINE" Then mid(kp1, 4, 1) = "1" 'untuk menunjukan bahwa pada alamat ini sudah ada nilai yang terbaca
            End Select
            save8 = jumlahsalah1
            alamatsalah1()

            If jumlahsalah1 = 0 Then
                If jumlahsalah = 0 Then
                    My.Computer.Audio.Stop()
                    Button16.Visible = False
                End If
            End If
            'mendeteksi unit yang offline
            If commqualityTZ.responsePV.Text <> "" Then
                'commqualityTZ.TextBox1.Text = poll
                checktrue(poll)
            Else checkfalse(poll)
            End If

            commqualityTZ.responsePV.Text = ""
            commqualityTZ.Poll.Text = poll
            Waktu.Enabled = True
            TimerPV.Enabled = False
            TimerSV.Enabled = True
        End If
    End Sub
    Function KonversiRuang(ByVal alamat As Integer)
        Dim ruang As String = Nothing
        Select Case alamat
            Case 1 : ruang = "BIO/COLD/032"
            Case 2 : ruang = "BIO/COLD/037A"
            Case 3 : ruang = "BIO/COLD/038A"
            Case 4 : ruang = "BIO/COLD/038B"
            Case 5 : ruang = "BIO/COLD/039"
            Case 6 : ruang = "BIO/COLD/043"
            Case 7 : ruang = "BIO/COLD/040"
            Case 8 : ruang = "BIO/COLD/041A"
            Case 9 : ruang = "BIO/COLD/041A"
            Case 10 : ruang = "BIO/COLD/042A"
            Case 11 : ruang = "BIO/COLD/042B"
            Case 12 : ruang = "BIO/COLD/044"
            Case 100 : ruang = "POS 3"
            Case 101 : ruang = "PPIC"
            Case 102 : ruang = "2"
            Case 103 : ruang = "3"
            Case 104 : ruang = "Emulsi 2"
            Case 105 : ruang = "Emulsi 1"
            Case 106 : ruang = "POS 4"
            Case 107 : ruang = "INO"
            Case 108 : ruang = "Panen 4"
            Case 109 : ruang = "Panen 3"
            Case 110 : ruang = "Koridor"
            Case 111 : ruang = "INO 3"
            Case 112 : ruang = "PRA INO 2"
            Case 113 : ruang = "Passage"
            Case 114 : ruang = "PRA INO 1"
            Case 115 : ruang = "Alat Steril"
            Case 116 : ruang = "STERILISASI"
            Case 117 : ruang = "Persiapan"
            Case 118 : ruang = "INO 2"
            Case 119 : ruang = "INO 1"
            Case 120 : ruang = "POS INO 2"
            Case 121 : ruang = "POS INO 1"
            Case 122 : ruang = "Panen 1"
            Case 123 : ruang = "Panen 2"
            Case 124 : ruang = "Inaktif 1"
            Case 125 : ruang = "Inaktif 2"
        End Select
        Return ruang
    End Function
    Private Sub TimerSV_Tick(sender As Object, e As EventArgs) Handles TimerSV.Tick
        Dim PauseTime, Start As Double
        If MSComm2.PortOpen = True Then
            Command_Head = "S"

            'commqualityTZ.ResponseSV.Text = ""
            If Command_Head = "S" Then
                Select Case poll
                    Case 2, 13, 20, 31 : Bcode = "j"
                    Case 1, 10, 23, 32 : Bcode = "i"
                    Case 11, 22 : Bcode = "h"
                    Case 7, 16, 25 : Bcode = "o"
                    Case 6, 17, 24 : Bcode = "n"
                    Case 5, 14, 27 : Bcode = "m"
                    Case 4, 15, 26 : Bcode = "l"
                    Case 29 : Bcode = "c"
                    Case 28 : Bcode = "b"
                    Case 3, 12, 21, 30 : Bcode = "k"
                    Case 9, 18 : Bcode = "a"
                    Case 8, 19 : Bcode = "`"
                End Select
            End If

            If poll > 9 Then
                commqualityTZ.CommandSV.Text = Chr(2) + Chr(Asc(var2)) + Chr(Asc(var1)) + "RX" + Command_Head + "0" +
                Chr(3) + Bcode
            Else
                commqualityTZ.CommandSV.Text = Chr(2) + "0" + Chr(Asc(var1)) + "RX" + Command_Head + "0" +
                Chr(3) + Bcode
            End If

            MSComm2.Output = commqualityTZ.CommandSV.Text
            Do While MSComm2.OutBufferCount > 0
            Loop
            PauseTime = 1
            Start = Microsoft.VisualBasic.DateAndTime.Timer
            Do While (Microsoft.VisualBasic.DateAndTime.Timer < Start + PauseTime)

            Loop
            commqualityTZ.ResponseSV.Text = commqualityTZ.ResponseSV.Text + MSComm2.Input

            Select Case Mid(commqualityTZ.ResponseSV.Text, 3, 5)
                Case "15RDS" : SV032.Text = Val(Mid(commqualityTZ.ResponseSV.Text, 9, 5)) / 10
                Case "01RDS" : SV037A.Text = Val(Mid(commqualityTZ.ResponseSV.Text, 9, 5)) / 10
                Case "02RDS" : SV037B.Text = Val(Mid(commqualityTZ.ResponseSV.Text, 9, 5)) / 10
                Case "03RDS" : SV038A.Text = Val(Mid(commqualityTZ.ResponseSV.Text, 9, 5)) / 10
                Case "04RDS" : SV038B.Text = Val(Mid(commqualityTZ.ResponseSV.Text, 9, 5)) / 10
                Case "05RDS" : SV039.Text = Val(Mid(commqualityTZ.ResponseSV.Text, 9, 5)) / 10
                Case "07RDS" : SV040.Text = Val(Mid(commqualityTZ.ResponseSV.Text, 9, 5)) / 10
                Case "08RDS" : SV041A.Text = Val(Mid(commqualityTZ.ResponseSV.Text, 9, 5)) / 10
                Case "09RDS" : SV041B.Text = Val(Mid(commqualityTZ.ResponseSV.Text, 9, 5)) / 10
                Case "10RDS" : SV042A.Text = Val(Mid(commqualityTZ.ResponseSV.Text, 9, 5)) / 10
                Case "11RDS" : SV042B.Text = Val(Mid(commqualityTZ.ResponseSV.Text, 9, 5)) / 10
                Case "06RDS" : SV043.Text = Val(Mid(commqualityTZ.ResponseSV.Text, 9, 5)) / 10
                Case "12RDS" : SV044.Text = Val(Mid(commqualityTZ.ResponseSV.Text, 9, 5)) / 10
            End Select
            If poll = 16 Then
                poll = 0
                var1 = 0
                var2 = 0
            End If
        End If
        commqualityTZ.ResponseSV.Text = ""
        commqualityTZ.Poll.Text = poll

        TimerSV.Enabled = False
        TimerPV.Enabled = True
    End Sub

    Private Sub TZDbase_Tick(sender As Object, e As EventArgs) Handles TZDbase.Tick
        'dei_dbase.Enabled = False
        Dim provider, dataFile, connString As String
        Dim myConnection As OleDbConnection = New OleDbConnection
        Dim data1, data2, data3, data4 As String
        If Val(BIOCOLD037A.Text) > Val(BIOCOLD037B.Text) Then data1 = BIOCOLD037A.Text Else data1 = BIOCOLD037B.Text
        If Val(BIOCOLD038A.Text) > Val(BIOCOLD038B.Text) Then data2 = BIOCOLD038A.Text Else data2 = BIOCOLD038B.Text
        If Val(BIOCOLD041A.Text) > Val(BIOCOLD041B.Text) Then data3 = BIOCOLD041A.Text Else data3 = BIOCOLD041B.Text
        If Val(BIOCOLD042A.Text) > Val(BIOCOLD042B.Text) Then data4 = BIOCOLD042A.Text Else data4 = BIOCOLD042B.Text
        provider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
        dataFile = "E:\PENS\BiocoldDB.mdb"
        connString = provider & dataFile
        myConnection.ConnectionString = connString
        myConnection.Open()
        Dim str As String
        'Digunakan untuk memasukan data yang terbaca ke database sesuai dengan field yang ada
        str = "Insert into Biocold ([Tanggal/Jam],[BIO/COLD/032],[BIO/COLD/037],[BIO/COLD/038],[BIO/COLD/039],[BIO/COLD/040],[BIO/COLD/041],[BIO/COLD/042],[BIO/COLD/043],[BIO/COLD/044]) Values (?,?,?,?,?,?,?,?,?,?) "
        Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
        cmd.Parameters.Add(New OleDbParameter("Tanggal/Jam", CDate(now.Text)))
        'record data ke database TZ, apabila ada alamat yang belum ada nilainya maka, akan disimpan sebagai nilai 0 di database. hal ini telah diantisipasi sebelumnya dengan adanya variable kp1 yang memastikan semua alamat yang diinginkan memiliki nilai terlebih dahulu meskipun niilainya 0
        If BIOCOLD032.Text <> "OFFLINE" Then cmd.Parameters.Add(New OleDbParameter("BIO/COLD/032", CDbl(BIOCOLD032.Text))) Else cmd.Parameters.Add(New OleDbParameter("BIO/COLD/032", 0))
        '  If data1 <> "OFFLINE" Then cmd.Parameters.Add(New OleDbParameter("BIO/COLD/032", CDbl(data1))) Else cmd.Parameters.Add(New OleDbParameter("BIO/COLD/032", 0))
        If data1 <> "OFFLINE" Then cmd.Parameters.Add(New OleDbParameter("BIO/COLD/037", CDbl(data1))) Else cmd.Parameters.Add(New OleDbParameter("BIO/COLD/037", 0))
        If data2 <> "OFFLINE" Then cmd.Parameters.Add(New OleDbParameter("BIO/COLD/038", CDbl(data2))) Else cmd.Parameters.Add(New OleDbParameter("BIO/COLD/038", 0))
        If BIOCOLD039.Text <> "OFFLINE" Then cmd.Parameters.Add(New OleDbParameter("BIO/COLD/039", CDbl(BIOCOLD039.Text))) Else cmd.Parameters.Add(New OleDbParameter("BIO/COLD/039", 0))
        If BIOCOLD040.Text <> "OFFLINE" Then cmd.Parameters.Add(New OleDbParameter("BIO/COLD/040", CDbl(BIOCOLD040.Text))) Else cmd.Parameters.Add(New OleDbParameter("BIO/COLD/040", 0))
        If data3 <> "OFFLINE" Then cmd.Parameters.Add(New OleDbParameter("BIO/COLD/041", CDbl(data3))) Else cmd.Parameters.Add(New OleDbParameter("BIO/COLD/041", 0))
        If data4 <> "OFFLINE" Then cmd.Parameters.Add(New OleDbParameter("BIO/COLD/042", CDbl(data4))) Else cmd.Parameters.Add(New OleDbParameter("BIO/COLD/042", 0))
        If BIOCOLD043.Text <> "OFFLINE" Then cmd.Parameters.Add(New OleDbParameter("BIO/COLD/043", CDbl(BIOCOLD043.Text))) Else cmd.Parameters.Add(New OleDbParameter("BIO/COLD/043", 0))
        If BIOCOLD044.Text <> "OFFLINE" Then cmd.Parameters.Add(New OleDbParameter("BIO/COLD/044", CDbl(BIOCOLD044.Text))) Else cmd.Parameters.Add(New OleDbParameter("BIO/COLD/044", 0))

        Label34.Text = DateAndTime.Now()

        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            myConnection.Close()

        Catch ex As Exception

            MsgBox(ex.Message)
        End Try
        '   dei_dbase.Enabled = True
        ' TZDbase.Enabled = False
    End Sub

    ' Private Sub DbaseAlarmOn(ByVal address As Integer, ByVal waktu As String, ByVal nilai As Double, ByVal keterangan As String)
    'dei_dbase.Enabled = False
    'TZDbase.Enabled = False
    'Dim provider, dataFile As String
    'Dim myConnection As OleDbConnection = New OleDbConnection
    '  provider = "Provider=Microsoft.ACE.OleDB.12.0;Data Source="
    '  dataFile = "E:\PENS\alarmlog.mdb"
    '  myConnection.ConnectionString = provider & dataFile
    '  myConnection.Open()
    'Dim str As String
    'Digunakan untuk memasukan data yang terbaca ke database sesuai dengan field yang ada
    '   str = "insert Into Table1 ([Ruangan],[Alarm On],[Nilai],[Keterangan],) Values(?,?,?,?)"
    'Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
    '  cmd.Parameters.Add(New OleDbParameter("Ruangan", CStr(KonversiRuang(address))))
    '  cmd.Parameters.Add(New OleDbParameter("Alarm On", CStr(waktu)))
    '  cmd.Parameters.Add(New OleDbParameter("Nilai", CDbl(nilai)))
    '  cmd.Parameters.Add(New OleDbParameter("Keterangan", CStr(keterangan)))

    'Try
    '  cmd.ExecuteNonQuery()
    '  cmd.Dispose()
    '  myConnection.Close()
    'Catch ex As Exception
    '   MsgBox(ex.Message)
    'End Try
    '  End Sub

    '  Private Sub DbaseAlarmOff(ByVal address As Integer, ByVal waktu As String, ByVal nilai As Double, ByVal keterangan As String)

    'Dim provider, dataFile As String
    '  Dim myConnection As OleDbConnection = New OleDbConnection
    '  provider = "Provider=Microsoft.ACE.OleDB.12.0;Data Source="
    '  dataFile = "E:\PENS\alarmlog.mdb"
    '  myConnection.ConnectionString = provider & dataFile
    '  myConnection.Open()
    'Dim str As String
    'Digunakan untuk memasukan data yang terbaca ke database sesuai dengan field yang ada
    '   str = "insert Into Table1 ([Ruangan],[Alarm Off],[Nilai],[Keterangan],) Values(?,?,?,?)"
    'Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
    '   cmd.Parameters.Add(New OleDbParameter("Ruangan", CStr(KonversiRuang(address))))
    '   cmd.Parameters.Add(New OleDbParameter("Alarm Off", CStr(waktu)))
    '  cmd.Parameters.Add(New OleDbParameter("Nilai", CDbl(nilai)))
    '  cmd.Parameters.Add(New OleDbParameter("Keterangan", CStr(keterangan)))

    'Try
    '   cmd.ExecuteNonQuery()
    '  cmd.Dispose()
    '   myConnection.Close()
    'Catch ex As Exception
    '      MsgBox(ex.Message)
    'End Try
    '  End Sub


    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        'Timer1.Enabled = True
        dei_dbase.Enabled = False
        TZDbase.Enabled = False
        Dim provider, dataFile As String
        Dim myConnection As OleDbConnection = New OleDbConnection
        provider = "Provider=Microsoft.ACE.OleDB.12.0;Data Source="
        dataFile = "E:\PENS\alarmlog.mdb"
        myConnection.ConnectionString = provider & dataFile
        myConnection.Open()
        Dim str As String
        'Digunakan untuk memasukan data yang terbaca ke database sesuai dengan field yang ada
        str = "insert Into table ([Ruangan],[Alarm Off]) Values(?,?)"
        Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
        cmd.Parameters.Add(New OleDbParameter("Ruangan", CInt(TextBox5.Text)))
        cmd.Parameters.Add(New OleDbParameter("Alarm Off", CStr(TextBox3.Text)))
        TextBox6.Text = "Alarm OFF"
        dei_dbase.Enabled = True
        'DEI_timer.Enabled = False
        'DigitalOutput.Enabled = False
        'Dim data As String
        'Dim t1 As Long
        'Dim k, i As Integer


        ''Slave = 100
        'On Error Resume Next
        't1 = (CLng(197) * 256)
        'data = Chr(TextBox3.Text) + Chr(3) + Chr(Val(startaddlabel.Text) \ 256) + Chr((startaddlabel.Text) Mod 256) + Chr(0) + Chr(Val(quantitylabel.Text))
        'CRC_16(data, 6)
        'data = data + Chr(CRC_High) + Chr(CRC_Low)
        'Dim PauseTime, Start, Finish
        'baru.dei_response.Text = ""
        'MSComm1.InputLen = 0

        'baru.dei_request.Text = Val(Asc(Mid(data, 1, 3))) & Val(Asc(Mid(data, 2, 1))) & Val(Asc(Mid(data, 3, 1))) _
        '& (Val(Asc(Mid(data, 4, 1))) + (Val(Asc(Mid(data, 5, 1))))) _
        '& (Val(Asc(Mid(data, 6, 1))) + (Val(Asc(Mid(data, 7, 1))))) _
        '& (Mid(data, 8, 1)) & (Mid(data, 9, 1))
        ''baru.dei_request.Text = Hex(Asc(Mid(data, 1, 3))) & Hex(Asc(Mid(data, 2, 1))) & Hex(Asc(Mid(data, 3, 1))) _
        ''& (Hex(Asc(Mid(data, 4, 1))) + (Hex(Asc(Mid(data, 5, 1))))) _
        ''& (Hex(Asc(Mid(data, 6, 1))) + (Hex(Asc(Mid(data, 7, 1))))) _
        ''& (Mid(data, 8, 1)) & (Mid(data, 9, 1))
        'MSComm1.Output = data

        'Do While MSComm1.OutBufferCount > 0
        'Loop
        'PauseTime = 0.1  ' Set duration.
        'Start = Microsoft.VisualBasic.Timer    ' Set start time.
        'Do While (Microsoft.VisualBasic.Timer < Start + PauseTime) And (MSComm1.InBufferCount < Val(quantitylabel.Text) * 2 + 5)
        '    Application.DoEvents()
        'Loop

        'baru.dei_inBuffer.Text = (MSComm1.InBufferCount)
        ''quantitylabel.text = quantitydei
        ''startaddlabel.text = startadd
        'Response_mod = MSComm1.Input

        'baru.dei_response.Text = Val(Asc(Mid(Response_mod, 1, 1))) _
        '& Val(Asc(Mid(Response_mod, 2, 1))) _
        '& Val(Asc(Mid(Response_mod, 3, 1))) _
        '& Val(Asc(Mid(Response_mod, 4, 1))) * CLng(256) + (Val(Asc(Mid(Response_mod, 5, 1)))) _
        '& Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + (Val(Asc(Mid(Response_mod, 7, 1)))) _
        '& Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + (Val(Asc(Mid(Response_mod, 9, 1)))) _
        '& Val(Asc(Mid(Response_mod, 10, 1))) * CLng(256) + (Val(Asc(Mid(Response_mod, 11, 1)))) _
        '& (Mid(Response_mod, 11, 1)) & (Mid(Response_mod, 12, 1))

        'TextBox5.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
        'TextBox4.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
        'DEI_timer.Enabled = True
    End Sub

    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click

    End Sub

    Private Sub BIOCOLD039_Click(sender As Object, e As EventArgs) Handles BIOCOLD039.Click

    End Sub

    Private Sub SV039_Click(sender As Object, e As EventArgs) Handles SV039.Click

    End Sub

    Private Sub pa118_Click(sender As Object, e As EventArgs) Handles pa118.Click

    End Sub

    Private Sub Pa_121_Click(sender As Object, e As EventArgs) Handles Pa_121.Click

    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Label77.Visible = False
        Label88.Visible = False
        Label89.Visible = False
        Label86.Visible = False
        Label85.Visible = False
        TextBox2.Visible = False
        setval.Visible = False
        ComboBox2.Visible = False
        waktuon.Visible = False
        waktuoff.Visible = False
        Button15.Visible = False
        Button17.Visible = False
        Button18.Visible = False
        Button22.Visible = False
    End Sub

    Private Sub hapusdata_tz_Tick(sender As Object, e As EventArgs)
        Dim provider, dataFile As String
        Dim myConnection As OleDbConnection = New OleDbConnection
        provider = "Provider=Microsoft.ACE.OleDB.12.0;Data Source="
        dataFile = "E:\PENS\BiocoldDB.mdb"
        myConnection.ConnectionString = provider & dataFile
        myConnection.Open()
        Dim str As String

        str = "DELETE FROM Biocold"
        Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
        cmd.ExecuteNonQuery()
        myConnection.Close()
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
    End Sub

    Private Sub AlarmLogDB_Tick(sender As Object, e As EventArgs) Handles AlarmLogDBon.Tick
        dei_dbase.Enabled = False
        TZDbase.Enabled = False
        Dim provider, dataFile As String
        Dim myConnection As OleDbConnection = New OleDbConnection
        provider = "Provider=Microsoft.ACE.OleDB.12.0;Data Source="
        dataFile = "E:\PENS\alarmlog.mdb"
        myConnection.ConnectionString = provider & dataFile
        myConnection.Open()
        Dim str As String
        'Digunakan untuk memasukan data yang terbaca ke database sesuai dengan field yang ada
        str = "insert Into table ([Ruangan],[Alarm On]) Values(?,?)"
        Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
        cmd.Parameters.Add(New OleDbParameter("Ruangan", CInt(address(n))))
        cmd.Parameters.Add(New OleDbParameter("Alarm On", CStr(tanggal.Text)))
        TextBox3.Text = "AlarmON"
        dei_dbase.Enabled = True 'f
        AlarmLogDBon.Enabled = False
    End Sub

    Private Sub AlarmLogDBoff_Tick(sender As Object, e As EventArgs) Handles AlarmLogDBoff.Tick
        dei_dbase.Enabled = False
        TZDbase.Enabled = False
        Dim provider, dataFile As String
        Dim myConnection As OleDbConnection = New OleDbConnection
        provider = "Provider=Microsoft.ACE.OleDB.12.0;Data Source="
        dataFile = "E:\PENS\alarmlog.mdb"
        myConnection.ConnectionString = provider & dataFile
        myConnection.Open()
        Dim str As String
        'Digunakan untuk memasukan data yang terbaca ke database sesuai dengan field yang ada
        str = "insert Into table ([Ruangan],[Alarm Off]) Values(?,?)"
        Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
        cmd.Parameters.Add(New OleDbParameter("Ruangan", CInt(address(n))))
        cmd.Parameters.Add(New OleDbParameter("Alarm Off", CStr(tanggal.Text)))
        TextBox5.Text = "Alarm OFF"
        dei_dbase.Enabled = True
        AlarmLogDBoff.Enabled = False
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        password.Label2.Text = 3
        password.Show()
    End Sub



    'Private Sub DEISV(alamat As Integer, setvalue As Double)
    '    Dim data As String
    '    Dim t1 As Long
    '    DEI_timer.Enabled = False
    '    DigitalOutput.Enabled = False
    '    'Select Case ComboBox1.Text
    '    '    Case "SPIH" : opo = 340
    '    '    Case "SPIC" : opo = 338
    '    '    Case "L-RANGE" : opo = 336
    '    '    Case "R-RANGE" : opo = 335
    '    'End Select
    '    ' DEI_timer.Enabled = False
    '    On Error Resume Next
    '    t1 = (CLng(197) * 256)
    '    data = Chr(Val(alamat)) + Chr(6) + Chr(Val(338) \ 256) + Chr((338) Mod 256) + Chr(0) + Chr(CDbl(Val(setvalue)) * 10)
    '    CRC_16(data, 6)
    '    data = data + Chr(CRC_High) + Chr(CRC_Low)
    '    Dim PauseTime, Start, Finish, TotalTime

    '    MSComm2.InputLen = 0

    '    MSComm1.Output = data

    '    Do While MSComm1.OutBufferCount > 0
    '    Loop
    '    PauseTime = 0.25    ' Set duration.
    '    Start = Microsoft.VisualBasic.Timer     ' Set start time.
    '    Do While (Microsoft.VisualBasic.Timer < Start + PauseTime) And (MSComm1.InBufferCount < Val(quantitylabel.Text) * 2 + 5)
    '        Application.DoEvents()

    '    Loop

    '    'InBuffer.Text = (MSComm2.InBufferCount)

    '    Response_mod = MSComm1.Input

    '    TextBox6.Text = Hex(Asc(Mid(Response_mod, 1, 1))) _
    '     & " " _
    '     & Hex(Asc(Mid(Response_mod, 2, 1))) _
    '     & " " _
    '     & Hex(Asc(Mid(Response_mod, 3, 1))) _
    '     & " " _
    '     & Hex(Asc(Mid(Response_mod, 4, 1))) _
    '     & " " _
    '     & Hex(Asc(Mid(Response_mod, 5, 1))) _
    '     & " " _
    '     & Hex(Asc(Mid(Response_mod, 6, 1))) _
    '     & " " _
    '     & Hex(Asc(Mid(Response_mod, 7, 1))) _
    '     & " " _
    '     & Hex(Asc(Mid(Response_mod, 8, 1))) _
    '     & " " _
    '     & Hex(Asc(Mid(Response_mod, 9, 1))) _
    '     & " " _
    '    & Hex(Asc(Mid(Response_mod, 10, 1))) _
    '     & " " _
    '    & Hex(Asc(Mid(Response_mod, 11, 1))) _
    '     & " " _
    '    & Hex(Asc(Mid(Response_mod, 12, 1))) _
    '     & " " _
    '    & Hex(Asc(Mid(Response_mod, 13, 1)))
    '    Finish = Microsoft.VisualBasic.Timer    ' Set end time.
    '    DEI_timer.Enabled = True

    'End Sub

    Private Sub tutupform_Click(sender As Object, e As EventArgs) Handles tutupform.Click
        Me.Close()
    End Sub

    Private Sub minimize_Click(sender As Object, e As EventArgs) Handles minimize.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Button15_Click_1(sender As Object, e As EventArgs) Handles Button15.Click
        Dim data As String
        Dim t1 As Long
        Dim k, i As Integer
        TextBox2.Text = ""
        DEI_timer.Enabled = False
        DigitalOutput.Enabled = False
        Button18.Visible = False
        Button17.Visible = False
        Select Case ComboBox2.Text

            Case "POS 3" : opo = 100
            Case "PPIC" : opo = 101
            Case "2" : opo = 102
            Case "3" : opo = 103
            Case "Emulsi 2" : opo = 104
            Case "Emulsi 1" : opo = 105
            Case "POS 4" : opo = 106
            Case "INO" : opo = 107
            Case "Panen 4" : opo = 108
            Case "Panen 3" : opo = 109
            Case "Koridor" : opo = 110
            Case "INO 3" : opo = 111
            Case "PRA INO 2" : opo = 112
            Case "Passage" : opo = 113
            Case "PRA INO 1" : opo = 114
            Case "Alat Steril" : opo = 115
            Case "STERILISASI" : opo = 116
            Case "Persiapan" : opo = 117
            Case "INO 2" : opo = 118
            Case "INO 1" : opo = 119
            Case "POS INO 2" : opo = 120
            Case "POS INO 1" : opo = 121
            Case "Panen 1" : opo = 122
            Case "Panen 2" : opo = 123
            Case "Inaktif 1" : opo = 124
            Case "Inaktif 2" : opo = 125
        End Select

        On Error Resume Next
        t1 = (CLng(197) * 256)
        data = Chr(opo) + Chr(3) + Chr(Val(338) \ 256) + Chr(Val(338) Mod 256) + Chr(0) + Chr(Val(1))
        CRC_16(data, 6)
        data = data + Chr(CRC_High) + Chr(CRC_Low)
        Dim PauseTime, Start, Finish
        MSComm1.InputLen = 0

        MSComm1.Output = data

        Do While MSComm1.OutBufferCount > 0
        Loop
        PauseTime = 0.1  ' Set duration.
        Start = Microsoft.VisualBasic.Timer    ' Set start time.
        Do While (Microsoft.VisualBasic.Timer < Start + PauseTime) And (MSComm1.InBufferCount < Val(quantitylabel.Text) * 2 + 5)
            Application.DoEvents()
        Loop

        baru.dei_inBuffer.Text = (MSComm1.InBufferCount)

        Response_mod = MSComm1.Input

        If Val(Asc(Mid(Response_mod, 1, 1))) = opo Then
            If Val(Asc(Mid(Response_mod, 2, 1))) = 3 Then
                If Len(Val(Asc(Mid(Response_mod, 5, 1))) / Val(10)) = 1 Then
                    TextBox2.Text = "0," + Val(Asc(Mid(Response_mod, 5, 1))) / Val(10)
                Else TextBox2.Text = Val(Asc(Mid(Response_mod, 5, 1))) / Val(10)
                End If
            End If
        End If
        ListBox3.Items.Add(tanggal.Text + " : " + Label83.Text + " :" + " Cek SV " + ": " + ComboBox2.Text + " : " + TextBox2.Text)
        DEI_timer.Enabled = True
        Button17.Visible = True
        Button18.Visible = True
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        If waktuon.Text <> "" Then
            mid(waktuon.Text, 3, 1) = ":"
            waktuon.Text = Mid(waktuon.Text, 1, 5)
        End If
        If waktuoff.Text <> "" Then
            mid(waktuoff.Text, 3, 1) = ":"
            waktuoff.Text = Mid(waktuoff.Text, 1, 5)
        End If
        ListBox3.Items.Add(tanggal.Text + " : " + Label83.Text + " :" + " Set DO " + ": " + ComboBox2.Text + " : " + waktuon.Text + " : " + waktuoff.Text)
        bacaON()
        bacaOFF()
        tulisON()
        tulisOFF()
        flag = 1
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        DEI_timer.Enabled = False
        DigitalOutput.Enabled = False

        Dim data As String
        Dim t1 As Long
        Dim setvalue As Double
        Dim k, i As Integer
        'If opo1 = 6 Then
        '    DEI_timer.Enabled = True
        '    Timer1.Enabled = False
        '    opo1 = 0
        'End If
        'Select Case ComboBox1.Text
        '    Case "SPIH" : opo = 340
        '    Case "SPIC" : opo = 338
        '    Case "L-RANGE" : opo = 336
        '    Case "R-RANGE" : opo = 335
        'End Select
        'CDbl(Val(0))
        ' DEI_timer.Enabled = False
        On Error Resume Next
        t1 = (CLng(197) * 256)
        data = Chr(105) + Chr(3) + Chr(Val(inc) \ 256) + Chr((inc) Mod 256) + Chr(0) + Chr(Val(1))
        CRC_16(data, 6)
        data = data + Chr(CRC_High) + Chr(CRC_Low)
        Dim PauseTime, Start, Finish, TotalTime

        MSComm2.InputLen = 0

        MSComm1.Output = data
        TextBox6.Text = Mid(data, 1, 1)
        Do While MSComm1.OutBufferCount > 0
        Loop
        PauseTime = 0.25    ' Set duration.
        Start = Microsoft.VisualBasic.Timer     ' Set start time.
        Do While (Microsoft.VisualBasic.Timer < Start + PauseTime) And (MSComm1.InBufferCount < Val(quantitylabel.Text) * 2 + 5)
            Application.DoEvents()

        Loop
        Response_mod = MSComm1.Input
        'If Response_mod <> "" Then TextBox6.Text = "cek"
        TextBox7.Text = Val(Asc(Mid(Response_mod, 1, 1)))
        Finish = Microsoft.VisualBasic.Timer    ' Set end time.
        'inc = inc + 1
        TextBox5.Text = (Val(Asc(Mid(Response_mod, 6, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 7, 1)))) / 10
        TextBox4.Text = (Val(Asc(Mid(Response_mod, 8, 1))) * CLng(256) + Val(Asc(Mid(Response_mod, 9, 1)))) / 10
        'opo1 = opo1 + 1
        If Val(Asc(Mid(Response_mod, 1, 1))) = 69 Then
            'DEI_timer.Enabled = True
            Timer1.Enabled = False
        End If
    End Sub

    Private Sub tutup_tz_Click(sender As Object, e As EventArgs) Handles tutup_tz.Click
        Me.Close()
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim data As String
        Dim t1 As Long
        Dim k, i As Integer

        DEI_timer.Enabled = False
        DigitalOutput.Enabled = False
        Button15.Visible = False
        Button18.Visible = False

        If Val(Len(setval.Text)) > 2 Then
            mid(setval.Text, 3, 1) = ","
        End If
        setval.Text = CDbl(Mid(setval.Text, 1, 4))
        Select Case ComboBox2.Text
            Case "POS 3" : opo = 100
            Case "PPIC" : opo = 101
            Case "2" : opo = 102
            Case "3" : opo = 103
            Case "Emulsi 2" : opo = 104
            Case "Emulsi 1" : opo = 105
            Case "POS 4" : opo = 106
            Case "INO" : opo = 107
            Case "Panen 4" : opo = 108
            Case "Panen 3" : opo = 109
            Case "Koridor" : opo = 110
            Case "INO 3" : opo = 111
            Case "PRA INO 2" : opo = 112
            Case "Passage" : opo = 113
            Case "PRA INO 1" : opo = 114
            Case "Alat Steril" : opo = 115
            Case "STERILISASI" : opo = 116
            Case "Persiapan" : opo = 117
            Case "INO 2" : opo = 118
            Case "INO 1" : opo = 119
            Case "POS INO 2" : opo = 120
            Case "POS INO 1" : opo = 121
            Case "Panen 1" : opo = 122
            Case "Panen 2" : opo = 123
            Case "Inaktif 1" : opo = 124
            Case "Inaktif 2" : opo = 125
        End Select
        'Select Case ComboBox1.Text
        '    Case "SPIH" : opo = 340
        '    Case "SPIC" : opo = 338
        '    Case "L-RANGE" : opo = 336
        '    Case "R-RANGE" : opo = 335
        'End Select
        'CDbl(Val(0))
        ' DEI_timer.Enabled = False
        On Error Resume Next
        t1 = (CLng(197) * 256)
        data = Chr(Val(opo)) + Chr(6) + Chr(Val(338) \ 256) + Chr((338) Mod 256) + Chr(0) + Chr(CDbl(setval.Text) * Val(10))
        CRC_16(data, 6)
        data = data + Chr(CRC_High) + Chr(CRC_Low)
        Dim PauseTime, Start, Finish, TotalTime

        MSComm2.InputLen = 0

        MSComm1.Output = data

        Do While MSComm1.OutBufferCount > 0
        Loop
        PauseTime = 0.25    ' Set duration.
        Start = Microsoft.VisualBasic.Timer     ' Set start time.
        Do While (Microsoft.VisualBasic.Timer < Start + PauseTime) And (MSComm1.InBufferCount < Val(quantitylabel.Text) * 2 + 5)
            Application.DoEvents()

        Loop
        Response_mod = MSComm1.Input
        Finish = Microsoft.VisualBasic.Timer    ' Set end time.
        ListBox3.Items.Add(tanggal.Text + " : " + Label83.Text + " :" + " Set Value " + ": " + Str(KonversiRuang(opo)) + " : " + setval.Text)
        DEI_timer.Enabled = True
        Button15.Visible = True
        Button18.Visible = True
    End Sub

    Private Sub minimize_tz_Click(sender As Object, e As EventArgs) Handles minimize_tz.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub DEISV_Tick(sender As Object, e As EventArgs) Handles DEISV.Tick
        Dim data As String
        Dim t1 As Long
        Dim setvalue As Double
        Dim k, i, i1, lenlen, slavez As Integer
        DEI_timer.Enabled = False
        DigitalOutput.Enabled = False
        Button15.Visible = False
        Button17.Visible = False
        i1 = 0
        For i = 1 To 30 Step 1
            If digout(i) <> "" Then i1 = i1 + 1 ' maksimal / total address
        Next
        'lenlen = Val(Len(digout(i1))) - Val(5) 'kl itu iterasi dari timer
        slavez = Val(digout(KL)) + Val(99)

        If KL = 26 Then
            DEI_timer.Enabled = True
            Button15.Visible = True
            For k = 0 To 30 Step 1
                digout(k) = ""
            Next
            KL = 0
            slavez = 100
            DEISV.Enabled = False
        End If

        If rx = 1 Then
            DOn = 1
        End If
        If rx = 2 Then
            DOn = 0
        End If
        'Select Case ComboBox1.Text
        '    Case "SPIH" : opo = 340
        '    Case "SPIC" : opo = 338
        '    Case "L-RANGE" : opo = 336
        '    Case "R-RANGE" : opo = 335
        'End Select
        ' DEI_timer.Enabled = False
        'starting address
        'DO1 220
        'DO2 233
        'DO3 246
        'DO4 259
        'DO5 272
        'DO6 285
        'DO7 298
        'DO8 311
        On Error Resume Next
        t1 = (CLng(197) * 256)
        data = Chr(Val(slavez)) + Chr(6) + Chr(Val(233) \ 256) + Chr((233) Mod 256) + Chr(0) + Chr(DOn)
        CRC_16(data, 6)
        data = data + Chr(CRC_High) + Chr(CRC_Low)
        Dim PauseTime, Start, Finish, TotalTime

        MSComm2.InputLen = 0

        MSComm1.Output = data

        Do While MSComm1.OutBufferCount > 0
        Loop
        PauseTime = 0.25    ' Set duration.
        Start = Microsoft.VisualBasic.Timer     ' Set start time.
        Do While (Microsoft.VisualBasic.Timer < Start + PauseTime) And (MSComm1.InBufferCount < Val(quantitylabel.Text) * 2 + 5)
            Application.DoEvents()

        Loop

        'InBuffer.Text = (MSComm2.InBufferCount)

        Response_mod = MSComm1.Input

        'TextBox6.Text = Hex(Asc(Mid(Response_mod, 1, 1))) _
        ' & " " _
        ' & Hex(Asc(Mid(Response_mod, 2, 1))) _
        ' & " " _
        ' & Hex(Asc(Mid(Response_mod, 3, 1))) _
        ' & " " _
        ' & Hex(Asc(Mid(Response_mod, 4, 1))) _
        ' & " " _
        ' & Hex(Asc(Mid(Response_mod, 5, 1))) _
        ' & " " _
        ' & Hex(Asc(Mid(Response_mod, 6, 1))) _
        ' & " " _
        ' & Hex(Asc(Mid(Response_mod, 7, 1))) _
        ' & " " _
        ' & Hex(Asc(Mid(Response_mod, 8, 1))) _
        ' & " " _
        ' & Hex(Asc(Mid(Response_mod, 9, 1))) _
        ' & " " _
        '& Hex(Asc(Mid(Response_mod, 10, 1))) _
        ' & " " _
        '& Hex(Asc(Mid(Response_mod, 11, 1))) _
        ' & " " _
        '& Hex(Asc(Mid(Response_mod, 12, 1))) _
        ' & " " _
        '& Hex(Asc(Mid(Response_mod, 13, 1)))
        Finish = Microsoft.VisualBasic.Timer    ' Set end time.

        ' slaveDEISV = slaveDEISV + 1
        ' inc1 = inc1 + 1
        KL = KL + 1
        DEI_timer.Enabled = True
        Button15.Visible = True
        Button17.Visible = True
    End Sub

    Private Sub SP1Cdatabase()
        If Mid(TextBox1.Text, 3, 1) = "." Then
            Dim provider As String
            Dim myConnection As OleDbConnection = New OleDbConnection
            provider = "Provider=Microsoft.ACE.OleDB.12.0;Data Source= E:\PENS\SPcooling.mdb"
            'dataFile = "E:\PENS\SPcooling.mdb"
            myConnection.ConnectionString = provider
            myConnection.Open()
            Dim str As String
            'Digunakan untuk memasukan data yang terbaca ke database sesuai dengan field yang ada
            str = "UPDATE SP1C SET WAKTUON =@TextBox1 " '[OFF],[100],[101],[102],[103],[Suhu Emulsi 1],[Suhu POS 4],[Suhu INO],[Suhu Panen 4],[Suhu Panen 3],[Suhu Koridor],[Suhu INO 3],[Suhu Pra INO 2],[Suhu Passage],[Suhu Pra INO 1],[Suhu Alat Steril],[Suhu STERILISASI],[Suhu Persiapan],[Suhu INO 2],[Suhu INO 1],[Suhu POS INO 2],[Suhu POS INO 1],[Suhu Panen 1],[Suhu Panen 2],[Suhu Inaktif 1],[Suhu Inaktif2],[Tekanan POS 3],[Tekanan PPIC],[press102],[press103],[Tekanan Emulsi 2],[Tekanan Emulsi 1],[Tekanan POS 4],[Tekanan INO],[Tekanan Panen 4],[Tekanan Panen 3],[Tekanan Koridor],[Tekanan INO 3],[Tekanan Pra INO 2],[Tekanan Passage],[Tekanan Pra INO 1],[Tekanan Alat Steril],[Tekanan STERILISASI],[Tekanan Persiapan],[Tekanan INO 2],[Tekanan INO 1],[Tekanan POS INO 2],[Tekanan POS INO 1],[Tekanan Panen 1],[Tekanan Panen 2],[Tekanan Inaktif 1],[Tekanan Inaktif 2]) Values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
            ' str = "insert into SP1C([ON]) Values(?)"
            Dim cmd As OleDbCommand = New OleDbCommand(str, myConnection)
            'Dim cmd1 As OleDbCommand = New OleDbCommand(str1, myConnection)
            cmd.Parameters.Add(New OleDbParameter("ON", TextBox1.Text))

            Try
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                myConnection.Close()
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else TextBox1.Text = "format salah"
        End If
    End Sub

    Private Sub bacaON()
        Dim Conn As OleDbConnection = New OleDbConnection
        Dim Provider = "Provider=Microsoft.ACE.OleDB.12.0;Data Source= E:\PENS\SPcooling.mdb"
        Conn.ConnectionString = Provider
        Conn.Open()
        Dim cmd As OleDbCommand = New OleDbCommand("SELECT [100], [101], [102],[103],[104],[105],[106],[107],[108],[109],[110],[111],[112],[113],[114],[115],[116],[117],[118],[119],[120],[121],[122],[123],[124],[125] FROM [SP1Con]", Conn)
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        While dr.Read
            var10(1) = dr("100")
            var10(2) = dr("101")
            var10(3) = dr("102")
            var10(4) = dr("103")
            var10(5) = dr("104")
            var10(6) = dr("105")
            var10(7) = dr("106")
            var10(8) = dr("107")
            var10(9) = dr("108")
            var10(10) = dr("109")
            var10(11) = dr("110")
            var10(12) = dr("111")
            var10(13) = dr("112")
            var10(14) = dr("113")
            var10(15) = dr("114")
            var10(16) = dr("115")
            var10(17) = dr("116")
            var10(18) = dr("117")
            var10(19) = dr("118")
            var10(20) = dr("119")
            var10(21) = dr("120")
            var10(22) = dr("121")
            var10(23) = dr("122")
            var10(24) = dr("123")
            var10(25) = dr("124")
            var10(26) = dr("125")
        End While
        Try
            ' cmd.ExecuteNonQuery()
            cmd.Dispose()
            Conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub tulisON()
        Dim Conn As OleDbConnection = New OleDbConnection
        Dim Provider = "Provider=Microsoft.ACE.OleDB.12.0;Data Source= E:\PENS\SPcooling.mdb"
        Dim str As String
        Conn.ConnectionString = Provider
        Conn.Open()
        'If waktuon.Text <> "" Then
        '    start1 = waktuon.Text
        'End If

        If Len(waktuon.Text) <> 0 Then
            Select Case ComboBox2.Text
                Case "SEMUA RUANGAN" : For i = 1 To 30 Step 1
                        var10(i) = waktuon.Text
                    Next
                Case "POS 3" : var10(1) = waktuon.Text
                Case "PPIC" : var10(2) = waktuon.Text
                Case "2" : var10(3) = waktuon.Text
                Case "3" : var10(4) = waktuon.Text
                Case "Emulsi 2" : var10(5) = waktuon.Text
                Case "Emulsi 1" : var10(6) = waktuon.Text
                Case "POS 4" : var10(7) = waktuon.Text
                Case "INO" : var10(8) = waktuon.Text
                Case "Panen 4" : var10(9) = waktuon.Text
                Case "Panen 3" : var10(10) = waktuon.Text
                Case "Koridor" : var10(11) = waktuon.Text
                Case "INO 3" : var10(12) = waktuon.Text
                Case "PRA INO 2" : var10(13) = waktuon.Text
                Case "Passage" : var10(14) = waktuon.Text
                Case "PRA INO 1" : var10(15) = waktuon.Text
                Case "Alat Steril" : var10(16) = waktuon.Text
                Case "STERILISASI" : var10(17) = waktuon.Text
                Case "Persiapan" : var10(18) = waktuon.Text
                Case "INO 2" : var10(19) = waktuon.Text
                Case "INO 1" : var10(20) = waktuon.Text
                Case "POS INO 2" : var10(21) = waktuon.Text
                Case "POS INO 1" : var10(22) = waktuon.Text
                Case "Panen 1" : var10(23) = waktuon.Text
                Case "Panen 2" : var10(24) = waktuon.Text
                Case "Inaktif 1" : var10(25) = waktuon.Text
                Case "Inaktif 2" : var10(26) = waktuon.Text
            End Select
        End If

        str = "UPDATE SP1Con SET 100 =@unik0,101 =@unik1,102 =@unik2,103 =@unik3,104 =@unik4,105 =@unik5,106 =@unik6,107 =@unik7,108 =@unik8,109 =@unik9,110 =@unik10,111 =@unik11,112 =@unik12,113 =@unik13,114 =@unik14,115 =@unik15,116 =@unik16,117 =@unik17,118 =@unik18,119 =@unik19,120 =@unik20,121 =@unik21,122 =@unik22,123 =@unik23,124 =@unik24,125 =@unik25"
        Dim cmd1 As OleDbCommand = New OleDbCommand(str, Conn)
        cmd1.Parameters.Add(New OleDbParameter("100", CStr(var10(1))))
        cmd1.Parameters.Add(New OleDbParameter("101", CStr(var10(2))))
        cmd1.Parameters.Add(New OleDbParameter("102", CStr(var10(3))))
        cmd1.Parameters.Add(New OleDbParameter("103", CStr(var10(4))))
        cmd1.Parameters.Add(New OleDbParameter("104", CStr(var10(5))))
        cmd1.Parameters.Add(New OleDbParameter("105", CStr(var10(6))))
        cmd1.Parameters.Add(New OleDbParameter("106", CStr(var10(7))))
        cmd1.Parameters.Add(New OleDbParameter("107", CStr(var10(8))))
        cmd1.Parameters.Add(New OleDbParameter("108", CStr(var10(9))))
        cmd1.Parameters.Add(New OleDbParameter("109", CStr(var10(10))))
        cmd1.Parameters.Add(New OleDbParameter("110", CStr(var10(11))))
        cmd1.Parameters.Add(New OleDbParameter("111", CStr(var10(12))))
        cmd1.Parameters.Add(New OleDbParameter("112", CStr(var10(13))))
        cmd1.Parameters.Add(New OleDbParameter("113", CStr(var10(14))))
        cmd1.Parameters.Add(New OleDbParameter("114", CStr(var10(15))))
        cmd1.Parameters.Add(New OleDbParameter("115", CStr(var10(16))))
        cmd1.Parameters.Add(New OleDbParameter("116", CStr(var10(17))))
        cmd1.Parameters.Add(New OleDbParameter("117", CStr(var10(18))))
        cmd1.Parameters.Add(New OleDbParameter("118", CStr(var10(19))))
        cmd1.Parameters.Add(New OleDbParameter("119", CStr(var10(20))))
        cmd1.Parameters.Add(New OleDbParameter("120", CStr(var10(21))))
        cmd1.Parameters.Add(New OleDbParameter("121", CStr(var10(22))))
        cmd1.Parameters.Add(New OleDbParameter("122", CStr(var10(23))))
        cmd1.Parameters.Add(New OleDbParameter("123", CStr(var10(24))))
        cmd1.Parameters.Add(New OleDbParameter("124", CStr(var10(25))))
        cmd1.Parameters.Add(New OleDbParameter("125", CStr(var10(26))))
        Try
            cmd1.ExecuteNonQuery()
            cmd1.Dispose()
            Conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub bacaOFF()
        Dim Conn As OleDbConnection = New OleDbConnection
        Dim Provider = "Provider=Microsoft.ACE.OleDB.12.0;Data Source= E:\PENS\SPcooling.mdb"
        Conn.ConnectionString = Provider
        Conn.Open()
        Dim cmd As OleDbCommand = New OleDbCommand("SELECT [100], [101], [102],[103],[104],[105],[106],[107],[108],[109],[110],[111],[112],[113],[114],[115],[116],[117],[118],[119],[120],[121],[122],[123],[124],[125] FROM [SP1Coff]", Conn)
        Dim dr As OleDbDataReader = cmd.ExecuteReader
        While dr.Read
            var11(1) = dr("100")
            var11(2) = dr("101")
            var11(3) = dr("102")
            var11(4) = dr("103")
            var11(5) = dr("104")
            var11(6) = dr("105")
            var11(7) = dr("106")
            var11(8) = dr("107")
            var11(9) = dr("108")
            var11(10) = dr("109")
            var11(11) = dr("110")
            var11(12) = dr("111")
            var11(13) = dr("112")
            var11(14) = dr("113")
            var11(15) = dr("114")
            var11(16) = dr("115")
            var11(17) = dr("116")
            var11(18) = dr("117")
            var11(19) = dr("118")
            var11(20) = dr("119")
            var11(21) = dr("120")
            var11(22) = dr("121")
            var11(23) = dr("122")
            var11(24) = dr("123")
            var11(25) = dr("124")
            var11(26) = dr("125")
        End While
        Try
            ' cmd.ExecuteNonQuery()
            cmd.Dispose()
            Conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub tulisOFF()
        Dim Conn As OleDbConnection = New OleDbConnection
        Dim Provider = "Provider=Microsoft.ACE.OleDB.12.0;Data Source= E:\PENS\SPcooling.mdb"
        Dim str As String
        Conn.ConnectionString = Provider
        Conn.Open()

        'If waktuoff.Text <> "" Then
        '    finish1 = waktuoff.Text
        'End If
        If Len(waktuoff.Text) <> 0 Then
            Select Case ComboBox2.Text
                Case "SEMUA RUANGAN" : For i = 1 To 30 Step 1
                        var11(i) = waktuoff.Text
                    Next
                Case "POS 3" : var11(1) = waktuoff.Text
                Case "PPIC" : var11(2) = waktuoff.Text
                Case "2" : var11(3) = waktuoff.Text
                Case "3" : var11(4) = waktuoff.Text
                Case "Emulsi 2" : var11(5) = waktuoff.Text
                Case "Emulsi 1" : var11(6) = waktuoff.Text
                Case "POS 4" : var11(7) = waktuoff.Text
                Case "INO" : var11(8) = waktuoff.Text
                Case "Panen 4" : var11(9) = waktuoff.Text
                Case "Panen 3" : var11(10) = waktuoff.Text
                Case "Koridor" : var11(11) = waktuoff.Text
                Case "INO 3" : var11(12) = waktuoff.Text
                Case "PRA INO 2" : var11(13) = waktuoff.Text
                Case "Passage" : var11(14) = waktuoff.Text
                Case "PRA INO 1" : var11(15) = waktuoff.Text
                Case "Alat Steril" : var11(16) = waktuoff.Text
                Case "STERILISASI" : var11(17) = waktuoff.Text
                Case "Persiapan" : var11(18) = waktuoff.Text
                Case "INO 2" : var11(19) = waktuoff.Text
                Case "INO 1" : var11(20) = waktuoff.Text
                Case "POS INO 2" : var11(21) = waktuoff.Text
                Case "POS INO 1" : var11(22) = waktuoff.Text
                Case "Panen 1" : var11(23) = waktuoff.Text
                Case "Panen 2" : var11(24) = waktuoff.Text
                Case "Inaktif 1" : var11(25) = waktuoff.Text
                Case "Inaktif 2" : var11(26) = waktuoff.Text
            End Select
        End If

        str = "UPDATE SP1Coff SET 100 =@unik0,101 =@unik1,102 =@unik2,103 =@unik3,104 =@unik4,105 =@unik5,106 =@unik6,107 =@unik7,108 =@unik8,109 =@unik9,110 =@unik10,111 =@unik11,112 =@unik12,113 =@unik13,114 =@unik14,115 =@unik15,116 =@unik16,117 =@unik17,118 =@unik18,119 =@unik19,120 =@unik20,121 =@unik21,122 =@unik22,123 =@unik23,124 =@unik24,125 =@unik25"
        Dim cmd1 As OleDbCommand = New OleDbCommand(str, Conn)
        cmd1.Parameters.Add(New OleDbParameter("100", CStr(var11(1))))
        cmd1.Parameters.Add(New OleDbParameter("101", CStr(var11(2))))
        cmd1.Parameters.Add(New OleDbParameter("102", CStr(var11(3))))
        cmd1.Parameters.Add(New OleDbParameter("103", CStr(var11(4))))
        cmd1.Parameters.Add(New OleDbParameter("104", CStr(var11(5))))
        cmd1.Parameters.Add(New OleDbParameter("105", CStr(var11(6))))
        cmd1.Parameters.Add(New OleDbParameter("106", CStr(var11(7))))
        cmd1.Parameters.Add(New OleDbParameter("107", CStr(var11(8))))
        cmd1.Parameters.Add(New OleDbParameter("108", CStr(var11(9))))
        cmd1.Parameters.Add(New OleDbParameter("109", CStr(var11(10))))
        cmd1.Parameters.Add(New OleDbParameter("110", CStr(var11(11))))
        cmd1.Parameters.Add(New OleDbParameter("111", CStr(var11(12))))
        cmd1.Parameters.Add(New OleDbParameter("112", CStr(var11(13))))
        cmd1.Parameters.Add(New OleDbParameter("113", CStr(var11(14))))
        cmd1.Parameters.Add(New OleDbParameter("114", CStr(var11(15))))
        cmd1.Parameters.Add(New OleDbParameter("115", CStr(var11(16))))
        cmd1.Parameters.Add(New OleDbParameter("116", CStr(var11(17))))
        cmd1.Parameters.Add(New OleDbParameter("117", CStr(var11(18))))
        cmd1.Parameters.Add(New OleDbParameter("118", CStr(var11(19))))
        cmd1.Parameters.Add(New OleDbParameter("119", CStr(var11(20))))
        cmd1.Parameters.Add(New OleDbParameter("120", CStr(var11(21))))
        cmd1.Parameters.Add(New OleDbParameter("121", CStr(var11(22))))
        cmd1.Parameters.Add(New OleDbParameter("122", CStr(var11(23))))
        cmd1.Parameters.Add(New OleDbParameter("123", CStr(var11(24))))
        cmd1.Parameters.Add(New OleDbParameter("124", CStr(var11(25))))
        cmd1.Parameters.Add(New OleDbParameter("125", CStr(var11(26))))

        Try
            cmd1.ExecuteNonQuery()
            cmd1.Dispose()
            Conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        password.Label2.Text = 1
        password.Show()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Form4.Label9.Text = 1
        Form4.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        GrafikTZ.CrystalReport31.SetParameterValue("start", Form5.Date3.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.date2.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("039", strg2)
        GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
        GrafikTZ.Show()
        Form5.Close()
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        GrafikTZ.CrystalReport31.SetParameterValue("start", Form5.Date3.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.date2.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("043", strg2)
        GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
        GrafikTZ.Show()
        Form5.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        GrafikTZ.CrystalReport31.SetParameterValue("start", Form5.Date3.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.date2.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("040", strg2)
        GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
        GrafikTZ.Show()
        Form5.Close()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        GrafikTZ.CrystalReport31.SetParameterValue("start", Form5.Date3.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.date2.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("041", strg2)
        GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
        GrafikTZ.Show()
        Form5.Close()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        GrafikTZ.CrystalReport31.SetParameterValue("start", Form5.Date3.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.date2.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("042", strg2)
        GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
        GrafikTZ.Show()
        Form5.Close()
    End Sub

    Private Sub Button118_Click_1(sender As Object, e As EventArgs) Handles Button118.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs)
        baru.Show()

    End Sub

    Private Sub Button119_Click_1(sender As Object, e As EventArgs) Handles Button119.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button120_Click_1(sender As Object, e As EventArgs) Handles Button120.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button121_Click_1(sender As Object, e As EventArgs) Handles Button121.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button122_Click_1(sender As Object, e As EventArgs) Handles Button122.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()

    End Sub

    Private Sub Button123_Click_1(sender As Object, e As EventArgs) Handles Button123.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button124_Click(sender As Object, e As EventArgs) Handles Button124.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button125_Click(sender As Object, e As EventArgs) Handles Button125.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)
        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg2)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub button104_Click_1(sender As Object, e As EventArgs) Handles button104.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub button105_Click_1(sender As Object, e As EventArgs) Handles button105.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button110_Click_1(sender As Object, e As EventArgs) Handles Button110.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()

    End Sub

    Private Sub Button112_Click_1(sender As Object, e As EventArgs) Handles Button112.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button113_Click_1(sender As Object, e As EventArgs) Handles Button113.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button114_Click_1(sender As Object, e As EventArgs) Handles Button114.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button115_Click_1(sender As Object, e As EventArgs) Handles Button115.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button117_Click_1(sender As Object, e As EventArgs) Handles Button117.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button116_Click_1(sender As Object, e As EventArgs) Handles Button116.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button111_Click_1(sender As Object, e As EventArgs) Handles Button111.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button109_Click_1(sender As Object, e As EventArgs) Handles Button109.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button108_Click_1(sender As Object, e As EventArgs) Handles Button108.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button107_Click_1(sender As Object, e As EventArgs) Handles Button107.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub button106_Click_1(sender As Object, e As EventArgs) Handles button106.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub button103_Click_1(sender As Object, e As EventArgs) Handles button103.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub button102_Click_1(sender As Object, e As EventArgs) Handles button102.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub button101_Click_1(sender As Object, e As EventArgs) Handles button101.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub button100_Click_1(sender As Object, e As EventArgs) Handles button100.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        grafik.CrystalGrafik1.SetParameterValue("start", Form5.Date3.Text)
        grafik.CrystalGrafik1.SetParameterValue("end", Form5.date2.Text)

        grafik.CrystalGrafik1.SetParameterValue("check0", strg2)
        grafik.CrystalGrafik1.SetParameterValue("check1", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check2", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check3", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check4", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check5", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check6", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check7", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check8", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check9", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check10", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check11", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check12", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check13", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check14", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check15", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check16", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check17", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check18", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check19", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check20", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check21", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check22", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check23", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check24", strg1)
        grafik.CrystalGrafik1.SetParameterValue("check25", strg1)
        grafik.Show()
        Form5.Close()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        GrafikTZ.CrystalReport31.SetParameterValue("start", Form5.Date3.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.date2.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("037", strg2)
        GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
        GrafikTZ.Show()
        Form5.Close()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        GrafikTZ.CrystalReport31.SetParameterValue("start", Form5.Date3.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.date2.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("038", strg2)
        GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
        GrafikTZ.Show()
        Form5.Close()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        GrafikTZ.CrystalReport31.SetParameterValue("start", Form5.Date3.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.date2.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("coba", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("032", strg2)
        GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("044", strg1)
        GrafikTZ.Show()

        Form5.Close()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        strg1 = "True"
        strg2 = "False"
        Form5.date2.Value = DateAdd("d", -3, Form5.date1.Value)
        Form5.Date3.Value = DateAdd("d", 1, Form5.date1.Value)
        Form5.Show()
        GrafikTZ.CrystalReport31.SetParameterValue("start", Form5.Date3.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("end", Form5.date2.Text)
        GrafikTZ.CrystalReport31.SetParameterValue("032", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("037", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("038", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("039", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("040", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("041", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("042", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("043", strg1)
        GrafikTZ.CrystalReport31.SetParameterValue("044", strg2)
        GrafikTZ.Show()
        Form5.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        password.Label2.Text = 2
        password.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form4.Show()
    End Sub


End Class
