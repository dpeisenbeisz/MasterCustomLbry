Public Class TranspArray
    Friend Function TransToAlpha() As Dictionary(Of Integer, Byte)
        Dim transp As New Dictionary(Of Integer, Byte)
        transp(0) = 255
        transp(1) = 252
        transp(2) = 249
        transp(3) = 247
        transp(4) = 244
        transp(5) = 242
        transp(6) = 239
        transp(7) = 237
        transp(8) = 234
        transp(9) = 232
        transp(10) = 229
        transp(11) = 226
        transp(12) = 224
        transp(13) = 221
        transp(14) = 219
        transp(15) = 216
        transp(16) = 214
        transp(17) = 211
        transp(18) = 209
        transp(19) = 206
        transp(20) = 204
        transp(21) = 201
        transp(22) = 198
        transp(23) = 196
        transp(24) = 193
        transp(25) = 191
        transp(26) = 188
        transp(27) = 186
        transp(28) = 183
        transp(29) = 181
        transp(30) = 178
        transp(31) = 175
        transp(32) = 173
        transp(33) = 170
        transp(34) = 168
        transp(35) = 165
        transp(36) = 163
        transp(37) = 160
        transp(38) = 158
        transp(39) = 155
        transp(40) = 153
        transp(41) = 150
        transp(42) = 147
        transp(43) = 145
        transp(44) = 142
        transp(45) = 140
        transp(46) = 137
        transp(47) = 135
        transp(48) = 132
        transp(49) = 130
        transp(50) = 127
        transp(51) = 124
        transp(52) = 122
        transp(53) = 119
        transp(54) = 117
        transp(55) = 114
        transp(56) = 112
        transp(57) = 109
        transp(58) = 107
        transp(59) = 104
        transp(60) = 102
        transp(61) = 99
        transp(62) = 96
        transp(63) = 94
        transp(64) = 91
        transp(65) = 89
        transp(66) = 86
        transp(67) = 84
        transp(68) = 81
        transp(69) = 79
        transp(70) = 76
        transp(71) = 73
        transp(72) = 71
        transp(73) = 68
        transp(74) = 66
        transp(75) = 63
        transp(76) = 61
        transp(77) = 58
        transp(78) = 56
        transp(79) = 53
        transp(80) = 51
        transp(81) = 48
        transp(82) = 45
        transp(83) = 43
        transp(84) = 40
        transp(85) = 38
        transp(86) = 35
        transp(87) = 33
        transp(88) = 30
        transp(89) = 28
        transp(90) = 25
        transp(91) = 22
        transp(92) = 20
        transp(93) = 17
        transp(94) = 15
        transp(95) = 12
        transp(96) = 10
        transp(97) = 7
        transp(98) = 5
        transp(99) = 2
        transp(100) = 0
        Return transp
    End Function

    Friend Function AlphaToTrans() As Dictionary(Of Integer, Integer)
        Dim untransp As New Dictionary(Of Integer, Integer)
        untransp(255) = 0
        untransp(252) = 1
        untransp(249) = 2
        untransp(247) = 3
        untransp(244) = 4
        untransp(242) = 5
        untransp(239) = 6
        untransp(237) = 7
        untransp(234) = 8
        untransp(232) = 9
        untransp(229) = 10
        untransp(226) = 11
        untransp(224) = 12
        untransp(221) = 13
        untransp(219) = 14
        untransp(216) = 15
        untransp(214) = 16
        untransp(211) = 17
        untransp(209) = 18
        untransp(206) = 19
        untransp(204) = 20
        untransp(201) = 21
        untransp(198) = 22
        untransp(196) = 23
        untransp(193) = 24
        untransp(191) = 25
        untransp(188) = 26
        untransp(186) = 27
        untransp(183) = 28
        untransp(181) = 29
        untransp(178) = 30
        untransp(175) = 31
        untransp(173) = 32
        untransp(170) = 33
        untransp(168) = 34
        untransp(165) = 35
        untransp(163) = 36
        untransp(160) = 37
        untransp(158) = 38
        untransp(155) = 39
        untransp(153) = 40
        untransp(150) = 41
        untransp(147) = 42
        untransp(145) = 43
        untransp(142) = 44
        untransp(140) = 45
        untransp(137) = 46
        untransp(135) = 47
        untransp(132) = 48
        untransp(130) = 49
        untransp(127) = 50
        untransp(124) = 51
        untransp(122) = 52
        untransp(119) = 53
        untransp(117) = 54
        untransp(114) = 55
        untransp(112) = 56
        untransp(109) = 57
        untransp(107) = 58
        untransp(104) = 59
        untransp(102) = 60
        untransp(99) = 61
        untransp(96) = 62
        untransp(94) = 63
        untransp(91) = 64
        untransp(89) = 65
        untransp(86) = 66
        untransp(84) = 67
        untransp(81) = 68
        untransp(79) = 69
        untransp(76) = 70
        untransp(73) = 71
        untransp(71) = 72
        untransp(68) = 73
        untransp(66) = 74
        untransp(63) = 75
        untransp(61) = 76
        untransp(58) = 77
        untransp(56) = 78
        untransp(53) = 79
        untransp(51) = 80
        untransp(48) = 81
        untransp(45) = 82
        untransp(43) = 83
        untransp(40) = 84
        untransp(38) = 85
        untransp(35) = 86
        untransp(33) = 87
        untransp(30) = 88
        untransp(28) = 89
        untransp(25) = 90
        untransp(22) = 91
        untransp(20) = 92
        untransp(17) = 93
        untransp(15) = 94
        untransp(12) = 95
        untransp(10) = 96
        untransp(7) = 97
        untransp(5) = 98
        untransp(2) = 99
        untransp(0) = 100
        Return untransp
    End Function

    Friend Function GetAlpha(dec As Integer) As Byte

        Dim transp As Dictionary(Of Integer, Byte) = TransToAlpha()

        If dec > 100 Or dec < 0 Then
            Return Nothing
        Else
            Return transp(dec)
        End If

    End Function

    Friend Function GetTrans(t As Byte) As Integer

        If t > 255 Or t < 0 Then
            Return 0
        Else
            Dim tempDec As Integer = untransp(t)
            If tempDec = 0 And t <> 255 Then
                Return 0
            Else
                Return tempDec
            End If
        End If
    End Function

    Function getTransparencyIndex(ByVal lay As Autodesk.AutoCAD.DatabaseServices.LayerTableRecord) As Integer
        If lay.Transparency.IsInvalid = False AndAlso lay.Transparency.IsByAlpha = True Then
            Dim smallestDiff As Integer = Byte.MaxValue, foundValue As Integer = 0
            Dim transparencies As Dictionary(Of Integer, Byte) = TransToAlpha()
            For Each key As Byte In transparencies.Keys
                If Math.Abs(transparencies(key) - lay.Transparency.Alpha) < smallestDiff Then
                    smallestDiff = Math.Abs(transparencies(key) - lay.Transparency.Alpha)
                    foundValue = key
                End If
            Next
            Return foundValue
        Else
            Return 0
        End If
    End Function

    Function getTransparencyValue(ByVal index As Integer) As Autodesk.AutoCAD.Colors.Transparency
            If Transparencies.ContainsKey(index) = True Then
                Return New Autodesk.AutoCAD.Colors.Transparency(Transparencies(index))
            Else
                Return New Autodesk.AutoCAD.Colors.Transparency(Byte.MaxValue)
            End If
        End Function

End Class
