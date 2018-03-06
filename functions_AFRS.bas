Attribute VB_Name = "functions_AFRS"

'Converts AFRISS dates in the format YYYYMMDD to Excel's standard date format
Public Function convertAfrissDate(afrissDate As Variant)

    If Len(CStr(afrissDate)) = 8 Then
        convertAfrissDate = DateSerial(Left(afrissDate, 4), Mid(afrissDate, 5, 2), Mid(afrissDate, 7, 2))
    Else
        convertAfrissDate = 0
    End If

End Function


'Find fiscal year
Public Function fiscalYear(dt As Variant)

    If Month(dt) >= 10 Then
        fiscalYear = Year(dt) + 1 - 2000
    Else
        fiscalYear = Year(dt) - 2000
    End If

End Function



Public Function ssnRemoveDash(ssn As String)

    If Len(ssn) = 11 Then
        ssnRemoveDash = Val(Replace(ssn, "-", ""))
    ElseIf Len(ssn) = 9 And Len(CStr(Val(ssn))) = 9 Then
        ssnRemoveDash = Val(ssn)
    End If


End Function



'Fixes AFSC formatting errors:
' -Converting to scientific notation
Public Function fixAfsc(afsc As String)

    fixAfsc = Replace(afsc, "+", "0")

End Function



'Generates a random SSN
Public Function randomSSN(Optional includeDashes As Boolean)

    If includeDashes = True Then
        randomSSN = Format(Round(Rnd * 999), "000") & "-" & Format(Round(Rnd * 99), "00") & "-" & Format(Round(Rnd * 9999), "0000")
    Else
        randomSSN = Round(Rnd * 999999999)
    End If

End Function



Public Function mageAll(GS As Integer, AR As Integer, WK As Integer, PC As Integer, MK As Integer, EI As Integer, ASx As Integer, MC As Integer, AO As Integer, VE As Integer)
    
    Dim M As Integer, A As Integer, G As Integer, E As Integer
    
    M = mageM(GS, AR, WK, PC, MK, EI, ASx, MC, AO, VE)
    A = mageA(GS, AR, WK, PC, MK, EI, ASx, MC, AO, VE)
    G = mageG(GS, AR, WK, PC, MK, EI, ASx, MC, AO, VE)
    E = mageE(GS, AR, WK, PC, MK, EI, ASx, MC, AO, VE)

    mageAll = "M: " & M & ", A: " & A & ", G: " & G & ", E: " & E

End Function


Public Function mageM(GS As Integer, AR As Integer, WK As Integer, PC As Integer, MK As Integer, EI As Integer, ASx As Integer, MC As Integer, AO As Integer, VE As Integer) As Integer

    Dim mScore As Integer, M As Integer
    
    
    mScore = AR + 2 * VE + MC + ASx
    Select Case mScore
        Case Is <= 133: M = 1
        Case 147 To 155: M = 2
        Case 156 To 163: M = 3
        Case 164 To 169: M = 4
        Case 170 To 174: M = 5
        Case 175 To 179: M = 6
        Case 180 To 184: M = 7
        Case 185 To 187: M = 8
        Case 188 To 191: M = 9
        Case 192 To 194: M = 10
        Case 195 To 197: M = 11
        Case 198 To 199: M = 12
        Case 200 To 201: M = 13
        Case 202 To 204: M = 14
        Case 205 To 207: M = 15
        Case 208 To 209: M = 16
        Case 210: M = 17
        Case 211 To 212: M = 18
        Case 213 To 214: M = 19
        Case 215: M = 20
        Case 216 To 217: M = 21
        Case 218 To 219: M = 22
        Case 220: M = 23
        Case 221 To 222: M = 24
        Case 223: M = 25
        Case 224 To 225: M = 26
        Case 226: M = 27
        Case 227 To 228: M = 28
        Case 229 To 230: M = 29
        Case 231: M = 30
        Case 232 To 233: M = 31
        Case 234 To 235: M = 33
        Case 236: M = 34
        Case 237: M = 35
        Case 238: M = 36
        Case 239 To 240: M = 37
        Case 241: M = 38
        Case 242: M = 39
        Case 243: M = 40
        Case 244: M = 41
        Case 245: M = 42
        Case 246: M = 43
        Case 247: M = 44
        Case 248: M = 45
        Case 249: M = 46
        Case 250 To 251: M = 47
        Case 252: M = 48
        Case 253: M = 49
        Case 254: M = 50
        Case 255: M = 51
        Case 256: M = 52
        Case 257: M = 53
        Case 258: M = 54
        Case 259: M = 55
        Case 260: M = 56
        Case 261: M = 57
        Case 262: M = 58
        Case 263: M = 59
        Case 264: M = 60
        Case 265: M = 61
        Case 266: M = 62
        Case 267: M = 63
        Case 268: M = 64
        Case 269: M = 65
        Case 270: M = 66
        Case 271: M = 67
        Case 272: M = 68
        Case 273: M = 69
        Case 274: M = 70
        Case 275: M = 71
        Case 276: M = 72
        Case 277: M = 73
        Case 278: M = 74
        Case 279: M = 75
        Case 280: M = 76
        Case 281: M = 77
        Case 282 To 283: M = 78
        Case 284: M = 79
        Case 285: M = 80
        Case 286 To 287: M = 81
        Case 288: M = 82
        Case 289: M = 83
        Case 290: M = 84
        Case 291: M = 85
        Case 292 To 293: M = 86
        Case 294 To 295: M = 87
        Case 296 To 297: M = 88
        Case 298: M = 89
        Case 299 To 300: M = 90
        Case 301 To 302: M = 91
        Case 303 To 304: M = 92
        Case 305 To 307: M = 93
        Case 308 To 309: M = 94
        Case 310 To 313: M = 95
        Case 314 To 316: M = 96
        Case 317 To 321: M = 97
        Case 322 To 328: M = 98
        Case Is >= 329: M = 99
    End Select
    
    mageM = M
    
End Function


Public Function mageA(GS As Integer, AR As Integer, WK As Integer, PC As Integer, MK As Integer, EI As Integer, ASx As Integer, MC As Integer, AO As Integer, VE As Integer) As Integer

    Dim aScore As Integer, A As Integer

    aScore = VE + MK
    Select Case aScore
        Case Is <= 49: A = 1
        Case 56 To 59: A = 2
        Case 60 To 63: A = 3
        Case 64 To 66: A = 4
        Case 67 To 68: A = 5
        Case 69 To 70: A = 6
        Case 71 To 72: A = 7
        Case 73: A = 8
        Case 74: A = 9
        Case 75 To 76: A = 10
        Case 77: A = 11
        Case 78: A = 13
        Case 79: A = 14
        Case 80: A = 15
        Case 81: A = 16
        Case 82: A = 17
        Case 83: A = 19
        Case 84: A = 20
        Case 85: A = 21
        Case 86: A = 23
        Case 87: A = 25
        Case 88: A = 27
        Case 89: A = 29
        Case 90: A = 30
        Case 91: A = 32
        Case 92: A = 34
        Case 93: A = 35
        Case 94: A = 37
        Case 95: A = 39
        Case 96: A = 41
        Case 97: A = 43
        Case 98: A = 45
        Case 99: A = 47
        Case 100: A = 49
        Case 101: A = 50
        Case 102: A = 52
        Case 103: A = 55
        Case 104: A = 56
        Case 105: A = 59
        Case 106: A = 61
        Case 107: A = 63
        Case 108: A = 65
        Case 109: A = 67
        Case 110: A = 69
        Case 111: A = 71
        Case 112: A = 72
        Case 113: A = 74
        Case 114: A = 76
        Case 115: A = 78
        Case 116: A = 80
        Case 117: A = 82
        Case 118: A = 84
        Case 119: A = 85
        Case 120: A = 87
        Case 121: A = 88
        Case 122: A = 90
        Case 123: A = 91
        Case 124: A = 92
        Case 125: A = 93
        Case 126: A = 94
        Case 127: A = 95
        Case 128 To 129: A = 96
        Case 130 To 131: A = 97
        Case 132 To 133: A = 98
        Case Is >= 134: A = 99
    End Select

    mageA = A
    
End Function




Public Function mageG(GS As Integer, AR As Integer, WK As Integer, PC As Integer, MK As Integer, EI As Integer, ASx As Integer, MC As Integer, AO As Integer, VE As Integer) As Integer

    Dim gScore As Integer, G As Integer

    gScore = AR + VE
    Select Case gScore
        Case Is <= 45: G = 1
        Case 52 To 57: G = 2
        Case 58 To 61: G = 3
        Case 62 To 65: G = 4
        Case 66 To 67: G = 5
        Case 68 To 69: G = 6
        Case 70 To 71: G = 7
        Case 72 To 73: G = 8
        Case 74: G = 9
        Case 75 To 76: G = 10
        Case 77: G = 11
        Case 78: G = 12
        Case 79: G = 13
        Case 80: G = 14
        Case 81: G = 15
        Case 82: G = 16
        Case 83: G = 17
        Case 84: G = 19
        Case 85: G = 20
        Case 86: G = 21
        Case 87: G = 23
        Case 88: G = 24
        Case 89: G = 26
        Case 90: G = 28
        Case 91: G = 30
        Case 92: G = 32
        Case 93: G = 33
        Case 94: G = 36
        Case 95: G = 38
        Case 96: G = 40
        Case 97: G = 42
        Case 98: G = 44
        Case 99: G = 47
        Case 100: G = 49
        Case 101: G = 51
        Case 102: G = 53
        Case 103: G = 55
        Case 104: G = 57
        Case 105: G = 59
        Case 106: G = 62
        Case 107: G = 64
        Case 108: G = 66
        Case 109: G = 68
        Case 110: G = 70
        Case 111: G = 72
        Case 112: G = 74
        Case 113: G = 76
        Case 114: G = 78
        Case 115: G = 80
        Case 116: G = 81
        Case 117: G = 83
        Case 118: G = 84
        Case 119: G = 85
        Case 120: G = 87
        Case 121: G = 88
        Case 122: G = 89
        Case 123: G = 91
        Case 124: G = 92
        Case 125: G = 93
        Case 126: G = 94
        Case 127 To 128: G = 95
        Case 129 To 130: G = 96
        Case 131 To 132: G = 97
        Case 133 To 134: G = 98
        Case Is >= 135: G = 99
    End Select
    
    mageG = G

End Function



Public Function mageE(GS As Integer, AR As Integer, WK As Integer, PC As Integer, MK As Integer, EI As Integer, ASx As Integer, MC As Integer, AO As Integer, VE As Integer) As Integer

    Dim eScore As Integer, E As Integer
    
    eScore = GS + AR + MK + EI
    Select Case eScore
        Case Is <= 100: E = 1
        Case 120 To 127: E = 2
        Case 128 To 133: E = 3
        Case 134 To 137: E = 4
        Case 138 To 141: E = 5
        Case 142 To 144: E = 6
        Case 145 To 148: E = 7
        Case 149 To 150: E = 8
        Case 151 To 153: E = 9
        Case 154 To 155: E = 10
        Case 156 To 157: E = 11
        Case 158 To 159: E = 12
        Case 160 To 161: E = 13
        Case 162 To 163: E = 14
        Case 164: E = 15
        Case 165 To 166: E = 16
        Case 167 To 168: E = 17
        Case 169: E = 18
        Case 170 To 171: E = 19
        Case 172: E = 20
        Case 173: E = 21
        Case 174: E = 22
        Case 175 To 176: E = 23
        Case 177: E = 24
        Case 178: E = 25
        Case 179: E = 26
        Case 180: E = 27
        Case 181: E = 28
        Case 182: E = 30
        Case 183 To 184: E = 31
        Case 185: E = 33
        Case 186: E = 34
        Case 187: E = 35
        Case 188: E = 36
        Case 189: E = 37
        Case 190: E = 38
        Case 191: E = 39
        Case 192: E = 40
        Case 193: E = 41
        Case 194: E = 43
        Case 195: E = 44
        Case 196: E = 45
        Case 197: E = 46
        Case 198: E = 47
        Case 199: E = 49
        Case 200: E = 50
        Case 201: E = 51
        Case 202: E = 52
        Case 203: E = 53
        Case 204: E = 54
        Case 205: E = 55
        Case 206: E = 56
        Case 207: E = 58
        Case 208: E = 59
        Case 209: E = 60
        Case 210: E = 61
        Case 211: E = 62
        Case 212: E = 64
        Case 213: E = 65
        Case 214: E = 66
        Case 215: E = 67
        Case 216: E = 68
        Case 217: E = 69
        Case 218: E = 70
        Case 219: E = 71
        Case 220: E = 72
        Case 221: E = 73
        Case 222: E = 74
        Case 223: E = 75
        Case 224 To 225: E = 76
        Case 226: E = 77
        Case 227: E = 79
        Case 228: E = 80
        Case 229 To 230: E = 81
        Case 231: E = 82
        Case 232: E = 83
        Case 233 To 234: E = 84
        Case 235: E = 85
        Case 236: E = 86
        Case 237 To 238: E = 87
        Case 239 To 240: E = 88
        Case 241 To 242: E = 89
        Case 243: E = 90
        Case 244 To 246: E = 91
        Case 247 To 248: E = 92
        Case 249 To 250: E = 93
        Case 251 To 253: E = 94
        Case 254 To 256: E = 95
        Case 257 To 260: E = 96
        Case 261 To 264: E = 97
        Case 265 To 271: E = 98
        Case Is >= 272: E = 99
    End Select
        
    mageE = E
    
End Function




Function shipWeek(dtead As Date)

    shipWeek = dtead - (dtead + 4) Mod 7

End Function







