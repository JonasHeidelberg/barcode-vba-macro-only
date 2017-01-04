Rem  *****  BASIC  *****
Rem This software is distributd under The MIT License (MIT)
Rem Copyright © 2013 Madeta a.s. Jiri Gabriel
Rem Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
Rem The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
Rem THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
Rem
Rem Some code comments translated from Czech into English using an online translator by JonasHeidelberg - careful, might be quite wrong
Option Explicit
Const BCEnc128$ = "1A1B1B1B1A1B1B1B1A0B0B1C0B0C1B0C0B1B0B1B0C0B1C0B0C1B0B1B0B0C1B0C0B1C0B0B0A1B2B0B1A2B0B1B2A0A2B1B0B2A1B0B2B1A1B2B0A1B0A2B1B0B2A1A2B0B1B2A0B2A1A2A2A0B1B2B0A1B2B0B1A2A1B0B2B1A0B2B1B0A1A1A1C1A1C1A1C1A1A0A0C1C0C0A1C0C0C1A0A1C0C0C1A0C0C1C0A1A0C0C1C0A0C1C0C0A0A1A2C0A1C2A0C1A2A0A2A1C0A2C1A0C2A1A2A2A1A1A0C2A1C0A2A1A2A0C1A2C0A1A2A2A2A0A1C2A0C1A2C0A1A2A1A0C2A1C0A2C1A0A2A3A0A1B0D0A3C0A0A0A0B1D0A0D1B0B0A1D0B0D1A0D0A1B0D0B1A0A1B0D0A1D0B0B1A0D0B1D0A0D1A0B0D1B0A1D0B0A1B0A0D3A2A0A1D0A0B0C3A0A0A0B3B0B0A3B0B0B3A0A3B0B0B3A0B0B3B0A3A0B0B3B0A0B3B0B0A1A1A3A1A3A1A3A1A1A0A0A3C0A0C3A0C0A3A0A3A0C0A3C0A3A0A0C3A0C0A0A2A3A0A3A2A2A0A3A3A0A2A1A0D0B1A0B0D1A0B2B1C2A0A1"
Const BCEncE13$ = "C6A5A5B77B5AB6B5A6B66B6AB5B6B6A66A6BA8A5A5D55D5AA5C6B7A55A7BA6C5A7B55B7AA5A8D5A55A5DA7A6B5C55C5BA6A7C5B55B5CC5A6B5A77A5B"
Const BCEnc39$ = "0A0C3A3A03A0C0A0A30A3C0A0A33A3C0A0A00A0C3A0A33A0C3A0A00A3C3A0A00A0C0A3A33A0C0A3A00A3C0A3A03A0A0C0A30A3A0C0A33A3A0C0A00A0A3C0A33A0A3C0A00A3A3C0A00A0A0C3A33A0A0C3A00A3A0C3A00A0A3C3A03A0A0A0C30A3A0A0C33A3A0A0C00A0A3A0C33A0A3A0C00A3A3A0C00A0A0A3C33A0A0A3C00A3A0A3C00A0A3A3C03C0A0A0A30C3A0A0A33C3A0A0A00C0A3A0A33C0A3A0A00C3A3A0A00C0A0A3A33C0A0A3A00C3A0A3A00C0C0C0A00C0C0A0C00C0A0C0C00A0C0C0C0"
Const BCChs39$ = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%"
Const BCExt39$ = "%U$A$B$C$D$E$F$G$H$I$J$K$L$M$N$O$P$Q$R$S$T$U$V$W$X$Y$Z%A%B%C%D%Esp/A/B/C/D/E/F/G/H/I/J/K/L - ./O 0 1 2 3 4 5 6 7 8 9/Z%F%G%H%iJ%V A B C D E F G H I J K L M N O P Q R S T U V W X Y Z%K%L%M%N%O%W+A+B+C+D+E+F+G+H+I+J+K+L+M+N+O+P+Q+R+S+T+U+V+W+X+Y+Z%P%Q%R%S%T"
Const BCEnc25$ = "00110100010100111000001011010001100000111001001010AABBABAAABABAABBBAAAAABABBABAAABBAAAAABBBAABAABABA"
Const qralnum$ = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:"

Dim IsMs As Boolean
Sub Init()
  If VarType(Asc("A")) = 2 Then IsMs = True Else IsMs = False
End Sub
' MS: EncodeBarcode(POLICKO("SHEET");POLICKO("ADDRESS");A2;1;1;0;2)
'    Usage: EncodeBarcode(CELL("SHEET");CELL("ADDRESS");A2;1;1;0;2)
'                                                       /  | | | \
'                        The barcode content (string) -'   / | \  `- Protection zone for 1D codes (width)
' 0-Code128,1-EAN,2-2of5I,3-Code39,50-Datamatrix,51-QR ---'  /  `--- Parameters (depending on the barcode type)
'                       Output type: graphic=1,Font=0  -----'
Public Function EncodeBarcode(ShIx As Integer, xAddr As String, _
                code As String, pbctype%, Optional pgraficky%, _
                Optional pparams%, Optional pzones%) As String
  Dim s$, bctype%, graficky%, params%, zones%
  Dim oo As Object

  Call Init
  If IsMissing(pzones) Then zones = 2 Else zones = pzones
  If IsMissing(pparams) Then params = 0 Else params = pparams
  If IsMissing(pgraficky) Then graficky = 1 Else graficky = pgraficky
  If IsMissing(pbctype) Then bctype = 0 Else bctype = pbctype
  Select Case bctype
    Case 1 ' EAN8/13/UPCA/UPCE
           ' params 1,2,3,4 = EAN13,EAN8,UPCA,UPCE - type
           '        + 8 add checksum
      s = bc_EAN(code, params, zones)
    Case 2 ' Two of Five Interleaved
      s = bc_25I(code, zones)
    Case 3 ' Code39
           ' params extended charset 2 = disabled, 1 = always, 0 = automaticaly
           '        + 7 = add checksum
      s = bc_Code39(code, params, zones)
    Case 50 ' DataMatrix params: 1 = force ASCII encoding
      s = dmx_gen(code, Iif(params = 1, "ASCII", ""))
    Case 51 ' QRCode params: ECLevel 0=M 1=L 2=Q 3=H
      s = "mode=" & Mid("MLQH", (params Mod 4) + 1, 1)
      s = qr_gen(code, s)
    Case Else ' Code128
           ' params 1 = start subset A   2 = start subset B   3 = start subset C
      s = bc_Code128(code, params, zones)
  End Select
  If graficky <> 0 Then
    If bctype >= 50 Then
      If IsMs Then
        Call bc_2Dms(s)
      Else
        Call bc_2D(ShIx, xAddr, s)
      End If
    Else
      If IsMs Then
        Call bc_1Dms(s)
      Else
        Call bc_1D(ShIx, xAddr, s)
      End If
    End If
    EncodeBarcode = ""
  Else
    EncodeBarcode = s
  End If
  Exit Function
End Function

Function AscL(s As String) As Long
  If IsMs Then AscL = AscW(s) Else AscL = Asc(s)
End Function

Function bc_25I(chaine$, Optional zones%) As String
  ' start = "0A0A" stop = "1A0"
  Dim i%, j%, k%, l%, s$, q$, zon$
  If IsMissing(zones%) Then
    zon$ = "DD"
  Else
    zon$ = Iif(zones% <= 0, "", Mid$("DDDDDDDDDD", 1, zones%))
  End If
  q = chaine
  s = ""
  For i = 1 To Len(q)
    j = (AscL(Mid(q, i, 1)) Mod 256) - 48
    If (j >= 0 And j <= 9) Then s = s & Chr(48 + j)
  Next
  i = Len(s)
  If i <= 0 Then
    bc_25I = ""
    Exit Function
  End If
  If (i Mod 2) = 1 Then s = "0" & s
  q = zon & "0A0A" ' Start
  For i = 1 To Len(s) Step 2
    j = val(Mid(s, i, 1)) * 5
    k = 50 + val(Mid(s, i + 1, 1)) * 5
    For l = 1 To 5
      q = q & Mid(BCEnc25, j + l, 1) & Mid(BCEnc25, k + l, 1)
    Next
  Next
  bc_25I = q & "01A0" & zon
End Function

Function bc_Code39(chaine$, Optional params%, Optional zones%) As String
  ' params extended charset 2 = disabled, 1 = always, 0 = automaticaly
  '         4 = add checksum
  '[bWbwBwBwb]w[BwbwbWbwB]w[bWbwBwBwb] start = 0C0A2A2A0A  stop = A0C0A2A2A0
  Dim i, j%, s$, p$, q$, zon$, ext%, ch%, check%
  If IsMissing(zones) Then
    zon$ = "DD"
  Else
    zon$ = Iif(zones <= 0, "", Mid("DDDDDDDDDD", 1, zones))
  End If
  If IsMissing(params) Then
    check = 0
    ext = 0
  Else
    check = Int(params / 4) Mod 2
    ext = (params Mod 4) - 1
  End If
  s = chaine
  If Len(s) <= 0 Then
    bc_Code39 = ""
    Exit Function
  End If
  If ext = -1 Then
    ' Need extend ?
    For i = 1 To Len(s)
      p = Mid(s, i, 1)
      j = InStr(BCChs39, p)
      If j <= 0 Or AscL(p) > 90 Then
        ext = 1
        Exit For
      End If
    Next
  End If
  If ext = 1 Then
    p = s
    s = ""
    For i = 1 To Len(p)
      j = AscL(Mid(p, i, 1)) Mod 256
      If j = 32 Then
        s = s & " "
      ElseIf (j <= 127) Then
        s = s & Trim(Mid(BCExt39, 1 + j * 2, 2))
      End If
    Next
  End If
  q = zon & "0C0A2A2A0A" ' Start *
  ch = 0
  For i = 1 To Len(s)
    p = Mid(s, i, 1)
    j = InStr(BCChs39, p) - 1
    If j >= 0 And j < 43 Then
      ch = (ch + j) Mod 43
      q = q & Mid(BCEnc39, j * 9 + 1, 9) & "A"
    End If
  Next
  If check = 1 Then q = q & Mid(BCEnc39, ch * 9 + 1, 9) & "A"
  bc_Code39 = q & "0C0A2A2A0" & zon
End Function

Function bc_EAN(chaine$, Optional params%, Optional zones%) As String
  'Parameters : String up to 13 chars wide,
  ' params 1,2,3,4 = EAN13,EAN8,UPCA,UPCE - type
  '        + 8 add checksum
  Dim i%, j%, checksum%, first%, CodeBarre$, s$, p$, q$, zon$, subtyp%, check%
  Dim tableA As Boolean
  If IsMissing(zones) Then
    zon$ = "DD"
  Else
    zon$ = Iif(zones <= 0, "", Mid("DDDDDDDDDD", 1, zones))
  End If
  If IsMissing(params) Then
    check = 0
    subtyp = 0
  Else
    check = Int(params / 8) Mod 2
    subtyp = params Mod 8
  End If
  s = chaine
  p = ""
  CodeBarre = zon
  For i = 1 To Len(s)
    j = AscL(Mid(s, i, 1)) Mod 256
    If j >= 48 Or j <= 57 Then p = p & Chr(j)
  Next i
  s = p
  If subtyp = 4 Then
    While Len(s) < 6
      s = "0" & s
    Wend
    If Len(s) > 6 Then s = Left(s, 6)
    p = s
    first = val(right(p, 1))
    If first >= 5 Then
      s = "00" & Left(p, 5) & "0000" & right(p, 1)
    ElseIf first = 4 Then
      s = "00" & Left(p, 4) & "00000" & Mid(p, 5, 1)
    ElseIf first = 3 Then
      s = "00" & Left(p, 3) & "00000" & Mid(p, 4, 2)
    Else
      s = "00" & Left(p, 2) & right(p, 1) & "0000" & Mid(p, 3, 3)
    End If
  End If
  If check = 1 Or subtyp = 4 Then s = s & "0"
  While Len(s) < 13
    s = "0" & s
  Wend
  checksum = 0
  first = 1
  For i = 1 To 12
    j = AscL(Mid(s, i, 1)) Mod 256
    checksum = (checksum + first * (j - 48)) Mod 10
    first = (first + 2) Mod 4
  Next
  'Kontrolni soucet
  s = Left(s, 12) & Chr(48 + (10 - checksum Mod 10) Mod 10)
  If subtyp = 4 Then
    s = "000000" & right(s, 1) & p
  End If
  If Left(s, 12) <> "000000000000" Then
    CodeBarre = CodeBarre & "0A0"
    If subtyp = 0 And Left(s, 5) = "00000" Then subtyp = 2 ' EAN8
    If subtyp = 0 And Left(s, 1) = "0" Then subtyp = 3 ' UPC-A
    ' Jinak EAN13
    If subtyp = 0 Then subtyp = 1
    If subtyp = 2 Then ' EAN8
      j = 5
      p = "0000LLLLRRRR"
    ElseIf subtyp = 3 Then ' UPC-A
      j = 1
      p = "LLLLLLRRRRRR"
    ElseIf subtyp = 4 Then ' UPC-E
      first = val(Mid(s, 7, 1)) ' check
      j = 7
      p = "000000" & Mid("GGGLLLGGLGLLGGLLGLGGLLLGGLGGLLGLLGGLGLLLGGGLGLGLGLGLLGGLLGLG", 1 + first * 6, 6)
    Else ' EAN13
      j = 1
      first = val(Left(s, 1))
      p = Mid("LLLLLLLLGLGGLLGGLGLLGGGLLGLLGGLGGLLGLGGGLLLGLGLGLGLGGLLGGLGL", 1 + first * 6, 6) + "RRRRRR"
    End If
    For i = j To 12
      first = val(Mid(s, i + 1, 1))
'      L       G       R     BarsAndSpaces                          G=rev(R)=rev(inv(L)) R=Inv(L)
'0      0001101 0100111 1110010 C1A0 A0B2 2B0A
'1      0011001 0110011 1100110 B1B0 A1B1 1B1A
'2      0010011 0011011 1101100 B0B1 B1A1 1A1B
'3      0111101 0100001 1000010 A3A0 A0D0 0D0A
'4      0100011 0011101 1011100 A0C1 B2A0 0A2B
'5      0110001 0111001 1001110 A1C0 A2B0 0B2A
'6      0101111 0000101 1010000 A0A3 D0A0 0A0D
'7      0111011 0010001 1000100 A2A1 B0C0 0C0B
'8      0110111 0001001 1001000 A1A2 C0B0 0B0C
'9      0001011 0010111 1110100 C0A1 B0A2 2A0B
      q = Mid(BCEncE13, 1 + first * 12, 12)
      Select Case Mid(p, i, 1)
        Case "L"
          CodeBarre = CodeBarre & Mid(q, 1, 4)
        Case "G"
          CodeBarre = CodeBarre & Mid(q, 5, 4)
        Case "R"
          CodeBarre = CodeBarre & Mid(q, 9, 4)
      End Select
      Select Case subtyp
        Case 1: If i = 6 Then CodeBarre = CodeBarre & "A0A0A"
        Case 3: If i = 6 Then CodeBarre = CodeBarre & "A0A0A"
        Case 2: If i = 8 Then CodeBarre = CodeBarre & "A0A0A"
      End Select
    Next
    If subtyp = 4 Then CodeBarre = CodeBarre & "A0A"
    CodeBarre = CodeBarre & "0A0"
  End If
  bc_EAN = CodeBarre & zon
End Function

Function bc_Code128(chaine$, Optional params%, Optional zones%) As String
  'Parameters : a string
  'Return : * a string which give the bar code when it is dispayed with BarsAndSpaces.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i%, checksum&, checkw&, min$, n%, zon$, s$, c128$, tbl$, q$, j%
  If IsMissing(zones) Then
    zon$ = "DD"
  Else
    zon$ = Iif(zones <= 0, "", Mid("DDDDDDDDDD", 1, zones))
  End If
'  If IsMissing(params) Then
'  Else
  c128 = ""
  s = chaine
  If Len(s) <= 0 Then
    bc_Code128 = ""
    Exit Function
  End If
  'Calculation of the code string with optimized use of tables A,B and C
  min = ""
  If (params Mod 4) >= 1 And (params Mod 4) <= 3 Then
    tbl = Mid("ABC", params Mod 4, 1)
  Else
    tbl = ""
  End If
  i = 1 'i become the string index
  Do While i <= Len(s)
    n = AscL(Mid(s, i, 1)) Mod 256
    If n = 95 Then ' _ escape _1 FNC1 .. _4 FNC4 __ = podtrzitko
      i = i + 1
      If i > Len(s) Then n = 0 Else n = AscL(Mid(s, i, 1)) Mod 256
      If (n >= 49 And n <= 52) Then
        n = 48 - n
      ElseIf n >= 64 And n <= 94 Then
        n = n - 64
        ' _@,_A .. _^ = 0x00..0x1E
      ElseIf n = 48 Then
        ' _0 = 0x1F
        n = 31
      Else
        n = 95
      End If
    End If
    If n >= 128 Then
      n = n Mod 128
      min = min & "z"
      c128 = c128 & "-05" ' shift
    End If
    Select Case n
      Case 48 To 57, -1
        min = min & "C"
      Case -4 To -2
        min = min & "z"
      Case 0 To 31
        min = min & "A"
      Case 32 To 63
        min = min & "z"
      Case Else ' 64 to 127
        min = min & "B"
    End Select
    q = "000" & Trim(CStr(Abs(n)))
    If n < 0 Then q = "-" & right(q, 2) Else q = right(q, 3)
    c128 = c128 & q
    i = i + 1
  Loop
  s = zon
  If tbl = "" Then
    If Left(min, 4) = "CCCC" Then
      tbl = "C"
    ElseIf InStr(min, "A") <= 0 Or Left(min, 1) = "B" Then
      tbl = "B"
    Else
      tbl = "A"
    End If
  End If
  n = 103 + AscL(tbl) - 65 ' 103,104,105 = Start A,B,C
  s = s & Mid(BCEnc128, 6 * n + 1, 6)
  checksum = n
  checkw = 1
  i = 1
  Do While i <= Len(min)
    n = val(Mid(c128, -2 + (i * 3), 3))
    q = Mid(min, i, 1)
    Select Case tbl
      Case "C"
        If q <> "C" Then
          If q = "A" Or (q = "z" And InStr(Mid(min, i), "B") < 0) Then
            tbl = "A"
            n = 101
          Else
            tbl = "B"
            n = 100
          End If
          i = i - 1
        Else
          If (n = -1) Then
            n = 102 ' Fnc 1
          Else
            ' Dvojcislo
            j = (n - 48) * 10
            If (i >= Len(min) Or Mid(min, i + 1, 1) <> "C") Then
              tbl = "B"
              n = 100
              i = i - 1
            Else
              i = i + 1
              n = val(Mid(c128, -2 + (i * 3), 3))
              If n < 0 Then
                tbl = "B"
                n = 100
                i = i - 2
              Else
                n = j + (n - 48)
              End If
            End If
          End If
        End If
      Case "A"
        If q = "B" Then
          n = 100 ' Switch to B
          i = i - 1
          tbl = "B"
        ElseIf Mid(min, i, 4) = "CCCC" Then
          n = 99 ' Start C
          i = i - 1
          tbl = "C"
        Else
          Select Case n
          Case -5: n = 98
          Case -4: n = 101
          Case -3: n = 96
          Case -2: n = 97
          Case -1: n = 102
          Case 0 To 31
            n = n + 64
          Case Else
            n = n - 32
          End Select
        End If
      Case "B"
        If q = "A" Then
          n = 101 ' Switch to B
          i = i - 1
          tbl = "A"
        ElseIf Mid(min, i, 4) = "CCCC" Then
          n = 99 ' Start C
          i = i - 1
          tbl = "C"
        Else
          Select Case n
          Case -5: n = 98
          Case -4: n = 100
          Case -3: n = 96
          Case -2: n = 97
          Case -1: n = 102
          Case Else
            n = n - 32
          End Select
        End If
    End Select
    If n >= 0 And n <= 102 Then
      s = s & Mid(BCEnc128, 6 * n + 1, 6)
      checksum = (checksum + checkw * n) Mod 103
      checkw = checkw + 1
    End If
    i = i + 1
  Loop
  n = checksum Mod 103
  s = s & Mid(BCEnc128, 6 * n + 1, 6)
  s = s + "1C2A0A1"
  bc_Code128 = s & zon
End Function

Function dmx_place(parr As Variant, psiz As Integer, _
                   pbl As Integer, prow As Integer, pcol As Integer, _
                   pbit As Integer) As Boolean
  Dim ix%, va%, r%, c%, s%
  r = prow
  c = pcol
  If psiz > 0 Then
    s = psiz / pbl
    If r < 0 Then
      r = r + psiz
      c = c + 4 - ((psiz + 4) Mod 8)
    End If
    If c < 0 Then
      c = c + psiz
      r = r + 4 - ((psiz + 4) Mod 8)
    End If
    If c >= psiz Then
      c = c - psiz
      r = r + 1
    End If
    r = r + (Int(r / s) * 2)
    c = c + (Int(c / s) * 2)
  End If
  dmx_place = False
  r = r + 2
  c = c + 2
  ix = r * 20 + Int(c / 8) ' 20 bytes per row
  If ix > (UBound(parr, 2)) Or ix < 0 Then Exit Function
'  c = 2^(7 - (c MOD 8))
  c = 2 ^ (c Mod 8)
  va = parr(0, ix)
  If psiz > 0 Then
    If (Int(va / c) Mod 2) = 0 Then
      If pbit < 0 Then
        dmx_place = True
        Exit Function
      End If
      parr(0, ix) = va + c
    Else
      Exit Function
    End If
  End If
  If pbit > 0 Then
    va = parr(1, ix)
    If (Int(va / c) Mod 2) = 0 Then va = va + c ' else va = va - c
    parr(1, ix) = va
  End If
  dmx_place = True
End Function

Function dmx_placebyte(parr As Variant, psiz As Integer, pbl As Integer, _
                       ByRef prow As Variant, ByRef pcol As Variant, pbyte As Integer) As Boolean
  Dim bity(7) As Integer
  Dim xv As Boolean
  Dim i, x%
  x = pbyte
  For i = 7 To 0 Step -1
    bity(i) = x Mod 2
    x = Int(x / 2)
    If Not (dmx_place(parr, psiz, pbl, (prow(i)), (pcol(i)), -1)) Then
      dmx_placebyte = False
      Exit Function
    End If
  Next
  For i = 0 To 7
    xv = dmx_place(parr, psiz, pbl, (prow(i)), (pcol(i)), bity(i))
  Next
  dmx_placebyte = True
End Function

Function dmx_can_put(parr As Variant, psiz As Integer, pbl As Integer, _
                     prow As Integer, pcol As Integer, pbyte As Integer, _
                     pcorner As Boolean) As Boolean
  Dim dmxtype As Integer
  Dim wr As Variant
  Dim wc As Variant
  dmxtype = 0
  dmx_can_put = False
  If pcorner Then
    If prow = psiz And pcol = 0 Then
      dmxtype = 1   ' LowerLeft
    ElseIf prow = (psiz - 2) And pcol = 0 And (psiz Mod 4) <> 0 Then
      dmxtype = 2   ' lower left 2
    ElseIf prow = (psiz - 2) And pcol = 0 And (psiz Mod 8) = 4 Then
      dmxtype = 3   ' lower left 3
    ElseIf prow = (psiz + 4) And pcol = 2 And (psiz Mod 8) = 0 Then
      dmxtype = 4   ' lower right
    End If
    If dmxtype = 0 Then Exit Function
  End If
  Select Case dmxtype
    Case 1
      wr = Array(psiz - 1, psiz - 1, psiz - 1, 0, 0, 1, 2, 3)
      wc = Array(0, 1, 2, psiz - 2, psiz - 1, psiz - 1, psiz - 1, psiz - 1)
    Case 2
      wr = Array(psiz - 3, psiz - 2, psiz - 1, 0, 0, 0, 0, 1)
      wc = Array(0, 0, 0, psiz - 4, psiz - 3, psiz - 2, psiz - 1, psiz - 1)
     Case 3
      wr = Array(psiz - 3, psiz - 2, psiz - 1, 0, 0, 1, 2, 3)
      wc = Array(0, 0, 0, psiz - 2, psiz - 1, psiz - 1, psiz - 1, psiz - 1)
    Case 4
      wr = Array(psiz - 1, psiz - 1, 0, 0, 0, 1, 1, 1)
      wc = Array(0, psiz - 1, psiz - 3, psiz - 2, psiz - 1, psiz - 3, psiz - 2, psiz - 1)
    Case Else
      wr = Array(prow - 2, prow - 2, prow - 1, prow - 1, prow - 1, prow, prow, prow)
      wc = Array(pcol - 2, pcol - 1, pcol - 2, pcol - 1, pcol, pcol - 2, pcol - 1, pcol)
  End Select
  dmx_can_put = dmx_placebyte(parr, psiz, pbl, wr, wc, pbyte)
End Function ' dmx_can_put

'  Function exor(pa as integer, pb as integer) as INTEGER
'    Dim exorr as integer
'    Dim exorb as integer
'    exorr = 0 : exorb = 1
'    do while exorb <= pa or exorb <= pb
'      IF (int(pa / exorb) MOD 2) <> (int(pb / exorb) MOD 2) THEN exorr = exorr + exorb
'      exorb = exorb + exorb
'    loop
'    exor = exorr
'    exor = pa XOR pb
'  END Function ' exor

Sub dmx_rs(ppoly As Integer, pmemptr As Variant, ByVal psize As Integer, ByVal plen As Integer, ByVal pblocks As Integer)
    Dim v_x%, v_y%, v_z%, v_a%, v_b%, pa%, pb%, rp%
    Dim poly(512) As Integer
    Dim v_ply() As Integer
    
    ' generate reed solomon expTable and logTable
    '   for datamatrix GF256(0x012D) // 0x12d=301 => x^8 + x^5 + x^3 + x^2 + 1
    '   QR uses GF256(0x11d) // 0x11d=285 => x^8 + x^4 + x^3 + x^2 + 1
    v_x = 1: v_y = 0
    Do
      poly(v_x) = v_y   ' expTable
      poly(v_y + 256) = v_x ' logTable
      If v_x = 0 Then Exit Do
      v_x = v_x * 2
      v_y = v_y + 1
      If v_x > 255 Then v_x = v_x Xor ppoly
      If v_x = 1 Then v_x = 0
    Loop
' for 301 check:
' 255,0,1,240,2,225,241,53,3,38,226,133,242,43,54,210,4,195,39,114,227,106,134,28,243,140,44,23,55,118,211,234,5,219,196,96,40,222,115,103,228,78,107,125,135,8,29,162,244,186,141,180,45,99,24,49,56,13,119,153,212,199,235,91,6,76,220,217,197,11,97,184,41,36,223,253,116,138,104,193,229,86,79,171,108,165,126,145,136,34,9,74,30,32,163,84,245,173,187,204,142,81,181,190,46,88,100,159,25,231,50,207,57,147,14,67,120,128,154,248,213,167,200,63,236,110,92,176,7,161,77,124,221,102,218,95,198,90,12,152,98,48,185,179,42,209,37,132,224,52,254,239,117,233,139,22,105,27,194,113,230,206,87,158,80,189,172,203,109,175,166,62,127,247,146,66,137,192,35,252,10,183,75,216,31,83,33,73,164,144,85,170,246,65,174,61,188,202,205,157,143,169,82,72,182,215,191,251,47,178,89,151,101,94,160,123,26,112,232,21,51,238,208,131,58,69,148,18,15,16,68,17,121,149,129,19,155,59,249,70,214,250,168,71,201,156,64,60,237,130,111,20,93,122,177,150
' 1,2,4,8,16,32,64,128,45,90,180,69,138,57,114,228,229,231,227,235,251,219,155,27,54,108,216,157,23,46,92,184,93,186,89,178,73,146,9,18,36,72,144,13,26,52,104,208,141,55,110,220,149,7,14,28,56,112,224,237,247,195,171,123,246,193,175,115,230,225,239,243,203,187,91,182,65,130,41,82,164,101,202,185,95,190,81,162,105,210,137,63,126,252,213,135,35,70,140,53,106,212,133,39,78,156,21,42,84,168,125,250,217,159,19,38,76,152,29,58,116,232,253,215,131,43,86,172,117,234,249,223,147,11,22,44,88,176,77,154,25,50,100,200,189,87,174,113,226,233,255,211,139,59,118,236,245,199,163,107,214,129,47,94,188,85,170,121,242,201,191,83,166,97,194,169,127,254,209,143,51,102,204,181,71,142,49,98,196,165,103,206,177,79,158,17,34,68,136,61,122,244,197,167,99,198,161,111,222,145,15,30,60,120,240,205,183,67,134,33,66,132,37,74,148,5,10,20,40,80,160,109,218,153,31,62,124,248,221,151,3,6,12,24,48,96,192,173,119,238,241,207,179,75,150,0
    ReDim v_ply(plen + pblocks)
    For v_x = 1 To plen + 1
      pmemptr(v_x + psize) = 0
    Next
    For v_b = 0 To (pblocks - 1)
      v_ply(v_b + 1) = 1
      v_z = 1
      v_x = v_b + 1 + pblocks
      Do While v_x <= plen + pblocks
        v_ply(v_x) = v_ply(v_x - pblocks)
        v_y = v_x - pblocks
        Do While v_y >= v_b + 1 + pblocks
          pa = v_ply(v_y): pb = poly(256 + v_z): GoSub rsprod
          v_ply(v_y) = v_ply(v_y - pblocks) Xor rp
          v_y = v_y - pblocks
        Loop
        pa = v_ply(v_b + 1): pb = poly(256 + v_z): GoSub rsprod
        v_ply(v_b + 1) = rp
        v_z = v_z + 1
        v_x = v_x + pblocks
      Loop
      ' generate "nc" checkwords in the array
      v_x = v_b + 1
      Do While v_x <= psize
        v_y = v_b + 1
        v_z = pmemptr(v_y + psize) Xor pmemptr(v_x)
        v_a = plen - pblocks + 1 + v_b ' pro pblocks = 1 je to plen ; pro blocks = 2 to musiÂ­byt plen - pblocks + p_b + 1
        Do While v_y <= plen
          pa = v_z: pb = v_ply(v_a): GoSub rsprod
          pmemptr(v_y + psize) = pmemptr(v_y + psize + pblocks) Xor rp
          v_y = v_y + pblocks
          v_a = v_a - pblocks
        Loop
        v_x = v_x + pblocks
      Loop
    Next
    Exit Sub
rsprod:
    rp = 0
    If pa > 0 And pb > 0 Then rp = poly(256 + (poly(pa) + poly(pb)) Mod 255)
    Return
End Sub ' reed solomon dmx_rs

'  Sub setMousePointer(oWin, bEnable As Boolean)
'    Dim oPointer, iPoint%
'    If bEnable Then
'      iPoint = com.sun.star.awt.SystemPointer.ARROW
'    Else
'      iPoint = com.sun.star.awt.SystemPointer.Wait
'    End If
'    oPointer = createUnoService("com.sun.star.awt.Pointer")
'    oPointer.setType (iPoint)
'    oWin.setPointer (oPointer)
'  End Sub

Function dmx_gen(ptext As String, poptions As String) As String
    Dim encoded1(2200) As Integer
    Dim encoded2(3300) As Integer
    Dim encoded3(3300) As Integer
    Dim encix(3) As Integer
    Dim SavedPointer As Variant
    Dim enctype%, dmx_row%, dmx_col%
    Dim i&, j&, k&
    Dim ch%, bl%, s%, siz%
    Dim ascimatrix As String
    Dim err As String
    Dim arr() As Integer
    Dim x As Boolean
    dmx_row = 0
    ascimatrix = ""
    err = ""
    dmx_gen = ""
'    setMousePointer(wnd,False)
    If ptext = "" Then
      err = "Not data"
      Exit Function
    End If
    encix(1) = 0
    If (InStr(poptions, "ASCII") > 0) Then
      enctype = 1
      encix(2) = -1
      encix(3) = -1
    Else
      enctype = 0
      encix(2) = 0
      encix(3) = 0
    End If
    For i = 1 To Len(ptext)
      ch = AscL(Mid(ptext, i, 1)) Mod 256
      j = -5
      ' F1=_1 _=__ : ASCII=chr(232) C40&Text=chr(1) + chr(27)
      If i < Len(ptext) Then
        k = AscL(Mid(ptext, i + 1, 1)) Mod 256
        If Mid(ptext, i, 2) = "_1" Then
          j = -1
          i = i + 1
        ElseIf ch >= 48 And ch <= 57 And k >= 48 And k <= 57 Then
          j = val(Mid(ptext, i, 2))
        ElseIf Mid(ptext, i, 2) = "__" Then ' podtrzitko
          i = i + 1
          j = -5
        End If
      End If
      ' ascii encoding
      If encix(1) >= 0 And dmx_row = 0 Then
        If j = -1 Then
          encix(1) = encix(1) + 1
          encoded1(encix(1)) = 232
        Else
          If (ch >= 128) Then
            encix(1) = encix(1) + 1
            encoded1(encix(1)) = 235 ' hi bit
            ch = ch - 128
          End If
          If j >= 0 Then
            encix(1) = encix(1) + 1
            encoded1(encix(1)) = j + 130
            dmx_row = 1 ' SKIP NEXT ASCII
          Else
            encix(1) = encix(1) + 1
            encoded1(encix(1)) = ch + 1
          End If
        End If
      Else
        dmx_row = 0 ' no skip next ASCII
      End If
      ' C40 encoding
      If encix(2) >= 0 Then
        ' chr(230) Start C40
        If j = -1 Then ' FNC 1
          encix(2) = encix(2) + 2
          encoded2(encix(2) - 1) = 1  ' set 2
          encoded2(encix(2)) = 27 ' set2 FNC1
        Else
          If ch > 128 Then
            encix(2) = encix(2) + 2
            encoded2(encix(2) - 1) = 1
            encoded2(encix(2)) = 30  ' set2 hi-bit
            ch = ch - 128
          End If
          If ch < 32 Then
            encix(2) = encix(2) + 2
            encoded2(encix(2) - 1) = 0
            encoded2(encix(2)) = ch ' set1 control
          Else
            k = InStr(" 0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(ch))
            If ch <= 90 And k > 0 Then
              encix(2) = encix(2) + 1
              encoded2(encix(2)) = k + 2 ' set0 default
            Else
              k = InStr("!""#$%&'()*+,-./:;<=>?@(\)^_", Chr(ch))
              If (k > 0) Then
                encix(2) = encix(2) + 2
                encoded2(encix(2) - 1) = 1
                encoded2(encix(2)) = k - 1 ' set 2
              Else
                k = InStr("`abcdefghijklmnopqrstuvwxyz{|}~", Chr(ch))
                If (k > 0) Then
                  encix(2) = encix(2) + 2
                  encoded2(encix(2) - 1) = 2 ' set 3
                  encoded2(encix(2)) = k - 1
                Else
                  encix(2) = -1
                End If
              End If
            End If
          End If
        End If
        If encix(2) > UBound(encoded2) - 10 Then encix(2) = -1
      End If
      ' Text encoding
      If encix(3) >= 0 Then
        ' chr(239) start Text
        If j = -1 Then ' FNC 1
          encix(3) = encix(3) + 2
          encoded3(encix(3) - 1) = 1
          encoded3(encix(3)) = 27
        Else
          If ch > 128 Then
            encix(3) = encix(3) + 2
            encoded3(encix(3) - 1) = 1
            encoded3(encix(3)) = 30  ' set2 hi-bit
            ch = ch - 128
          End If
          If (ch < 32) Then
            encix(3) = encix(3) + 2
            encoded3(encix(3) - 1) = 0
            encoded3(encix(3)) = ch ' set1 control
          Else
            k = InStr(" 0123456789abcdefghijklmnopqrstuvwxyz", Chr(ch))
            If (ch < 65 Or ch >= 97) And k > 0 Then
              encix(3) = encix(3) + 1
              encoded3(encix(3)) = k + 2 ' set0
            Else
              k = InStr("!""#$%&'()*+,-./:;<=>?@(\)^_", Chr(ch))
              If (k > 0) Then
                encix(3) = encix(3) + 2
                encoded3(encix(3) - 1) = 1
                encoded3(encix(3)) = k - 1 ' set2
              Else
                k = InStr("`ABCDEFGHIJKLMNOPQRSTUVWXYZ{|}~", ch)
                If (k > 0) Then
                  encix(3) = encix(3) + 2
                  encoded3(encix(3) - 1) = 2
                  encoded3(encix(3)) = k - 1 ' set3
                Else
                  encix(3) = -1
                End If
              End If
            End If
          End If
        End If
        If encix(3) > UBound(encoded3) - 10 Then encix(3) = -1
      End If
      If encix(1) > UBound(encoded1) - 20 Then
         err = "Too long text"
'         setMousePointer(wnd,True)
         Exit Function
      End If
    Next i
    i = encix(1): j = 10000: k = 10000
    If enctype = 0 Then
      If encix(1) <= 0 Then i = 10000
      If encix(2) > 0 Then j = 2 + Int((encix(2) * 2) / 3)
      If encix(3) > 0 Then k = 2 + Int((encix(3) * 2) / 3)
      If i < j And i < k Then
        enctype = 1
      ElseIf j < i And j < k Then
        enctype = 2
      ElseIf k < i And k < j Then
        enctype = 3
      End If
    End If
    ' string not convertible
    If encix(enctype) <= 0 Then
      err = "Bad chars"
'      setMousePointer(wnd,True)
      Exit Function
    End If
    If enctype > 1 Then
      i = 1
      k = 1
      dmx_row = encix(enctype) + 1
      If enctype = 2 Then
        encoded1(1) = 230 ' start enc
        encoded2(dmx_row) = 0 ' padding
        encoded2(dmx_row + 1) = 0 ' padding
      Else
        encoded1(1) = 239 ' start enc
        encoded3(dmx_row) = 0 ' padding
        encoded3(dmx_row + 1) = 0 ' padding
      End If
      Do While i <= dmx_row:
        If enctype = 2 Then
          j = 1600& * encoded2(i) + 40& * encoded2(i + 1) + 1& + encoded2(i + 2)
        Else
          j = 1600& * encoded3(i) + 40& * encoded3(i + 1) + 1& + encoded3(i + 2)
        End If
        i = i + 3
        k = k + 2
        encoded1(k - 1) = Int(j / 256)
        encoded1(k) = j Mod 256
      Loop
      j = 254& + (129& * 256&) ' padding 254,129 = switch to text, end of message
      encix(1) = k
    Else
      j = 129 ' END = 129
'      j = 129& + (251& * 256&) + (147& * 65536)  ' END = 129 + padding 251,147 - test Wikipedia
    End If
    i = encix(1)
    k = 0: ch = 1
    ' only sqare types implemented
    Do
      Select Case i
        Case 3: siz = 8: k = 5
        Case 5: siz = 10: k = 7
        Case 8: siz = 12: k = 10
        Case 12: siz = 14: k = 12
        Case 18: siz = 16: k = 14
        Case 22: siz = 18: k = 18
        Case 30: siz = 20: k = 20
        Case 36: siz = 22: k = 24
        Case 44: siz = 24: k = 28
        Case 62: siz = 28: k = 36
        Case 86: siz = 32: k = 42
        Case 114: siz = 36: k = 48
        Case 144: siz = 40: k = 56
        Case 174: siz = 44: k = 68
        Case 204: siz = 48: k = 84: ch = 2
        Case 280: siz = 56: k = 112: ch = 2
        Case 368: siz = 64: k = 144: ch = 4
        Case 456: siz = 72: k = 192: ch = 4
        Case 576: siz = 80: k = 224: ch = 4
        Case 696: siz = 88: k = 272: ch = 4
        Case 816: siz = 96: k = 336: ch = 6
        Case 1050: siz = 108: k = 408: ch = 6
        Case 1304: siz = 120: k = 496: ch = 8
        Case 1558: siz = 132: k = 620: ch = 10
      End Select
      If k > 4 Or i > 1558 Then Exit Do
      i = i + 1
      If j = 0 Then encoded1(i) = ((149& * i) Mod 255) + 1 Else encoded1(i) = j Mod 256
      If j <> 0 Then j = Int(j / 256&)
    Loop
    If (i > 1558) Then
      err = "Too big code"
'      setMousePointer(wnd,True)
      Exit Function
    End If
    If siz >= 108 Then
      bl = 6
    ElseIf siz >= 56 Then
      bl = 4
    ElseIf siz >= 28 Then
      bl = 2
    Else
      bl = 1
    End If
    ReDim arr(0)
    ReDim arr(1, siz * 20& + 40& * (bl + 1)) ' 20 bytes per row
    arr(0, 0) = 128
    ' doplnime ECC
    Call dmx_rs(301, encoded1, 0 + i, 0 + k, ch)
'' Call arr2hexstr(encoded1)
    encix(1) = i + k
    dmx_row = 4: dmx_col = 0: i = 1
    Do
      ' only corners cases
      If dmx_can_put(arr, siz, bl, dmx_row, dmx_col, encoded1(i), True) Then i = i + 1
      Do
        If (dmx_row < siz) And (dmx_col >= 0) Then
          If dmx_can_put(arr, siz, bl, dmx_row, dmx_col, encoded1(i), False) Then i = i + 1
        End If
        dmx_row = dmx_row - 2
        dmx_col = dmx_col + 2
        If dmx_row < 0 Or dmx_col >= siz Or i > encix(1) Then Exit Do
      Loop
      dmx_row = dmx_row + 1
      dmx_col = dmx_col + 3
      Do ' downward diagonaly
        If (dmx_row >= 0) And (dmx_col < siz) Then
          If dmx_can_put(arr, siz, bl, dmx_row, dmx_col, encoded1(i), False) Then i = i + 1
        End If
        dmx_row = dmx_row + 2
        dmx_col = dmx_col - 2
        If dmx_row >= siz Or dmx_col < 0 Or i > encix(1) Then Exit Do
      Loop
      dmx_row = dmx_row + 3
      dmx_col = dmx_col + 1
      If (dmx_row >= siz And dmx_col >= siz) Or i > encix(1) Then Exit Do
    Loop
    k = siz * siz
    If (k Mod 8) = 4 Then ' right lower void
      x = dmx_place(arr, siz, bl, siz - 1, siz - 1, 1)
      x = dmx_place(arr, siz, bl, siz - 2, siz - 2, 1)
    End If
    s = Int(siz / bl)
    For i = -1 To s
      For k = 0 To bl ^ 2 - 1
        dmx_col = (k Mod bl) * (s + 2)
        dmx_row = Int(k / bl) * (s + 2)
        x = dmx_place(arr, 0, 0, dmx_row + i, dmx_col - 1, 1) ' leva cara
        x = dmx_place(arr, 0, 0, dmx_row + s, dmx_col + i, 1) ' spodni cara
        If ((i + 2) Mod 2) = 1 Then
          x = dmx_place(arr, 0, 0, dmx_row - 1, dmx_col + i, 1) ' horni tecky
        Else
          x = dmx_place(arr, 0, 0, dmx_row + i, dmx_col + s, 1) ' prave tecky
        End If
      Next
    Next
'    ascimatrix = trim(CStr(siz + 2))
'    ascimatrix = ascimatrix & "x" & ascimatrix & ","
    ascimatrix = ""
    k = siz + 2 * (bl + 1) - 1
    For dmx_row = 0 To k Step 2
      s = 0
      For dmx_col = 0 To k Step 2
        If (dmx_col Mod 8) = 0 Then
          ch = arr(1, s + 20 * dmx_row)
          i = arr(1, s + 20 * (dmx_row + 1))
          s = s + 1
        End If
        ascimatrix = ascimatrix _
           & Chr(97 + (ch Mod 4) + 4 * (i Mod 4))
        ch = Int(ch / 4)
        i = Int(i / 4)
      Next
      ascimatrix = ascimatrix & vbNewLine
    Next dmx_row
    ReDim arr(1, 1)
    dmx_gen = ascimatrix
End Function  ' dmx_gen

Sub qr_rs(ppoly As Integer, pmemptr As Variant, ByVal psize As Integer, ByVal plen As Integer, ByVal pblocks As Integer)
    Dim v_x%, v_y%, v_z%, v_a%, v_b%, pa%, pb%, rp%, v_last%, v_bs%, v_b2c%, vpo%, vdo%, v_es%
    Dim poly(512) As Byte
    Dim v_ply() As Byte
'    Dim dbg$
    ' generate reed solomon expTable and logTable
    '   QR uses GF256(0x11d) // 0x11d=285 => x^8 + x^4 + x^3 + x^2 + 1
    v_x = 1: v_y = 0
    For v_y = 0 To 255
      poly(v_y) = v_x         ' expTable
      poly(v_x + 256) = v_y   ' logTable
      v_x = v_x * 2
      If v_x > 255 Then v_x = v_x Xor ppoly
    Next
'    poly(257) =    ' pro QR logTable(1) = 0 not50
'Call arr2decstr(poly)
    For v_x = 1 To plen
      pmemptr(v_x + psize) = 0
    Next
    v_b2c = pblocks
    ' qr code has first x blocks shorter than lasts
    v_bs = Int(psize / pblocks) ' shorter block size
    v_es = Int(plen / pblocks) ' ecc block size
    v_x = psize Mod pblocks ' remain bytes
    v_b2c = pblocks - v_x ' on block number v_b2c
    ReDim v_ply(v_es + 1)
    v_z = 0 ' pro QR je v_z=0 pro dmx je v_z=1
    v_ply(1) = 1
    v_x = 2
    Do While v_x <= v_es + 1
      v_ply(v_x) = v_ply(v_x - 1)
      v_y = v_x - 1
      Do While v_y > 1
        pb = poly(v_z)
        pa = v_ply(v_y): GoSub rsprod
        v_ply(v_y) = v_ply(v_y - 1) Xor rp
        v_y = v_y - 1
      Loop
      pa = v_ply(1): pb = poly(v_z): GoSub rsprod
      v_ply(1) = rp
      v_z = v_z + 1
      v_x = v_x + 1
    Loop
'Call arr2hexstr(v_ply)
    For v_b = 0 To (pblocks - 1)
      vpo = v_b * v_es + 1 + psize ' ECC start
      vdo = v_b * v_bs + 1 ' data start
      If v_b > v_b2c Then vdo = vdo + v_b - v_b2c ' x longers before
      ' generate "nc" checkwords in the array
      v_x = 0
      v_z = v_bs
      If v_b >= v_b2c Then v_z = v_z + 1
      Do While v_x < v_z
        pa = pmemptr(vpo) Xor pmemptr(vdo + v_x)
        v_y = vpo
        v_a = v_es
        Do While v_a > 0
          pb = v_ply(v_a): GoSub rsprod
          If v_a = 1 Then
            pmemptr(v_y) = rp
          Else
            pmemptr(v_y) = pmemptr(v_y + 1) Xor rp
          End If
          v_y = v_y + 1
          v_a = v_a - 1
        Loop
        v_x = v_x + 1
'if v_b = 0 and v_x = v_z then call arr2hexstr(pmemptr)
      Loop
    Next
    Exit Sub
rsprod:
    rp = 0
    If pa > 0 And pb > 0 Then rp = poly((0& + poly(256 + pa) + poly(256 + pb)) Mod 255&)
    Return
End Sub ' reed solomon qr_rs

Sub bb_putbits(ByRef parr As Variant, ByRef ppos As Integer, pa As Variant, ByVal plen As Integer)
  Dim i%, b%, w&, l%, j%
  Dim dw As Double
  Dim x(7) As Byte
  Dim y As Variant
  w = VarType(pa)
  If w = 17 Or w = 2 Or w = 3 Or w = 5 Then ' byte,integer,long, double
    If plen > 56 Then Exit Sub
    dw = pa
    l = plen
    If l < 56 Then dw = dw * 2 ^ (56 - l)
    i = 0
    Do While i < 6 And dw > 0#
      w = Int(dw / 2 ^ 48)
      x(i) = w Mod 256
      dw = dw - 2 ^ 48 * w
      dw = dw * 256
      l = l - 8
      i = i + 1
    Loop
    y = x
  ElseIf InStr("Integer(),Byte(),Long(),Variant()", TypeName(pa)) > 0 Then
    y = pa
  Else
    MsgBox TypeName(pa), "Unknown type"
    Exit Sub
  End If
  i = Int(ppos / 8) + 1
  b = ppos Mod 8
  j = LBound(y)
  l = plen
  Do While l > 0
    If j <= UBound(y) Then
      w = y(j)
      j = j + 1
    Else
      w = 0
    End If
    If (l < 8) Then w = w And (256 - 2 ^ (8 - l))
    If b > 0 Then
      w = w * 2 ^ (8 - b)
      parr(i) = parr(i) Or Int(w / 256)
      parr(i + 1) = parr(i + 1) Or (w And 255)
    Else
      parr(i) = parr(i) Or (w And 255)
    End If
    If l < 8 Then
      ppos = ppos + l
      l = 0
    Else
      ppos = ppos + 8
      i = i + 1
      l = l - 8
    End If
  Loop
End Sub

Function qr_numbits(ByVal num As Long) As Integer
  Dim n%, a&
  a = 1: n = 0
  Do While a <= num
    a = a * 2
    n = n + 1
  Loop
  qr_numbits = n
End Function

' padding 0xEC,0x11,0xEC,0x11...
' TYPE_INFO_MASK_PATTERN = 0x5412
' TYPE_INFO_POLY = 0x537  [(ecLevel << 3) | maskPattern] : 5 + 10 = 15 bitu
' VERSION_INFO_POLY = 0x1f25 : 5 + 12 = 17 bitu
Sub qr_bch_calc(ByRef data As Long, ByVal poly As Long)
  Dim b%, n%, rv&, x&
  b = qr_numbits(poly) - 1
  If data = 0 Then
'    data = poly
    Exit Sub
  End If
  x = data * 2 ^ b
  rv = x
  Do
    n = qr_numbits(rv)
    If n <= b Then Exit Do
    rv = rv Xor (poly * 2 ^ (n - b - 1))
  Loop
  data = x + rv
End Sub

Sub qr_params(ByVal pcap As Long, ByVal ecl As Integer, ByRef rv As Variant, ByRef ecx_poc As Variant)
  Dim siz%, totby&, s$, i&, syncs%, ccsiz%, ccblks%, j&, ver%
'  Dim rv(15) as Integer ' 1:version,2:size,3:ccs,4:ccb,5:totby,6-12:syncs(7),13-15:versinfo(3)
'  ecl:M=0,L=1,H=2,Q=3
  If ecl < 0 Or ecl > 3 Then Exit Sub
  For i = 1 To UBound(rv): rv(i) = 0: Next i
  j = Int((pcap + 18 * ecx_poc(1) + 17 * ecx_poc(2) + 20 * ecx_poc(3) + 7) / 8)
  If ecl = 0 And j > 2334 Or _
     ecl = 1 And j > 2956 Or _
     ecl = 2 And j > 1276 Or _
     ecl = 3 And j > 1666 Then
    Exit Sub
  End If
  j = Int((pcap + 14 * ecx_poc(1) + 13 * ecx_poc(2) + 12 * ecx_poc(3) + 7) / 8)
  For ver = 1 To 40
    If ver = 10 Then j = Int((pcap + 16 * ecx_poc(1) + 15 * ecx_poc(2) + 20 * ecx_poc(3) + 7) / 8)
    If ver = 27 Then j = Int((pcap + 18 * ecx_poc(1) + 17 * ecx_poc(2) + 20 * ecx_poc(3) + 7) / 8)
    siz = 4 * ver + 17
    i = (ver - 1) * 12 + ecl * 3
    s = Mid("D01A01K01G01J01D01V01P01T01I01P02L02L02N01J04T02R02T01P04L04J04L02V04R04L04N02T05L06P04R02T06P06P05X02R08N08T05L04V08R08X05N04R11V08P08R04V11T10P09T04P16R12R09X04R16N16R10P06R18X12V10R06X16R17V11V06V19V16T13X06V21V18T14V07T25T21T16V08V25X20T17V08X25V23V17V09R34X23V18X09X30X25V20X10X32X27V21T12X35X29V23V12X37V34V25X12X40X34V26X13X42X35V28X14X45X38V29X15X48X40V31X16X51X43V33X17X54X45V35X18X57X48V37X19X60X51V38X19X63X53V40X20X66X56V43X21X70X59V45X22X74X62V47X24X77X65V49X25X81X68" _
            , i + 1, 3)
    ccsiz = AscL(Left(s, 1)) - 65 + 7
    ccblks = val(right(s, 2))
    If ver = 1 Then
      syncs = 0
      totby = 26
    Else
      syncs = ((Int(ver / 7) + 2) ^ 2) - 3
      totby = siz - 1
      totby = ((totby ^ 2) / 8) - (3& * syncs) - 24
      If ver > 6 Then totby = totby - 4
      If syncs = 1 Then totby = totby - 1
    End If
'MsgBox "ver:" & ver & " tot: " & totby & " dat:" & (totby - ccsiz * ccblks) & " need:" & j
    If totby - ccsiz * ccblks >= j Then Exit For
  Next
  If ver > 1 Then
    syncs = Int(ver / 7) + 2
    rv(6) = 6
    rv(5 + syncs) = siz - 7
    If syncs > 2 Then
      i = Int((siz - 13) / 2 / (syncs - 1) + 0.7) * 2
      rv(7) = rv(5 + syncs) - i * (syncs - 2)
      If syncs > 3 Then
        For j = 3 To syncs - 1
          rv(5 + j) = rv(4 + j) + i
        Next
      End If
    End If
  End If
  rv(1) = ver
  rv(2) = siz
  rv(3) = ccsiz: rv(4) = ccblks
  rv(5) = totby
  If ver >= 7 Then
    i = ver
    Call qr_bch_calc(i, &H1F25)
    rv(13) = Int(i / 65536)
    rv(14) = Int(i / 256&) Mod 256
    rv(15) = i Mod 256
  End If
End Sub

Function qr_bit(parr As Variant, ByVal psiz As Integer, _
                ByVal prow As Integer, ByVal pcol As Integer, _
                ByVal pbit As Integer) As Boolean
  Dim ix%, va%, r%, c%, s%
  r = prow
  c = pcol
  qr_bit = False
  ix = r * 24 + Int(c / 8) ' 24 bytes per row
  If ix > (UBound(parr, 2)) Or ix < 0 Then Exit Function
  c = 2 ^ (c Mod 8)
  va = parr(0, ix)
  If psiz > 0 Then ' Kontrola masky
    If (va And c) = 0 Then
      If pbit <> 0 Then parr(1, ix) = parr(1, ix) Or c
      qr_bit = True
    Else
      qr_bit = False
    End If
  Else
    qr_bit = True
    parr(1, ix) = parr(1, ix) And (255 - c) ' reset bit for psiz <= 0
    If pbit > 0 Then parr(1, ix) = parr(1, ix) Or c
    If psiz < 0 Then parr(0, ix) = parr(0, ix) Or c ' mask for psiz < 0
  End If
End Function

Sub qr_mask(parr As Variant, pb As Variant, ByVal pbits As Integer, ByVal pr As Integer, ByVal pc As Integer)
' max 8 bites wide
  Dim i%, w&, r%, c%, j%
  Dim x As Boolean
  If pbits > 8 Or pbits < 1 Then Exit Sub
  r = pr: c = pc
  w = VarType(pb)
  If w = 17 Or w = 2 Or w = 3 Or w = 5 Then ' byte,integer,long, double
    w = Int(pb)
    i = 2 ^ (pbits - 1)
    Do While i > 0
      x = qr_bit(parr, -1, r, c, w And i)
      c = c + 1
      i = Int(i / 2)
    Loop
  ElseIf InStr("Integer(),Byte(),Long(),Variant()", TypeName(pb)) > 0 Then
    For j = LBound(pb) To UBound(pb)
      w = Int(pb(j))
      i = 2 ^ (pbits - 1)
      c = pc
      Do While i > 0
        x = qr_bit(parr, -1, r, c, w And i)
        c = c + 1
        i = Int(i / 2)
      Loop
      r = r + 1
    Next
  End If
End Sub

Sub qr_fill(parr As Variant, ByVal psiz%, pb As Variant, ByVal pblocks As Integer, ByVal pdlen As Integer, ByVal ptlen As Integer)
  ' vyplni pole parr (psiz x 24 bytes) z pole pb pdlen = pocet dbytes, pblocks = bloku, ptlen celkem
  ' podle logiky qr_kodu - s prokladem
  Dim vx%, vb%, vy%, vdnlen%, vds%, ves%, c%, r%, wa%, wb%, w%, smer%, vsb%
  ' qr code has first x blocks shorter than lasts but datamatrix has first longer and shorter last
  vds = Int(pdlen / pblocks) ' shorter data block size
  ves = Int((ptlen - pdlen) / pblocks) ' ecc block size
  vdnlen = vds * pblocks ' potud jsou databloky stejne velike
  vsb = pblocks - (pdlen Mod pblocks) ' mensich databloku je ?
  
  c = psiz - 1: r = c ' start position on right lower corner
  smer = 0 ' nahoru :  3 <- 2 10  dolu: 1 <- 0  32
           '           1 <- 0 10        3 <- 2  32
  vb = 1: w = pb(1): vx = 0
  Do While c >= 0 And vb <= ptlen
    If qr_bit(parr, psiz, r, c, (w And 128)) Then
      vx = vx + 1
      If vx = 8 Then
        GoSub qrfnb ' first byte
        vx = 0
      Else
        w = (w * 2) Mod 256
      End If
    End If
    Select Case smer
      Case 0, 2 ' nahoru nebo dolu a jsem vpravo
        c = c - 1
        smer = smer + 1
      Case 1 ' nahoru a jsem vlevo
        If r = 0 Then ' nahoru uz to nejde
          c = c - 1
          If c = 6 And psiz >= 21 Then c = c - 1 ' preskoc sync na sloupci 6
          smer = 2 ' a jedeme dolu
        Else
          c = c + 1
          r = r - 1
          smer = 0 ' furt nahoru
        End If
      Case 3 ' dolu a jsem vlevo
        If r = (psiz - 1) Then ' dolu uz to nepude
          c = c - 1
          If c = 6 And psiz >= 21 Then c = c - 1 ' preskoc sync na sloupci 6
          smer = 0
        Else
          c = c + 1
          r = r + 1
          smer = 2
        End If
    End Select
  Loop
  Exit Sub
qrfnb:
  ' next byte
        ' plen = 14 pbl = 3   => 1x4 + 2x5 (v_b2c = 3 - 2 = 1; v_bs1 = 4)
        '     v_b = 0 => v_last = 0 + 4 * 3 - 2 = 10 => 1..12 by 3   1,4,7,10
        '     v_b = 1 => v_last = 1 + 4 * 3     = 13 => 2..13 by 3   2,5,8,11,13
        '     v_b = 2 => v_last = 2 + 4 * 3     = 14 => 3..14 by 3   3,6,9,12,14
        ' plen = 15 pbl = 3   => 3x5 (v_b2c = 3; v_bs1 = 5)
        '     v_b = 0 => v_last = 0 + 5 * 3 - 2 = 13 => 1..13 by 3   1,4,7,10,13
        '     v_b = 1 => v_last = 1 + 5 * 3 - 2 = 14 => 2..14 by 3   2,5,8,11,14
        '     v_b = 2 => v_last = 2 + 5 * 3 - 2 = 15 => 3..15 by 3   3,6,9,12,15
  If vb < pdlen Then ' Datovy byte
    wa = vb
    If vb >= vdnlen Then
      wa = wa + vsb
    End If
    wb = wa Mod pblocks
    wa = Int(wa / pblocks)
    If wb > vsb Then wa = wa + wb - vsb
'    If vb >= vdnlen Then MsgBox "D:" & (1 + vds * wb + wa)
    w = pb(1 + vds * wb + wa)
  ElseIf vb < ptlen Then ' ecc byte
    wa = vb - pdlen ' kolikaty ecc 0..x
    wb = wa Mod pblocks ' z bloku
    wa = Int(wa / pblocks) ' kolikaty
'    MsgBox "E:" & (1 + pdlen + ves * wb + wa)
    w = pb(1 + pdlen + ves * wb + wa)
  End If
  vb = vb + 1
  Return
End Sub

' Black If 0: (c+r) mod 2 = 0    4: ((r div 2) + (c div 3)) mod 2 = 0
'          1: r mod 2 = 0        5: (c*r) mod 2 + (c*r) mod 3 = 0
'          2: c mod 3 = 0        6: ((c*r) mod 2 + (c*r) mod 3) mod 2 = 0
'          3: (c+r) mod 3 = 0    7: ((c+r) mod 2 + (c*r) mod 3) mod 2 = 0
Function qr_xormask(parr As Variant, ByVal siz As Integer, ByVal pmod As Integer, ByVal final As Boolean) As Long
  Dim score&, bl&, rp&, rc&, c%, r%, m%, ix%, i%, w%
  Dim warr() As Byte
  Dim cols() As Long
  
  ReDim warr(siz * 24)
  For r = 0 To siz - 1
    m = 1
    ix = 24 * r
    warr(ix) = parr(1, ix)
    For c = 0 To siz - 1
      If (parr(0, ix) And m) = 0 Then ' nemaskovany
        Select Case pmod
         Case 0: i = (c + r) Mod 2
         Case 1: i = r Mod 2
         Case 2: i = c Mod 3
         Case 3: i = (c + r) Mod 3
         Case 4: i = (Int(r / 2) + Int(c / 3)) Mod 2
         Case 5: i = (c * r) Mod 2 + (c * r) Mod 3
         Case 6: i = ((c * r) Mod 2 + (c * r) Mod 3) Mod 2
         Case 7: i = ((c + r) Mod 2 + (c * r) Mod 3) Mod 2
        End Select
        If i = 0 Then warr(ix) = warr(ix) Xor m
      End If
      If m = 128 Then
        m = 1
        If final Then parr(1, ix) = warr(ix)
        ix = ix + 1
        warr(ix) = parr(1, ix)
      Else
        m = m * 2
      End If
    Next c
    If m <> 128 And final Then parr(1, ix) = warr(ix)
  Next r
  If final Then
    qr_xormask = 0
    Exit Function
  End If
 ' score computing
 ' a) adjacent modules colors in row or column 5+i mods = 3 + i penatly
 ' b) block same color MxN = 3*(M-1)*(N-1) penalty OR every 2x2 block penalty + 3
 ' c) 4:1:1:3:1:1 or 1:1:3:1:1:4 in row or column = 40 penalty rmks: 00001011101 or 10111010000 = &H05D or &H5D0
 ' d) black/light ratio : k=(abs(ratio% - 50) DIV 5) means 10*k penalty
  score = 0: bl = 0
'Dim s(4) as Integer
  ReDim cols(1, siz)
  rp = 0: rc = 0
  For r = 0 To siz - 1
    m = 1
    ix = 24 * r
    rp = 0: rc = 0
    For c = 0 To siz - 1
      rp = (rp And &H3FF) * 2 ' only last 12 bits
      cols(1, c) = (cols(1, c) And &H3FF) * 2
      If (warr(ix) And m) <> 0 Then
        If rc < 0 Then ' in row x whites
          If rc <= -5 Then score = score - 2 - rc  ': s(0) = s(0) - 2 - rc
          rc = 0
        End If
        rc = rc + 1 ' one more black
        If cols(0, c) < 0 Then ' color changed
          If cols(0, c) <= -5 Then score = score - 2 - cols(0, c) ': s(1) = s(1) - 2 - cols(0,c)
          cols(0, c) = 0
        End If
        cols(0, c) = cols(0, c) + 1 ' one more black
        rp = rp Or 1
        cols(1, c) = cols(1, c) Or 1
        bl = bl + 1 ' balck modules count
      Else
        If rc > 0 Then ' in row x black
          If rc >= 5 Then score = score - 2 + rc ': s(0) = s(0) - 2 + rc
          rc = 0
        End If
        rc = rc - 1 ' one more white
        If cols(0, c) > 0 Then ' color changed
          If cols(0, c) >= 5 Then score = score - 2 + cols(0, c) ': s(1) = s(1) - 2 + cols(0,c)
          cols(0, c) = 0
        End If
        cols(0, c) = cols(0, c) - 1 ' one more white
      End If
      If c > 0 And r > 0 Then ' penalty block 2x2
        i = rp And 3 ' current row pair
        If (cols(1, c - 1) And 3) >= 2 Then i = i + 8
        If (cols(1, c) And 3) >= 2 Then i = i + 4
        If i = 0 Or i = 15 Then
          score = score + 3 ': s(2) = s(2) + 3
          ' b) penalty na 2x2 block same color
        End If
      End If
      If c >= 10 And (rp = &H5D Or rp = &H5D0) Then  ' penalty pattern c in row
        score = score + 40 ': s(3) = s(3) + 40
      End If
      If r >= 10 And (cols(1, c) = &H5D Or cols(1, c) = &H5D0) Then ' penalty pattern c in column
        score = score + 40 ': s(3) = s(3) + 40
      End If
      ' next mask / byte
      If m = 128 Then
        m = 1
        ix = ix + 1
      Else
        m = m * 2
      End If
    Next
    If rc <= -5 Then score = score - 2 - rc ': s(0) = s(0) - 2 - rc
    If rc >= 5 Then score = score - 2 + rc ': s(0) = s(0) - 2 + rc
  Next
  For c = 0 To siz - 1 ' after last row count column blocks
    If cols(0, c) <= -5 Then score = score - 2 - cols(0, c) ': s(1) = s(1) - 2 - cols(0,c)
    If cols(0, c) >= 5 Then score = score - 2 + cols(0, c) ': s(1) = s(1) - 2 + cols(0,c)
  Next
  bl = Int(Abs((bl * 100&) / (siz * siz) - 50&) / 5) * 10
'MsgBox "mask:" + pmod + " " + s(0) + "+" + s(1) + "+" + s(2) + "+" + s(3) + "+" + bl
  qr_xormask = score + bl
End Function

Function qr_gen(ptext As String, poptions As String) As String
  Dim encoded1() As Byte ' byte mode (ASCII) all max 3200 bytes
  Dim encix1%
  Dim ecx_cnt(3) As Integer
  Dim ecx_pos(3) As Integer
  Dim ecx_poc(3) As Integer
  Dim eb(20, 4) As Integer 'store how many characters should be in which ECI mode. This is a list of rows, each row corresponding a the next batch of characters with a different ECI mode.
  ' eb(i, 1) - ECI mode (1 = numeric, 2 = alphanumeric, 3 = byte)
  ' eb(i, 2) - last character in previous row
  ' eb(i, 3) - number of characters in THIS row
  ' eb(i, 4) - number of bits for THIS row
  Dim ascimatrix$, mode$, err$
  Dim ecl%, r%, c%, mask%, utf8%, ebcnt%
  Dim i&, j&, k&, m&
  Dim ch%, s%, siz%
  Dim x As Boolean
  Dim qrarr() As Byte ' final matrix
  Dim qrpos As Integer
  Dim qrp(15) As Integer     ' 1:version,2:size,3:ccs,4:ccb,5:totby,6-12:syncs(7),13-15:versinfo(3)
  Dim qrsync1(1 To 8) As Byte
  Dim qrsync2(1 To 5) As Byte

  ascimatrix = ""
  err = ""
  mode = "M"
  i = InStr(poptions, "mode=")
  If i > 0 Then mode = Mid(poptions, i + 5, 1)
' M=0,L=1,H=2,Q=3
  ecl = InStr("MLHQ", mode) - 1
  If ecl < 0 Then mode = "M": ecl = 0
  If ptext = "" Then
    err = "Not data"
    Exit Function
  End If
  For i = 1 To 3
    ecx_pos(i) = 0
    ecx_cnt(i) = 0
    ecx_poc(i) = 0
  Next i
  ebcnt = 1
  utf8 = 0
  For i = 1 To Len(ptext) + 1
    ' Decide how many bytes this character has
	If i > Len(ptext) Then
      k = -5 ' End of text --> skip several code sections 
    Else ' need to parse character i of ptext and decide how many bytes it has
      k = AscL(Mid(ptext, i, 1))
      If k >= &H1FFFFF Then ' FFFF - 1FFFFFFF
        m = 4
        k = -1
      ElseIf k >= &H7FF Then ' 7FF-FFFF 3 bytes
        m = 3
        k = -1
      ElseIf k >= 128 Then
        m = 2
        k = -1
      Else ' normal 7bit ASCII character, so it is worth it to check if it belong to the Numeric or Alphanumeric subsets defined in ECI (array qralnum)
        m = 1
        k = InStr(qralnum, Mid(ptext, i, 1)) - 1
      End If
    End If
	' Depending on k and a lot of other things, increase ebcnt
    If (k < 0) Then ' Treat mult-byte case or exit? (bude byte nebo konec)
      If ecx_cnt(1) >= 9 Or (k = -5 And ecx_cnt(1) = ecx_cnt(3)) Then ' Until now it was possible numeric??? (Az dosud bylo mozno pouzitelne numeric)
        If (ecx_cnt(2) - ecx_cnt(1)) >= 8 Or (ecx_cnt(3) = ecx_cnt(2)) Then ' pred num je i pouzitelny alnum
          If (ecx_cnt(3) > ecx_cnt(2)) Then ' Jeste pred alnum bylo byte
            eb(ebcnt, 1) = 3         ' Typ byte
            eb(ebcnt, 2) = ecx_pos(3) ' Position pozice
            eb(ebcnt, 3) = ecx_cnt(3) - ecx_cnt(2) ' delka
            ebcnt = ebcnt + 1
            ecx_poc(3) = ecx_poc(3) + 1
          End If
          eb(ebcnt, 1) = 2         ' Typ alnum
          eb(ebcnt, 2) = ecx_pos(2)
          eb(ebcnt, 3) = ecx_cnt(2) - ecx_cnt(1) ' delka
          ebcnt = ebcnt + 1
          ecx_poc(2) = ecx_poc(2) + 1
          ecx_cnt(2) = 0
        ElseIf ecx_cnt(3) > ecx_cnt(1) Then ' byly bytes pred numeric
          eb(ebcnt, 1) = 3         ' Typ byte
          eb(ebcnt, 2) = ecx_pos(3) ' Position pozice
          eb(ebcnt, 3) = ecx_cnt(3) - ecx_cnt(1) ' delka
          ebcnt = ebcnt + 1
          ecx_poc(3) = ecx_poc(3) + 1
        End If
      ElseIf (ecx_cnt(2) >= 8) Or (k = -5 And ecx_cnt(2) = ecx_cnt(3)) Then ' Az dosud bylo mozno pouzitelne alnum
        If (ecx_cnt(3) > ecx_cnt(2)) Then ' Jeste pred alnum bylo byte
          eb(ebcnt, 1) = 3         ' Typ byte
          eb(ebcnt, 2) = ecx_pos(3) ' Position pozice
          eb(ebcnt, 3) = ecx_cnt(3) - ecx_cnt(2) ' delka
          ebcnt = ebcnt + 1
          ecx_poc(3) = ecx_poc(3) + 1
        End If
        eb(ebcnt, 1) = 2         ' Typ alnum
        eb(ebcnt, 2) = ecx_pos(2)
        eb(ebcnt, 3) = ecx_cnt(2) ' delka
        ebcnt = ebcnt + 1
        ecx_poc(2) = ecx_poc(2) + 1
        ecx_cnt(3) = 0
        ecx_cnt(2) = 0 ' vse zpracovano
      ElseIf (k = -5 And ecx_cnt(3) > 0) Then ' konec ale mam co ulozit
        eb(ebcnt, 1) = 3         ' Typ byte
        eb(ebcnt, 2) = ecx_pos(3) ' Position pozice
        eb(ebcnt, 3) = ecx_cnt(3) ' delka
        ebcnt = ebcnt + 1
        ecx_poc(3) = ecx_poc(3) + 1
      End If
    End If
    If k = -5 Then Exit For
    If (k >= 0) Then ' We can alphanumeric? (Muzeme alnum)
      If (k >= 10 And ecx_cnt(1) >= 12) Then ' Until now it was perhaps numeric (Az dosud bylo mozno num)
        If (ecx_cnt(2) - ecx_cnt(1)) >= 8 Or (ecx_cnt(3) = ecx_cnt(2)) Then ' There is also an alphanumeric which is worth it(Je tam i alnum ktery stoji za to)
          If (ecx_cnt(3) > ecx_cnt(2)) Then ' Even before it was alnum byte (Jeste pred alnum bylo byte)
            eb(ebcnt, 1) = 3         ' Typ byte
            eb(ebcnt, 2) = ecx_pos(3) ' Position (pozice)
            eb(ebcnt, 3) = ecx_cnt(3) - ecx_cnt(2) ' length (delka)
            ebcnt = ebcnt + 1
            ecx_poc(3) = ecx_poc(3) + 1
          End If
          eb(ebcnt, 1) = 2         ' Typ alnum
          eb(ebcnt, 2) = ecx_pos(2)
          eb(ebcnt, 3) = ecx_cnt(2) - ecx_cnt(1) ' length (delka)
          ebcnt = ebcnt + 1
          ecx_poc(2) = ecx_poc(2) + 1
          ecx_cnt(2) = 0 ' processed everything (vse zpracovano)
        ElseIf (ecx_cnt(3) > ecx_cnt(1)) Then ' Previous Num is byte (Pred Num je byte)
          eb(ebcnt, 1) = 3         ' Typ byte
          eb(ebcnt, 2) = ecx_pos(3) ' Position (pozice)
          eb(ebcnt, 3) = ecx_cnt(3) - ecx_cnt(1) ' length (delka)
          ebcnt = ebcnt + 1
          ecx_poc(3) = ecx_poc(3) + 1
        End If
        eb(ebcnt, 1) = 1         ' Typ numerix
        eb(ebcnt, 2) = ecx_pos(1)
        eb(ebcnt, 3) = ecx_cnt(1) ' length (delka)
        ebcnt = ebcnt + 1
        ecx_poc(1) = ecx_poc(1) + 1
        ecx_cnt(1) = 0
        ecx_cnt(2) = 0
        ecx_cnt(3) = 0 ' processed everything (vse zpracovano)
      End If
      If ecx_cnt(2) = 0 Then ecx_pos(2) = i
      ecx_cnt(2) = ecx_cnt(2) + 1
    Else ' possible alnum (mozno alnum)
      ecx_cnt(2) = 0
    End If
    If k >= 0 And k < 10 Then ' Can be numeric (muze byt numeric)
      If ecx_cnt(1) = 0 Then ecx_pos(1) = i
      ecx_cnt(1) = ecx_cnt(1) + 1
    Else
      ecx_cnt(1) = 0
    End If
    If ecx_cnt(3) = 0 Then ecx_pos(3) = i
    ecx_cnt(3) = ecx_cnt(3) + m
    utf8 = utf8 + m
    If ebcnt >= 16 Then ' We have already taken 3 other blocks of bits (Uz by se mi tri dalsi bloky stejne nevesli)
      ecx_cnt(1) = 0
      ecx_cnt(2) = 0
    End If
    If MsgBox("Character:'" & Mid(ptext, i, 1) & "'(" & k & _
        ") ebn=" & ecx_pos(1) & "." & ecx_cnt(1) & _
         " eba=" & ecx_pos(2) & "." & ecx_cnt(2) & _
         " ebb=" & ecx_pos(3) & "." & ecx_cnt(3), vbOKCancel) > 1 Then
        Exit Function
    End If
  Next
  ebcnt = ebcnt - 1 ' ebcnt now has its final value
  c = 0
  For i = 1 To ebcnt
    Select Case eb(i, 1)
      Case 1: eb(i, 4) = Int(eb(i, 3) / 3) * 10 + (eb(i, 3) Mod 3) * 3 + Iif((eb(i, 3) Mod 3) > 0, 1, 0)
      Case 2: eb(i, 4) = Int(eb(i, 3) / 2) * 11 + (eb(i, 3) Mod 2) * 6
      Case 3: eb(i, 4) = eb(i, 3) * 8
    End Select
    c = c + eb(i, 4)
  Next i
'  UTF-8 is default not need ECI value - zxing cannot recognize
'  Call qr_params(i * 8 + utf8,mode,qrp)
  Call qr_params(c, ecl, qrp, ecx_poc)
  If qrp(1) <= 0 Then
    err = "Too long"
    Exit Function
  End If
  siz = qrp(2)
'MsgBox "ver:" & qrp(1) & mode & " size " & siz & " ecc:" & qrp(3) & "x" & qrp(4) & " d:" & (qrp(5) - qrp(3) * qrp(4))
  ReDim encoded1(qrp(5) + 2)
  ' Table 3 — Number of bits in character count indicator for QR Code 2005:
  ' mode indicator (1=num,2=AlNum,4=Byte,8=kanji,ECI=7)
  '      mode: Byte Alphanum  Numeric  Kanji
  ' ver 1..9 :  8      9       10       8
  '   10..26 : 16     11       12      10
  '   27..40 : 16     13       14      12
' UTF-8 is default not need ECI value - zxing cannot recognize
'  if utf8 > 0 Then
'    k = &H700 + 26 ' UTF-8=26 ; Win1250 = 21; 8859-2 = 4 viz http://strokescribe.com/en/ECI.html
'    bb_putbits(encoded1,encix1,k,12)
'  End If
  encix1 = 0
  For i = 1 To ebcnt
    Select Case eb(i, 1)
      Case 1: c = Iif(qrp(1) < 10, 10, Iif(qrp(1) < 27, 12, 14)): k = 2 ^ c + eb(i, 3) ' encoding mode "Numeric"
      Case 2: c = Iif(qrp(1) < 10, 9, Iif(qrp(1) < 27, 11, 13)): k = 2 * (2 ^ c) + eb(i, 3) ' encoding mode "alphanum
      Case 3: c = Iif(qrp(1) < 10, 8, 16): k = 4 * (2 ^ c) + eb(i, 3) ' encoding mode "Byte"
    End Select
    Call bb_putbits(encoded1, encix1, k, c + 4)
    j = 0
    m = eb(i, 2)
    r = 0
    While j < eb(i, 3)
      k = AscL(Mid(ptext, m, 1))
      m = m + 1
      If eb(i, 1) = 1 Then
        r = (r * 10) + ((k - &H30) Mod 10)
        If (j Mod 3) = 2 Then
          Call bb_putbits(encoded1, encix1, r, 10)
          r = 0
        End If
        j = j + 1
      ElseIf eb(i, 1) = 2 Then
        r = (r * 45) + ((InStr(qralnum, Chr(k)) - 1) Mod 45)
        If (j Mod 2) = 1 Then
          Call bb_putbits(encoded1, encix1, r, 11)
          r = 0
        End If
        j = j + 1
      Else
        If k > &H1FFFFF Then ' FFFF - 1FFFFFFF
          ch = &HF0 + Int(k / &H40000) Mod 8
          Call bb_putbits(encoded1, encix1, ch, 8)
          ch = 128 + Int(k / &H1000) Mod 64
          Call bb_putbits(encoded1, encix1, ch, 8)
          ch = 128 + Int(k / 64) Mod 64
          Call bb_putbits(encoded1, encix1, ch, 8)
          ch = 128 + k Mod 64
          Call bb_putbits(encoded1, encix1, ch, 8)
          j = j + 4
        ElseIf k > &H7FF Then ' 7FF-FFFF 3 bytes
          ch = &HE0 + Int(k / &H1000) Mod 16
          Call bb_putbits(encoded1, encix1, ch, 8)
          ch = 128 + Int(k / 64) Mod 64
          Call bb_putbits(encoded1, encix1, ch, 8)
          ch = 128 + k Mod 64
          Call bb_putbits(encoded1, encix1, ch, 8)
          j = j + 3
        ElseIf k > &H7F Then ' 2 bytes
          ch = &HC0 + Int(k / 64) Mod 32
          Call bb_putbits(encoded1, encix1, ch, 8)
          ch = 128 + k Mod 64
          Call bb_putbits(encoded1, encix1, ch, 8)
          j = j + 2
        Else
          ch = k Mod 256
          Call bb_putbits(encoded1, encix1, ch, 8)
          j = j + 1
        End If
      End If
    Wend
    Select Case eb(i, 1)
      Case 1:
        If (j Mod 3) = 1 Then
          Call bb_putbits(encoded1, encix1, r, 4)
        ElseIf (j Mod 3) = 2 Then
          Call bb_putbits(encoded1, encix1, r, 7)
        End If
      Case 2:
        If (j Mod 2) = 1 Then Call bb_putbits(encoded1, encix1, r, 6)
    End Select
'MsgBox "blk[" & i & "] t:" & eb(i,1) & "from " & eb(i,2) & " to " & eb(i,3) + eb(i,2) & " bits=" & encix1
  Next i
  Call bb_putbits(encoded1, encix1, 0, 4) ' end of chain
  If (encix1 Mod 8) <> 0 Then  ' round to byte
    Call bb_putbits(encoded1, encix1, 0, 8 - (encix1 Mod 8))
  End If
  ' padding
  i = (qrp(5) - qrp(3) * qrp(4)) * 8
  If encix1 > i Then
    err = "Encode length error"
    Exit Function
  End If
  ' padding 0xEC,0x11,0xEC,0x11...
  Do While encix1 < i
    Call bb_putbits(encoded1, encix1, &HEC11, 16)
  Loop
  ' doplnime ECC
  i = qrp(3) * qrp(4) 'ppoly, pmemptr , psize , plen , pblocks
  Call qr_rs(&H11D, encoded1, qrp(5) - i, i, qrp(4))
'Call arr2hexstr(encoded1)
  encix1 = qrp(5)
  ' Pole pro vystup
  ReDim qrarr(0)
  ReDim qrarr(1, qrp(2) * 24& + 24&) ' 24 bytes per row
  qrarr(0, 0) = 0
  ch = 0
  Call bb_putbits(qrsync1, ch, Array(&HFE, &H82, &HBA, &HBA, &HBA, &H82, &HFE, 0), 64)
  Call qr_mask(qrarr, qrsync1, 8, 0, 0) ' sync UL
  Call qr_mask(qrarr, 0, 8, 8, 0)   ' fmtinfo UL under - bity 14..9 SYNC 8
  Call qr_mask(qrarr, qrsync1, 8, 0, siz - 7) ' sync UR ( o bit vlevo )
  Call qr_mask(qrarr, 0, 8, 8, siz - 8)   ' fmtinfo UR - bity 7..0
  Call qr_mask(qrarr, qrsync1, 8, siz - 7, 0) ' sync DL (zasahuje i do quiet zony)
  Call qr_mask(qrarr, 0, 8, siz - 8, 0)   ' blank nad DL
  For i = 0 To 6
    x = qr_bit(qrarr, -1, i, 8, 0) ' svisle fmtinfo UL - bity 0..5 SYNC 6,7
    x = qr_bit(qrarr, -1, i, siz - 8, 0) ' svisly blank pred UR
    x = qr_bit(qrarr, -1, siz - 1 - i, 8, 0) ' svisle fmtinfo DL - bity 14..8
  Next
  x = qr_bit(qrarr, -1, 7, 8, 0) ' svisle fmtinfo UL - bity 0..5 SYNC 6,7
  x = qr_bit(qrarr, -1, 7, siz - 8, 0) ' svisly blank pred UR
  x = qr_bit(qrarr, -1, 8, 8, 0) ' svisle fmtinfo UL - bity 0..5 SYNC 6,7
  x = qr_bit(qrarr, -1, siz - 8, 8, 1) ' black dot DL
  If qrp(13) <> 0 Or qrp(14) <> 0 Then ' versioninfo
  ' UR ver 0 1 2;3 4 5;...;15 16 17
  ' LL ver 0 3 6 9 12 15;1 4 7 10 13 16; 2 5 8 11 14 17
    k = 65536 * qrp(13) + 256& * qrp(14) + 1& * qrp(15)
    c = 0: r = 0
    For i = 0 To 17
      ch = k Mod 2
      x = qr_bit(qrarr, -1, r, siz - 11 + c, ch) ' UR ver
      x = qr_bit(qrarr, -1, siz - 11 + c, r, ch) ' DL ver
      c = c + 1
      If c > 2 Then c = 0: r = r + 1
      k = Int(k / 2&)
    Next
  End If
  c = 1
  For i = 8 To siz - 9 ' sync lines
    x = qr_bit(qrarr, -1, i, 6, c) ' vertical on column 6
    x = qr_bit(qrarr, -1, 6, i, c) ' horizontal on row 6
    c = (c + 1) Mod 2
  Next
  ' other syncs
  ch = 0
  Call bb_putbits(qrsync2, ch, Array(&H1F, &H11, &H15, &H11, &H1F), 40)
  ch = 6
  Do While ch > 0 And qrp(6 + ch) = 0
    ch = ch - 1
  Loop
  If ch > 0 Then
    For c = 0 To ch
      For r = 0 To ch
        ' corners
        If (c <> 0 Or r <> 0) And _
           (c <> ch Or r <> 0) And _
           (c <> 0 Or r <> ch) Then
          Call qr_mask(qrarr, qrsync2, 5, qrp(r + 6) - 2, qrp(c + 6) - 2)
        End If
      Next r
    Next c
  End If
 ' qr_fill(parr as Variant, psiz%, pb as Variant, pblocks%, pdlen%, ptlen%)
 ' vyplni pole parr (psiz x 24 bytes) z pole pb pdlen = pocet dbytes, pblocks = bloku, ptlen celkem
  Call qr_fill(qrarr, siz, encoded1, qrp(4), qrp(5) - qrp(3) * qrp(4), qrp(5))
  mask = 8 ' auto
  i = InStr(poptions, "mask=")
  If i > 0 Then mask = val(Mid(poptions, i + 5, 1))
  If mask < 0 Or mask > 7 Then
    j = -1
    For mask = 0 To 7
      GoSub addmm
      i = qr_xormask(qrarr, siz, mask, False)
'      MsgBox "score mask " & mask & " is " & i
      If i < j Or j = -1 Then j = i: s = mask
    Next mask
    mask = s
'    MsgBox "best is " & mask & " with score " & j
  End If
  GoSub addmm
  i = qr_xormask(qrarr, siz, mask, True)
  ascimatrix = ""
  For r = 0 To siz Step 2
    s = 0
    For c = 0 To siz Step 2
      If (c Mod 8) = 0 Then
        ch = qrarr(1, s + 24 * r)
        If r < siz Then i = qrarr(1, s + 24 * (r + 1)) Else i = 0
        s = s + 1
      End If
      ascimatrix = ascimatrix _
         & Chr(97 + (ch Mod 4) + 4 * (i Mod 4))
      ch = Int(ch / 4)
      i = Int(i / 4)
    Next
    ascimatrix = ascimatrix & vbNewLine
  Next r
  ReDim qrarr(0)
  qr_gen = ascimatrix
  Exit Function
addmm:
  k = ecl * 8 + mask
  ' poly: 101 0011 0111
  Call qr_bch_calc(k, &H537)
'MsgBox "mask :" & hex(k,3) & " " & hex(k xor &H5412,3)
  k = k Xor &H5412 ' micro xor &H4445
  r = 0
  c = siz - 1
  For i = 0 To 14
    ch = k Mod 2
    k = Int(k / 2)
    x = qr_bit(qrarr, -1, r, 8, ch) ' svisle fmtinfo UL - bity 0..5 SYNC 6,7 .... 8..14 dole
    x = qr_bit(qrarr, -1, 8, c, ch) ' vodorovne odzadu 0..7 ............ 8,SYNC,9..14
    c = c - 1
    r = r + 1
    If i = 7 Then c = 7: r = siz - 7
    If i = 5 Then r = r + 1 ' preskoc sync vodorvny
    If i = 8 Then c = c - 1 ' preskoc sync svisly
  Next
  Return
End Function  ' qr_gen

Sub bc_2D(ShIx As Integer, xAddr As String, xBC As String)
  Dim xPage As Object
  Dim xShape As Object
  Dim xDoc As Object
  Dim xView As Object
  Dim xProv As Object
  Dim xSheet As Object
  Dim xRange As Object
  Dim xCell As Object
  Dim xPos As New com.sun.star.awt.Point
  Dim xPosOld As New com.sun.star.awt.Point
  Dim xSize As New com.sun.star.awt.Size
  Dim xSizeOld As New com.sun.star.awt.Size
  Dim xGrp As Object
  Dim xSolid As Long
  Dim x&, y&, n%, w%, s$, p$, m&, dm&, a&, b%
  
  xDoc = ThisComponent
  On Error GoTo e2derr
  xView = ThisComponent.getCurrentController()
  xSheet = xDoc.Sheets.getByIndex(ShIx - 1)
  xCell = xSheet.getCellRangeByName(xAddr)
  xPage = xSheet.getDrawPage()
  On Error GoTo 0
  m = 60 ' block size
  xSolid = 1 ' com.sun.star.drawing.FillStyle.SOLID = 1
  xPosOld.x = xCell.Position.x
  xPosOld.y = xCell.Position.y
  xSizeOld.Width = 0
  xSizeOld.Height = 0
  s = "BC" & xAddr & "#GR"
  If xPage.hasElements() Then
    For n = (xPage.getCount() - 1) To 0 Step -1
      xShape = xPage.getByIndex(n)
      If xShape.Name = s Then
        xPosOld.x = xShape.Position.x
        xPosOld.y = xShape.Position.y
        xSizeOld.Width = xShape.Size.Width
        xSizeOld.Height = xShape.Size.Height
        xPage.remove (xShape)
      End If
    Next n
  End If
  x = 0
  y = 0
  a = 0
  dm = m * 2&
  n = 1
  p = Trim(xBC)
  b = Len(p)
  'bbccddeeffgghhiijjkkllmmnnoopp
  '^  ^^^. I .^I^ .^. I^I..I..III
  Do While n <= b
    w = AscL(Mid(p, n, 1)) Mod 256
    If w >= 97 And w <= 112 Then
      a = a + dm
    End If
    If w = 10 Or n = b Then
      y = y + dm
      If a > x Then x = a
      a = 0
    End If
    n = n + 1
  Loop
  If x = 0 Or y = 0 Then Exit Sub
  xGrp = xDoc.createInstance("com.sun.star.drawing.GroupShape")
  xGrp.Name = s
  xPage.add (xGrp)
  xShape = xDoc.createInstance("com.sun.star.drawing.RectangleShape")
  xShape.LineWidth = 0
  xShape.LineStyle = com.sun.star.drawing.LineStyle.NONE
  xShape.FillStyle = xSolid
  xShape.FillColor = RGB(255, 255, 255)
  xPos.x = 0
  xPos.y = 0
  xShape.Position = xPos
  xSize.Width = x
  xSize.Height = y
  xShape.Size = xSize
  xGrp.add (xShape)
  x = 0
  y = 0
  a = 1
  For n = 1 To b
    w = AscL(Mid(p, n, 1)) Mod 256
    If w = 10 Then
      y = y + dm
      x = 0
    ElseIf (w >= 97 And w <= 112) Then
      w = w - 97
      xSize.Height = m: xSize.Width = m: xPos.x = x: xPos.y = y
      Select Case w
        Case 1: GoSub crrect
        Case 2: xPos.x = x + m: GoSub crrect
        Case 3: xSize.Width = dm: GoSub crrect
        Case 4: xPos.y = y + m: GoSub crrect
        Case 5: xSize.Height = dm: GoSub crrect
        Case 6: xPos.x = x + m: GoSub crrect: xPos.x = x: xPos.y = y + m: GoSub crrect
        Case 7: xSize.Width = dm: GoSub crrect: xSize.Width = m: xPos.y = y + m: GoSub crrect
        Case 8: xPos.y = y + m: xPos.x = x + m: GoSub crrect
        Case 9: GoSub crrect: xPos.y = y + m: xPos.x = x + m: GoSub crrect
        Case 10: xPos.x = x + m: xSize.Height = dm: GoSub crrect
        Case 11: GoSub crrect: xPos.x = x + m: xSize.Height = dm: GoSub crrect
        Case 12: xPos.y = y + m: xSize.Width = dm: GoSub crrect
        Case 13: GoSub crrect: xPos.y = y + m: xSize.Width = dm: GoSub crrect
        Case 14: xPos.x = x + m: GoSub crrect: xPos.x = x: xPos.y = y + m: xSize.Width = dm: GoSub crrect
        Case 15: xSize.Width = dm: xSize.Height = dm: GoSub crrect
      End Select
      x = x + dm
    End If
  Next n
  xGrp.Visible = True
  xGrp.Position = xPosOld
  If xSizeOld.Width > 0 Then xGrp.Size = xSizeOld
  Erase xPos
  Erase xSize
  Erase xPosOld
  Erase xSizeOld
  Exit Sub
crrect:
  xShape = xDoc.createInstance("com.sun.star.drawing.RectangleShape")
  xShape.LineWidth = 0
  xShape.LineStyle = com.sun.star.drawing.LineStyle.NONE
  xShape.LineColor = RGB(255, 255, 255)
  xShape.FillStyle = xSolid
  xShape.FillColor = RGB(0, 0, 0)
  xShape.Position = xPos
  xShape.Size = xSize
  xShape.Name = xAddr & "#BR" & a
  xGrp.add (xShape)
  a = a + 1
  Return
e2derr:
  On Error GoTo 0
End Sub

Sub bc_1D(ShIx As Integer, xAddr As String, xBC As String)
  Dim xPage As Object
  Dim xShape As Object
  Dim xDoc As Object
  Dim xView As Object
  Dim xProv As Object
  Dim xSheet As Object
  Dim xRange As Object
  Dim xCell As Object
  Dim xPos As New com.sun.star.awt.Point
  Dim xPosOld As New com.sun.star.awt.Point
  Dim xSize As New com.sun.star.awt.Size
  Dim xSizeOld As New com.sun.star.awt.Size
  Dim xGrp As Object
  Dim xSolid As Long
  Dim x&, n%, w%, s$, m&
  
  xDoc = ThisComponent
  On Error GoTo e1derr
  xView = ThisComponent.getCurrentController()
  xSheet = xDoc.Sheets.getByIndex(ShIx - 1)
  xCell = xSheet.getCellRangeByName(xAddr)
  xPage = xSheet.getDrawPage()
  On Error GoTo 0
  m = 60&
  xSolid = 1 ' com.sun.star.drawing.FillStyle.SOLID = 1
  xPosOld.x = xCell.Position.x
  xPosOld.y = xCell.Position.y
  xSizeOld.Width = 0
  xSizeOld.Height = 0
  s = "BC" & xAddr & "#GR"
  If xPage.hasElements() Then
    For n% = (xPage.getCount() - 1) To 0 Step -1
      xShape = xPage.getByIndex(n%)
      If xShape.Name = s Then
        xPosOld.x = xShape.Position.x
        xPosOld.y = xShape.Position.y
        xSizeOld.Width = xShape.Size.Width
        xSizeOld.Height = xShape.Size.Height
        xPage.remove (xShape)
      End If
    Next n%
  End If
  x = 0
  For n = 1 To Len(xBC)
    w = AscL(Mid(xBC, n, 1)) Mod 256
    If (w >= 48 And w <= 57) Then
      w = (w - 48) Mod 5 + 1
    ElseIf (w >= 65 And w <= 69) Then
      w = w - 64
    Else
      w = 0
    End If
    x = x + m * w
  Next n
  If x = 0 Then Exit Sub
  xGrp = xDoc.createInstance("com.sun.star.drawing.GroupShape")
  xGrp.Name = s
  xPage.add (xGrp)
  xShape = xDoc.createInstance("com.sun.star.drawing.RectangleShape")
  xShape.LineWidth = 0
  xShape.LineStyle = com.sun.star.drawing.LineStyle.NONE
  xShape.FillStyle = xSolid
  xShape.FillColor = RGB(255, 255, 255)
  xPos.x = 0
  xPos.y = 0
  xShape.Position = xPos
  xSize.Width = x
  xSize.Height = m * 18
  xShape.Size = xSize
  xGrp.add (xShape)
  x = 0
  For n = 1 To Len(xBC)
    w = AscL(Mid(xBC, n, 1)) Mod 256
    If (w >= 48 And w <= 57) Then
      xShape = xDoc.createInstance("com.sun.star.drawing.RectangleShape")
      xShape.LineWidth = 0
      xShape.LineStyle = com.sun.star.drawing.LineStyle.NONE
      If w >= 53 Then xSize.Height = m * 15 Else xSize.Height = m * 17
      w = (w - 48) Mod 5 + 1
      xShape.FillStyle = xSolid
      xShape.FillColor = RGB(0, 0, 0)
      xPos.x = x
      xPos.y = 0
      xSize.Width = m * w
      xShape.Position = xPos
      xShape.Size = xSize
      xShape.Name = xAddr & "#BR" & x
      xGrp.add (xShape)
    ElseIf (w >= 65 And w <= 69) Then
      w = w - 64
    Else
      w = 0
    End If
    x = x + m * w
  Next n
  xGrp.Visible = True
  xGrp.Position = xPosOld
  If xSizeOld.Width > 0 Then xGrp.Size = xSizeOld
  Erase xPos
  Erase xSize
  Erase xPosOld
  Erase xSizeOld
  Exit Sub
e1derr:
  Exit Sub
End Sub

Sub bc_2Dms(xBC As String, Optional xNam As String)
 Dim xShape As Shape, xBkgr As Shape
 Dim xSheet As Worksheet
 Dim xRange As Range, xCell As Range
 Dim xAddr As String
 Dim xPosOldX As Double, xPosOldY As Double
 Dim xSizeOldW As Double, xSizeOldH As Double
 Dim x, y, m, dm, a As Double
 Dim b%, n%, w%, p$, s$, h%, g%
 
 If TypeName(Application.Caller) = "Range" Then
   Set xSheet = Application.Caller.Worksheet
   Set xRange = Application.Caller
   xAddr = xRange.Address
   xPosOldX = xRange.Left
   xPosOldY = xRange.Top
 Else
   Set xSheet = Worksheets(1)
   If IsMissing(xNam) Then
     xAddr = "QR"
   Else
     xAddr = xNam
   End If
 End If
 xSizeOldW = 0
 xSizeOldH = 0
 s = "BC" & xAddr & "#GR"
 x = 0#
 y = 0#
 m = 2.5
 dm = m * 2#
 a = 0#
 p = Trim(xBC)
 b = Len(p)
 For n = 1 To b
   w = AscL(Mid(p, n, 1)) Mod 256
   If (w >= 97 And w <= 112) Then
     a = a + dm
   ElseIf w = 10 Or n = b Then
     If x < a Then x = a
     y = y + dm
     a = 0#
   End If
 Next n
 If x <= 0# Then Exit Sub
 On Error Resume Next
 Set xShape = xSheet.Shapes(s)
 On Error GoTo 0
 If Not (xShape Is Nothing) Then
   xPosOldX = xShape.Left
   xPosOldY = xShape.Top
   xSizeOldW = xShape.Width
   xSizeOldH = xShape.Height
   xShape.Delete
 End If
 On Error Resume Next
 xSheet.Shapes("BC" & xAddr & "#BK").Delete
 On Error GoTo 0
 Set xBkgr = xSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, x, y)
 xBkgr.Line.Visible = msoFalse
 xBkgr.Line.Weight = 0#
 xBkgr.Line.ForeColor.RGB = RGB(255, 255, 255)
 xBkgr.Fill.Solid
 xBkgr.Fill.ForeColor.RGB = RGB(255, 255, 255)
 xBkgr.Name = "BC" & xAddr & "#BK"
 Set xShape = Nothing
 x = 0#
 y = 0#
 g = 0
 For n = 1 To b
   w = AscL(Mid(p, n, 1)) Mod 256
   If w = 10 Then
     y = y + dm
     x = 0#
   ElseIf (w >= 97 And w <= 112) Then
     w = w - 97
     With xSheet.Shapes
     Select Case w
       Case 1: Set xShape = .AddShape(msoShapeRectangle, x, y, m, m): GoSub fmtxshape
       Case 2: Set xShape = .AddShape(msoShapeRectangle, x + m, y, m, m): GoSub fmtxshape
       Case 3: Set xShape = .AddShape(msoShapeRectangle, x, y, dm, m): GoSub fmtxshape
       Case 4: Set xShape = .AddShape(msoShapeRectangle, x, y + m, m, m): GoSub fmtxshape
       Case 5: Set xShape = .AddShape(msoShapeRectangle, x, y, m, dm): GoSub fmtxshape
       Case 6: Set xShape = .AddShape(msoShapeRectangle, x + m, y, m, m): GoSub fmtxshape
               Set xShape = .AddShape(msoShapeRectangle, x, y + m, m, m): GoSub fmtxshape
       Case 7: Set xShape = .AddShape(msoShapeRectangle, x, y, dm, m): GoSub fmtxshape
               Set xShape = .AddShape(msoShapeRectangle, x, y + m, m, m): GoSub fmtxshape
       Case 8: Set xShape = .AddShape(msoShapeRectangle, x + m, y + m, m, m): GoSub fmtxshape
       Case 9: Set xShape = .AddShape(msoShapeRectangle, x, y, m, m): GoSub fmtxshape
               Set xShape = .AddShape(msoShapeRectangle, x + m, y + m, m, m): GoSub fmtxshape
       Case 10: Set xShape = .AddShape(msoShapeRectangle, x + m, y, m, dm): GoSub fmtxshape
       Case 11: Set xShape = .AddShape(msoShapeRectangle, x, y, dm, m): GoSub fmtxshape
                Set xShape = .AddShape(msoShapeRectangle, x + m, y + m, m, m): GoSub fmtxshape
       Case 12: Set xShape = .AddShape(msoShapeRectangle, x, y + m, dm, m): GoSub fmtxshape
       Case 13: Set xShape = .AddShape(msoShapeRectangle, x, y, m, m): GoSub fmtxshape
                Set xShape = .AddShape(msoShapeRectangle, x, y + m, dm, m): GoSub fmtxshape
       Case 14: Set xShape = .AddShape(msoShapeRectangle, x + m, y, m, m): GoSub fmtxshape
                Set xShape = .AddShape(msoShapeRectangle, x, y + m, dm, m): GoSub fmtxshape
       Case 15: Set xShape = .AddShape(msoShapeRectangle, x, y, dm, dm): GoSub fmtxshape
     End Select
     End With
     x = x + dm
   End If
 Next n
 On Error Resume Next
 Set xShape = xSheet.Shapes(s)
 On Error GoTo 0
 If Not (xShape Is Nothing) Then
   xShape.Left = xPosOldX
   xShape.Top = xPosOldY
   If xSizeOldW > 0 Then
     xShape.Width = xSizeOldW
     xShape.Height = xSizeOldH
   End If
 Else
   If Not (xBkgr Is Nothing) Then xBkgr.Delete
 End If
 Exit Sub
fmtxshape:
  xShape.Line.Visible = msoFalse
  xShape.Line.Weight = 0#
  xShape.Fill.Solid
  xShape.Fill.ForeColor.RGB = RGB(0, 0, 0)
  g = g + 1
  xShape.Name = "BC" & xAddr & "#BR" & g
  If g = 1 Then
    xSheet.Shapes.Range(Array(xBkgr.Name, xShape.Name)).Group.Name = s
  Else
    xSheet.Shapes.Range(Array(s, xShape.Name)).Group.Name = s
  End If
  Return
End Sub

Sub bc_1Dms(xBC As String)
 Dim xShape As Shape, xBkgr As Shape
 Dim xSheet As Worksheet
 Dim xRange As Range, xCell As Range
 Dim xAddr As String
 Dim xPosOldX As Double, xPosOldY As Double
 Dim xSizeOldW As Double, xSizeOldH As Double
' Dim xGrp As ShapeRange
 Dim x As Double
 Dim n%, w%, s$, h%, g%
 
 If TypeName(Application.Caller) <> "Range" Then
   Exit Sub
 End If
 Set xSheet = Application.Caller.Worksheet
 Set xRange = Application.Caller
' Set xCell = xRange("A1")
 xAddr = xRange.Address
 xPosOldX = xRange.Left
 xPosOldY = xRange.Top
 xSizeOldW = 0
 xSizeOldH = 0
 s = "BC" & xAddr & "#GR"
 x = 0
 For n = 1 To Len(xBC)
   w = AscL(Mid(xBC, n, 1)) Mod 256
   If (w >= 48 And w <= 57) Then
     w = (w - 48) Mod 5 + 1
   ElseIf (w >= 65 And w <= 69) Then
     w = w - 64
   Else
     w = 0
   End If
   x = x + 1.5 * w
 Next n
 If x <= 0# Then Exit Sub
 On Error Resume Next
 Set xShape = xSheet.Shapes(s)
 On Error GoTo 0
 If Not (xShape Is Nothing) Then
   xPosOldX = xShape.Left
   xPosOldY = xShape.Top
   xSizeOldW = xShape.Width
   xSizeOldH = xShape.Height
   xShape.Delete
 End If
 On Error Resume Next
 xSheet.Shapes("BC" & xAddr & "#BK").Delete
 On Error GoTo 0
 Set xBkgr = xSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, x, 51#)
 xBkgr.Line.Visible = msoFalse
 xBkgr.Line.Weight = 0#
 xBkgr.Line.ForeColor.RGB = RGB(255, 255, 255)
 xBkgr.Fill.Solid
 xBkgr.Fill.ForeColor.RGB = RGB(255, 255, 255)
 xBkgr.Name = "BC" & xAddr & "#BK"
 Set xShape = Nothing
 x = 0#
 g = 0
 For n = 1 To Len(xBC)
   w = AscL(Mid(xBC, n, 1)) Mod 256
   If (w >= 48 And w <= 57) Then
     If w >= 53 Then h = 47 Else h = 50
     w = (w - 48) Mod 5 + 1
     Set xShape = xSheet.Shapes.AddShape(msoShapeRectangle, x, 0, 1.5 * w, h)
     xShape.Line.Visible = msoFalse
     xShape.Line.Weight = 0#
     xShape.Fill.Solid
     xShape.Fill.ForeColor.RGB = RGB(0, 0, 0)
     g = g + 1
     xShape.Name = "BC" & xAddr & "#BR" & g
     If g = 1 Then
       xSheet.Shapes.Range(Array(xBkgr.Name, xShape.Name)).Group.Name = s
     Else
       xSheet.Shapes.Range(Array(s, xShape.Name)).Group.Name = s
     End If
   ElseIf (w >= 65 And w <= 69) Then
     w = w - 64
   Else
     w = 0
   End If
   x = x + 1.5 * w
 Next n
 On Error Resume Next
 Set xShape = xSheet.Shapes(s)
 On Error GoTo 0
 If Not (xShape Is Nothing) Then
   xShape.Left = xPosOldX
   xShape.Top = xPosOldY
   If xSizeOldW > 0 Then
     xShape.Width = xSizeOldW
     xShape.Height = xSizeOldH
   End If
 Else
   If Not (xBkgr Is Nothing) Then xBkgr.Delete
 End If
End Sub

