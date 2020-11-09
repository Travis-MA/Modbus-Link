Attribute VB_Name = "CRC"
'高位函数
Function GetCRCHi(Ind As Long) As Byte
GetCRCHi = Choose(Ind + 1, &H0, &HC0, &HC1, &H1, &HC3, &H3, &H2, &HC2, &HC6, _
    &H6, &H7, &HC7, &H5, &HC5, &HC4, &H4, &HCC, &HC, &HD, &HCD, &HF, &HCF, &HCE, _
    &HE, &HA, &HCA, &HCB, &HB, &HC9, &H9, &H8, &HC8, &HD8, &H18, &H19, &HD9, &H1B, _
    &HDB, &HDA, &H1A, &H1E, &HDE, &HDF, &H1F, &HDD, &H1D, &H1C, &HDC, &H14, &HD4, _
    &HD5, &H15, &HD7, &H17, &H16, &HD6, &HD2, &H12, &H13, &HD3, &H11, &HD1, &HD0, _
    &H10, &HF0, &H30, &H31, &HF1, &H33, &HF3, &HF2, &H32, &H36, &HF6, &HF7, &H37, _
    &HF5, &H35, &H34, &HF4, &H3C, &HFC, &HFD, &H3D, &HFF, &H3F, &H3E, &HFE, &HFA, _
    &H3A, &H3B, &HFB, &H39, &HF9, &HF8, &H38, &H28, &HE8, &HE9, &H29, &HEB, &H2B, _
    &H2A, &HEA, &HEE, &H2E, &H2F, &HEF, &H2D, &HED, &HEC, &H2C, &HE4, &H24, &H25, _
    &HE5, &H27, &HE7, &HE6, &H26, &H22, &HE2, &HE3, &H23, &HE1, &H21, &H20, &HE0, _
    &HA0, &H60, &H61, &HA1, &H63, &HA3, &HA2, &H62, &H66, &HA6, &HA7, &H67, &HA5, _
    &H65, &H64, &HA4, &H6C, &HAC, &HAD, &H6D, &HAF, &H6F, &H6E, &HAE, &HAA, &H6A, _
    &H6B, &HAB, &H69, &HA9, &HA8, &H68, &H78, &HB8, &HB9, &H79, &HBB, &H7B, &H7A, _
    &HBA, &HBE, &H7E, &H7F, &HBF, &H7D, &HBD, &HBC, &H7C, &HB4, &H74, &H75, &HB5, _
    &H77, &HB7, &HB6, &H76, &H72, &HB2, &HB3, &H73, &HB1, &H71, &H70, &HB0, &H50, _
    &H90, &H91, &H51, &H93, &H53, &H52, &H92, &H96, &H56, &H57, &H97, &H55, &H95, _
    &H94, &H54, &H9C, &H5C, &H5D, &H9D, &H5F, &H9F, &H9E, &H5E, &H5A, &H9A, &H9B, _
    &H5B, &H99, &H59, &H58, &H98, &H88, &H48, &H49, &H89, &H4B, &H8B, &H8A, &H4A, _
    &H4E, &H8E, &H8F, &H4F, &H8D, &H4D, &H4C, &H8C, &H44, &H84, &H85, &H45, &H87, _
    &H47, &H46, &H86, &H82, &H42, &H43, &H83, &H41, &H81, &H80, &H40)
End Function

'低位函数:

Function GetCRCLo(Ind As Long) As Byte
GetCRCLo = Choose(Ind + 1, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, _
    &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H0, &HC1, _
    &H81, &H40, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, &H41, _
    &H0, &HC1, &H81, &H40, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H0, &HC1, _
    &H81, &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, _
    &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, &H0, &HC1, &H81, &H40, &H1, &HC0, _
    &H80, &H41, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, _
    &H41, &H0, &HC1, &H81, &H40, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, _
    &H41, &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, &H1, _
    &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, &H0, &HC1, &H81, _
    &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, &H41, &H0, _
    &HC1, &H81, &H40, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, _
    &H41, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, _
    &H80, &H41, &H0, &HC1, &H81, &H40, &H0, &HC1, &H81, &H40, &H1, _
    &HC0, &H80, &H41, &H1, &HC0, &H80, &H41, &H0, &HC1, _
    &H81, &H40, &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, &H0, _
    &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, _
    &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, _
    &H40, &H1, &HC0, &H80, &H41, &H0, &HC1, &H81, &H40, &H0, &HC1, _
    &H81, &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, &H41, &H0, _
    &HC1, &H81, &H40, &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, _
    &H0, &HC1, &H81, &H40, &H1, &HC0, &H80, &H41, &H1, &HC0, &H80, _
    &H41, &H0, &HC1, &H81, &H40)
End Function

'CRC计算函数:

Private Function CRC16(data() As Byte, crc() As Byte) As String
    Dim CRC16Hi As Byte
    Dim CRC16Lo As Byte
    CRC16Hi = &HFF
    CRC16Lo = &HFF
    Dim i As Integer
    Dim iIndex As Long
    For i = 0 To 5
        iIndex = CRC16Lo Xor data(i)
        CRC16Lo = CRC16Hi Xor GetCRCLo(iIndex)        '低位处理
        CRC16Hi = GetCRCHi(iIndex)                   '高位处理
    Next i
     crc(0) = CRC16Hi
     crc(1) = CRC16Lo
End Function

Public Function runCRC(ByVal a, b, c, d, e, f As Byte) As String
    ReDim buf(7) As Byte
    Dim dcrc1 As Long, crcc(1) As Byte, ret As String
    Dim astr
    buf(0) = a
    buf(1) = b
    buf(2) = c
    buf(3) = d
    buf(4) = e
    buf(5) = f
    ret = CRC16(buf, crcc()) 'CRC的入口函数，有的例子上还有数组长度，但是对于特定工程，那个不需要。一般来讲家还是那个CRC一共就是8位
    buf(6) = crcc(1) '高低位处理反了，先高后低
    buf(7) = crcc(0)
    For i = 0 To 7
    If Len(Hex(buf(i))) = 1 Then
        astr = astr & "0" & Hex(buf(i))
    Else
        astr = astr & Hex(buf(i))
    End If
Next
runCRC = astr
End Function
