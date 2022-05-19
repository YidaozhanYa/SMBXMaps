Attribute VB_Name = "Base64"
'Visual Basic 6 Base64 API Header

Private Declare Function CryptBinaryToString Lib "Crypt32.dll" Alias "CryptBinaryToStringW" (ByRef pbBinary As Byte, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
Private Declare Function CryptStringToBinary Lib "Crypt32.dll" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long

Public Function Base64Decode(sBase64Buf As String) As String
    Const CRYPT_STRING_BASE64 As Long = 1
    Dim bTmp() As Byte, lLen As Long, dwActualUsed As Long
    If CryptStringToBinary(StrPtr(sBase64Buf), Len(sBase64Buf), CRYPT_STRING_BASE64, StrPtr(vbNullString), lLen, 0&, dwActualUsed) = 0 Then Exit Function       'Get output buffer length
    ReDim bTmp(lLen - 1)
    If CryptStringToBinary(StrPtr(sBase64Buf), Len(sBase64Buf), CRYPT_STRING_BASE64, VarPtr(bTmp(0)), lLen, 0&, dwActualUsed) = 0 Then Exit Function    'Convert Base64 to binary.
    Base64Decode = StrConv(bTmp, vbUnicode)
End Function

Public Function Base64Encode(Text As String) As String
    Const CRYPT_STRING_BASE64 As Long = 1
    Dim lLen As Long, m_bData() As Byte, sBase64Buf As String
    m_bData = StrConv(Text, vbFromUnicode)
    If CryptBinaryToString(m_bData(0), UBound(m_bData) + 1, CRYPT_STRING_BASE64, StrPtr(vbNullString), lLen) = 0 Then Exit Function  'Determine Base64 output String length required.
    sBase64Buf = String$(lLen - 1, Chr$(0))    'Convert binary to Base64.
    If CryptBinaryToString(m_bData(0), UBound(m_bData) + 1, CRYPT_STRING_BASE64, StrPtr(sBase64Buf), lLen) = 0 Then Exit Function
    Base64Encode = sBase64Buf
End Function
