VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JsonWrite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_UTF8 As Long = 65001
Private Const BUFFER_INITIAL As Long = 32768

' flat utf-16 integer output buffer (zero BSTR concatenation churn)
Private m_Buf() As Integer
Private m_Pos As Long
Private m_Cap As Long

' tracks comma placement across nested containers (true = needs comma before next item)
Private m_NeedComma() As Boolean
Private m_Depth As Long
Private m_DepthCap As Long
Private m_ExpectValue As Boolean


Private Sub Class_Initialize()
    m_Cap = BUFFER_INITIAL
    ReDim m_Buf(m_Cap - 1&)
    m_Pos = 0

    m_DepthCap = 64
    ReDim m_NeedComma(m_DepthCap - 1&)
    m_Depth = 0
End Sub

Private Sub Class_Terminate()
    Erase m_Buf
    Erase m_NeedComma
End Sub

' resets write cursor without freeing allocated buffer capacity
Public Sub Clear()
    m_Pos = 0
    m_Depth = 0
    m_ExpectValue = False
    If m_DepthCap > 0 Then m_NeedComma(0) = False
End Sub

Public Property Get Length() As Long
    Length = m_Pos
End Property

Private Sub EnsureCapacity(ByVal extraChars As Long)
    If m_Pos + extraChars >= m_Cap Then
        Do While m_Pos + extraChars >= m_Cap
            m_Cap = m_Cap * 2&
        Loop
        ReDim Preserve m_Buf(m_Cap - 1&)
    End If
End Sub

Private Sub PushScope()
    m_Depth = m_Depth + 1
    If m_Depth >= m_DepthCap Then
        m_DepthCap = m_DepthCap * 2&
        ReDim Preserve m_NeedComma(m_DepthCap - 1&)
    End If
    m_NeedComma(m_Depth) = False
End Sub

Private Sub PopScope()
    If m_Depth > 0 Then m_Depth = m_Depth - 1
End Sub

Private Sub WriteChar(ByVal ch As Integer)
    If m_Pos >= m_Cap Then EnsureCapacity 1&
    m_Buf(m_Pos) = ch
    m_Pos = m_Pos + 1
End Sub

Private Sub WriteRawString(ByRef sText As String)
    Dim lLen As Long
    lLen = Len(sText)
    If lLen = 0 Then Exit Sub
    If m_Pos + lLen >= m_Cap Then EnsureCapacity lLen
    CopyMemory ByVal VarPtr(m_Buf(m_Pos)), ByVal StrPtr(sText), lLen * 2&
    m_Pos = m_Pos + lLen
End Sub

' --- STREAMING BUILDER API ---

Public Sub BeginObject()
    If m_NeedComma(m_Depth) Then WriteChar 44    ' ','
    WriteChar 123                               ' '{'
    m_NeedComma(m_Depth) = True
    PushScope
    m_ExpectValue = False
End Sub

Public Sub EndObject()
    If m_ExpectValue Then
        Err.Raise vbObjectError + 5024, "JsonWrite", "EndObject called before writing value for key"
        Exit Sub
    End If
    PopScope
    WriteChar 125                               ' '}'
End Sub

Public Sub BeginArray()
    If m_NeedComma(m_Depth) Then WriteChar 44    ' ','
    WriteChar 91                                ' '['
    m_NeedComma(m_Depth) = True
    PushScope
    m_ExpectValue = False

End Sub

Public Sub EndArray()
    PopScope
    WriteChar 93                                ' ']'
End Sub

' writes a "key": prefix inside an active object
Public Sub WriteKey(ByVal sKey As String)
    If m_ExpectValue Then
        Err.Raise vbObjectError + 5024, "JsonWrite", "Missing value for previous key"
        Exit Sub
    End If
    If m_NeedComma(m_Depth) Then WriteChar 44    ' ','
    WriteEscapedString sKey
    WriteChar 58                                ' ':'
    m_NeedComma(m_Depth) = False
    m_ExpectValue = True
End Sub

' adds a complete key/value pair inside an active object
Public Sub Add(ByVal sKey As String, ByVal vValue As Variant)
    WriteKey sKey
    WriteVariantValue vValue
    m_NeedComma(m_Depth) = True
    m_ExpectValue = False
End Sub

' adds a standalone value inside an active array (or root)
Public Sub AddValue(ByVal vValue As Variant)
    If m_NeedComma(m_Depth) Then WriteChar 44    ' ','
    WriteVariantValue vValue
    m_NeedComma(m_Depth) = True
End Sub


' --- DIRECT-BUFFER STRING & SCALAR WRITERS ---

' writes a double-quoted, json-escaped string directly into m_Buf without temp strings
Private Sub WriteEscapedString(ByRef sText As String)
    Dim lLen As Long
    Dim inArr() As Integer
    Dim i As Long, j As Long
    Dim charCode As Long, nextCode As Long
    Dim sHex As String

    lLen = Len(sText)
    If lLen = 0 Then
        EnsureCapacity 2&
        m_Buf(m_Pos) = 34
        m_Pos = m_Pos + 1
        m_Buf(m_Pos) = 34
        m_Pos = m_Pos + 1
        Exit Sub
    End If

    ' worst-case capacity: 2 quotes + 6 chars per input char
    EnsureCapacity (lLen * 6&) + 2&

    ReDim inArr(lLen - 1&)
    CopyMemory ByVal VarPtr(inArr(0)), ByVal StrPtr(sText), lLen * 2&

    m_Buf(m_Pos) = 34    ' '"'
    m_Pos = m_Pos + 1

    For i = 0 To lLen - 1&
        charCode = inArr(i)
        If charCode < 0 Then charCode = charCode + 65536

        ' pass valid utf-16 surrogate pairs through directly
        If charCode >= &HD800 And charCode <= &HDBFF And i < lLen - 1& Then
            nextCode = inArr(i + 1&)
            If nextCode < 0 Then nextCode = nextCode + 65536
            If nextCode >= &HDC00 And nextCode <= &HDFFF Then
                m_Buf(m_Pos) = inArr(i)
                m_Pos = m_Pos + 1
                m_Buf(m_Pos) = inArr(i + 1&)
                m_Pos = m_Pos + 1
                i = i + 1&
                GoTo NextChar
            End If
        End If

        Select Case charCode
            Case 34    ' \"
                m_Buf(m_Pos) = 92
                m_Pos = m_Pos + 1
                m_Buf(m_Pos) = 34
                m_Pos = m_Pos + 1

            Case 92    ' \\
                m_Buf(m_Pos) = 92
                m_Pos = m_Pos + 1
                m_Buf(m_Pos) = 92
                m_Pos = m_Pos + 1

            Case 47    ' \/
                m_Buf(m_Pos) = 92
                m_Pos = m_Pos + 1
                m_Buf(m_Pos) = 47
                m_Pos = m_Pos + 1

            Case 8     ' \b
                m_Buf(m_Pos) = 92
                m_Pos = m_Pos + 1
                m_Buf(m_Pos) = 98
                m_Pos = m_Pos + 1

            Case 12    ' \f
                m_Buf(m_Pos) = 92
                m_Pos = m_Pos + 1
                m_Buf(m_Pos) = 102
                m_Pos = m_Pos + 1

            Case 10    ' \n
                m_Buf(m_Pos) = 92
                m_Pos = m_Pos + 1
                m_Buf(m_Pos) = 110
                m_Pos = m_Pos + 1

            Case 13    ' \r
                m_Buf(m_Pos) = 92
                m_Pos = m_Pos + 1
                m_Buf(m_Pos) = 114
                m_Pos = m_Pos + 1

            Case 9     ' \t
                m_Buf(m_Pos) = 92
                m_Pos = m_Pos + 1
                m_Buf(m_Pos) = 116
                m_Pos = m_Pos + 1

            Case 0 To 31, 127 To 159, &HD800 To &HDFFF
                m_Buf(m_Pos) = 92
                m_Pos = m_Pos + 1
                m_Buf(m_Pos) = 117
                m_Pos = m_Pos + 1
                sHex = Right$("0000" & Hex$(charCode), 4)
                For j = 1 To 4
                    m_Buf(m_Pos) = AscW(Mid$(sHex, j, 1))
                    m_Pos = m_Pos + 1
                Next j
            Case Else
                m_Buf(m_Pos) = inArr(i)
                m_Pos = m_Pos + 1
        End Select
NextChar:
    Next i

    m_Buf(m_Pos) = 34    ' '"'
    m_Pos = m_Pos + 1
End Sub

Private Sub WriteBoolean(ByVal bVal As Boolean)
    If bVal Then
        EnsureCapacity 4&
        m_Buf(m_Pos) = 116    ' t
         m_Pos = m_Pos + 1
        m_Buf(m_Pos) = 114    ' r
         m_Pos = m_Pos + 1
        m_Buf(m_Pos) = 117    ' u
         m_Pos = m_Pos + 1
        m_Buf(m_Pos) = 101   ' e
         m_Pos = m_Pos + 1
    Else
        EnsureCapacity 5&
        m_Buf(m_Pos) = 102   ' f
         m_Pos = m_Pos + 1
        m_Buf(m_Pos) = 97     ' a
         m_Pos = m_Pos + 1
        m_Buf(m_Pos) = 108    ' l
         m_Pos = m_Pos + 1
        m_Buf(m_Pos) = 115    ' s
         m_Pos = m_Pos + 1
        m_Buf(m_Pos) = 101    ' e
         m_Pos = m_Pos + 1
    End If
End Sub

Private Sub WriteNull()
    EnsureCapacity 4&
    m_Buf(m_Pos) = 110       ' n
     m_Pos = m_Pos + 1
    m_Buf(m_Pos) = 117        ' u
     m_Pos = m_Pos + 1
    m_Buf(m_Pos) = 108        ' l
     m_Pos = m_Pos + 1
    m_Buf(m_Pos) = 108        ' l
     m_Pos = m_Pos + 1
End Sub

' Str$ is locale-invariant (always uses '.' decimal separator and adds leading space for positive numbers)
Private Sub WriteNumber(ByRef vNum As Variant)
    Dim sNum As String
    ' guard against IEEE 754 NaN (v <> v) or Infinity (1.#INF)
    If VarType(vNum) = vbDouble Or VarType(vNum) = vbSingle Then
        If vNum <> vNum Or Abs(vNum) > 1.79769313486231E+308 Then
            WriteNull
            Exit Sub
        End If
    End If
    sNum = Trim$(Str$(vNum))
    If Left$(sNum, 1) = "." Then
        sNum = "0" & sNum
    ElseIf Left$(sNum, 2) = "-." Then
        sNum = "-0." & Mid$(sNum, 3)
    End If
    WriteRawString sNum
End Sub


' converts a local vb6 Date to an ISO 8601 string (defaults to UTC "YYYY-MM-DDTHH:MM:SSZ")
Public Function LocalToIsoDate(ByVal dtLocalDate As Date, Optional ByVal blnUseLocalOffset As Boolean = False) As String
    Dim tzi As TIME_ZONE_INFORMATION
    Dim lngBias As Long
    Dim lngResult As Long
    Dim dtUtc As Date

    lngResult = GetTimeZoneInformation(tzi)
    If lngResult <> 0& Then
        lngBias = tzi.Bias
        If lngResult = 2& Then lngBias = lngBias + tzi.DaylightBias
    End If

    If blnUseLocalOffset Then
        LocalToIsoDate = Format$(dtLocalDate, "yyyy-mm-ddThh:nn:ss") & _
                         Format$(-lngBias / 60, "\+00;-00") & _
                         Format$(Abs(lngBias) Mod 60, ":00")
    Else
        dtUtc = DateAdd("n", lngBias, dtLocalDate)
        LocalToIsoDate = Format$(dtUtc, "yyyy-mm-ddThh:nn:ss") & "Z"
    End If
End Function



' --- HYBRID VARIANT SERIALIZER & OUTPUT EXTRACTORS ---

' serializes a Scripting.Dictionary, 2D key/value array, 1D array, or scalar into the buffer
Public Function Serialize(vData As Variant) As String
    Clear
    WriteVariantValue vData
    Serialize = ToString()
End Function

' appends any supported Variant type into the active write buffer
Public Sub WriteVariantValue(vData As Variant)
    Dim vType As Integer
    vType = VarType(vData)

    If (vType And vbArray) = vbArray Then
        WriteArrayValue vData
        Exit Sub
    End If

    Select Case vType
        Case vbString
            WriteEscapedString CStr(vData)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbByte, vbDecimal
            WriteNumber vData
        Case vbBoolean
            WriteBoolean CBool(vData)
        Case vbDate
            WriteEscapedString LocalToIsoDate(CDate(vData), False)
        Case vbNull, vbEmpty
            WriteNull
        Case vbObject
            If vData Is Nothing Then
                WriteNull
            ElseIf TypeName(vData) = "Dictionary" Then
                WriteDictionary vData
            Else
                Err.Raise vbObjectError + 5020, "JsonWrite", "Unsupported object type: " & TypeName(vData)
            End If
        Case Else
            Err.Raise vbObjectError + 5021, "JsonWrite", "Unsupported Variant type: " & vType
    End Select

    m_ExpectValue = False

End Sub

' serializes a late-bound Scripting.Dictionary as a json object (no scrrun.dll reference needed)
Private Sub WriteDictionary(ByVal objDict As Object)
    Dim vKeys As Variant
    Dim vItems As Variant
    Dim i As Long
    Dim lCount As Long

    lCount = objDict.Count
    WriteChar 123    ' '{'
    If lCount > 0 Then
        vKeys = objDict.keys
        vItems = objDict.Items
        For i = 0 To lCount - 1&
            If i > 0 Then WriteChar 44    ' ','
            WriteEscapedString CStr(vKeys(i))
            WriteChar 58                  ' ':'
            WriteVariantValue vItems(i)
        Next i
    End If
    WriteChar 125    ' '}'
End Sub

' handles 1D arrays as json arrays and 2D (0 to 1, 0 to N) arrays as json objects
Private Sub WriteArrayValue(vArr As Variant)
    Dim lDims As Long
    Dim i As Long
    Dim lLow1 As Long, lHigh1 As Long
    Dim lLow2 As Long, lHigh2 As Long

    lDims = GetArrayDimensions(vArr)

    If lDims = 0 Then
        ' uninitialized array -> []
        WriteChar 91
        WriteChar 93
    ElseIf lDims = 1 Then
        lLow1 = LBound(vArr, 1)
        lHigh1 = UBound(vArr, 1)
        WriteChar 91    ' '['
        For i = lLow1 To lHigh1
            If i > lLow1 Then WriteChar 44    ' ','
            WriteVariantValue vArr(i)
        Next i
        WriteChar 93    ' ']'
    ElseIf lDims = 2 Then
        lLow1 = LBound(vArr, 1)
        lHigh1 = UBound(vArr, 1)
        If (lHigh1 - lLow1) <> 1 Then
            Err.Raise vbObjectError + 5022, "JsonWrite", "2D object array first dimension must have 2 elements (key, value)"
        End If
        lLow2 = LBound(vArr, 2)
        lHigh2 = UBound(vArr, 2)
        WriteChar 123    ' '{'
        For i = lLow2 To lHigh2
            If i > lLow2 Then WriteChar 44    ' ','
            WriteEscapedString CStr(vArr(lLow1, i))
            WriteChar 58                      ' ':'
            WriteVariantValue vArr(lLow1 + 1&, i)
        Next i
        WriteChar 125    ' '}'
    Else
        Err.Raise vbObjectError + 5023, "JsonWrite", "Arrays with more than 2 dimensions are not supported"
    End If
End Sub

' reads cDims (first 2 bytes of the SAFEARRAY header) directly without triggering On Error traps
Private Function GetArrayDimensions(ByVal vArr As Variant) As Long
    Dim vt As Integer
    Dim pSafeArray As Long
    Dim cDims As Integer

    CopyMemory vt, ByVal VarPtr(vArr), 2&
    If (vt And vbArray) = 0 Then Exit Function

    CopyMemory pSafeArray, ByVal VarPtr(vArr) + 8&, 4&
    If (vt And &H4000) <> 0 And pSafeArray <> 0 Then    ' VT_BYREF
        CopyMemory pSafeArray, ByVal pSafeArray, 4&
    End If

    If pSafeArray <> 0 Then
        CopyMemory cDims, ByVal pSafeArray, 2&
        GetArrayDimensions = cDims
    End If
End Function

' extracts the accumulated json buffer as a vb6 string in a single CopyMemory call
Public Function ToString() As String
    If m_Pos <= 0 Then Exit Function
    ToString = Space$(m_Pos)
    CopyMemory ByVal StrPtr(ToString), ByVal VarPtr(m_Buf(0)), m_Pos * 2&
End Function

' converts the accumulated utf-16 buffer directly into a utf-8 Byte() array (ideal for files/sockets)
Public Function ToUtf8Bytes() As Byte()
    Dim lSize As Long
    Dim outBytes() As Byte

    If m_Pos <= 0 Then Exit Function
    lSize = WideCharToMultiByte(CP_UTF8, 0&, ByVal VarPtr(m_Buf(0)), m_Pos, 0&, 0&, 0&, 0&)
    If lSize > 0 Then
        ReDim outBytes(lSize - 1&)
        Call WideCharToMultiByte(CP_UTF8, 0&, ByVal VarPtr(m_Buf(0)), m_Pos, ByVal VarPtr(outBytes(0)), lSize, 0&, 0&)
        ToUtf8Bytes = outBytes
    End If
End Function






