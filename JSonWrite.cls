VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSonWrite"
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

Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLbound As Long
End Type

Private Declare Function CreateFileW Lib "kernel32" ( _
    ByVal lpFileName As Long, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByRef lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) As Long

Private Declare Function DispCallFunc Lib "oleaut32" ( _
                                      ByVal pvInstance As Long, _
                                      ByVal oVft As Long, _
                                      ByVal cc As Long, _
                                      ByVal vtReturn As Integer, _
                                      ByVal cActuals As Long, _
                                      ByRef prgvt As Integer, _
                                      ByRef prgpvarg As Long, _
                                      ByRef pvargResult As Variant) As Long
                                      
                                      
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Const GENERIC_WRITE As Long = &H40000000
Private Const CREATE_ALWAYS As Long = 2&
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80&
Private Const INVALID_HANDLE_VALUE As Long = -1&

Private Const CP_UTF8 As Long = 65001
Private Const BUFFER_INITIAL As Long = 32768

Private Const CC_STDCALL As Long = 4&
Private Const VTBL_DICT_COUNT As Long = 44&   ' Slot 11: IDictionary::get_Count
Private Const VTBL_DICT_ITEMS As Long = 52&   ' Slot 13: IDictionary::Items
Private Const VTBL_DICT_KEYS As Long = 60&  ' Slot 15: IDictionary::Keys

' Flat UTF-16 integer output buffer (zero BSTR concatenation churn)
Private m_Buf() As Integer
Private m_Pos As Long
Private m_Cap As Long

' Container nesting and comma state
Private m_NeedComma() As Boolean
Private m_ScopeType() As Byte
Private m_Depth As Long
Private m_DepthCap As Long
Private m_ExpectValue As Boolean

' Reusable scratch buffers
Private m_CharIn() As Integer
Private m_CharInCap As Long
Private m_HexChars(0 To 15) As Integer
Private m_NumBuf(0 To 23) As Integer

' Zero-copy 1D Variant view over 1D/2D Variant SAFEARRAYs
Private m_ViewSA As SAFEARRAY1D
Private m_ViewArr() As Variant
Private m_pViewSA As Long

' Cached Scripting.Dictionary vtable pointer
Private m_DictVTable As Long

' Cached time zone bias and preformatted offset suffix
Private m_TzBias As Long
Private m_TzOffsetSuffix As String
Private m_TzCached As Boolean


Private Sub Class_Initialize()
    Dim i As Long

    m_Cap = BUFFER_INITIAL
    ReDim m_Buf(m_Cap - 1&)
    m_Pos = 0&

    m_CharInCap = 1024&
    ReDim m_CharIn(m_CharInCap - 1&)

    m_DepthCap = 64&
    ReDim m_NeedComma(m_DepthCap - 1&)
    ReDim m_ScopeType(m_DepthCap - 1&)
    m_Depth = 0&

    ' Populate hex lookup table ('0'-'9' = 48-57, 'A'-'F' = 65-70)
    For i = 0 To 9
        m_HexChars(i) = 48 + i
    Next i
    For i = 10 To 15
        m_HexChars(i) = 55 + i
    Next i

    ' Initialize zero-copy SAFEARRAY1D view
    m_ViewSA.cDims = 1
    m_ViewSA.fFeatures = &H10 Or &H80   ' FADF_FIXEDSIZE Or FADF_STATIC
    m_ViewSA.cbElements = 16&          ' sizeof(Variant) = 16 bytes
    m_pViewSA = VarPtr(m_ViewSA)
    CopyMemory ByVal ArrPtr(m_ViewArr), m_pViewSA, 4&
End Sub

Private Sub Class_Terminate()
    ' Detach SAFEARRAY view before VB6 tears down module variables
    m_ViewSA.pvData = 0&
    m_ViewSA.cElements = 0&
    CopyMemory ByVal ArrPtr(m_ViewArr), 0&, 4&

    Erase m_Buf
    Erase m_CharIn
    Erase m_NeedComma
    Erase m_ScopeType
End Sub

' Resets write cursor without freeing allocated buffer capacity
Public Sub Clear()
    m_Pos = 0&
    m_Depth = 0&
    m_ExpectValue = False
    If m_DepthCap > 0& Then
        m_NeedComma(0) = False
        m_ScopeType(0) = 0
    End If
End Sub

Public Property Get Length() As Long
    Length = m_Pos
End Property

' Default property: jw("key") = value (inside object) or jw("") = value (inside array/root)
Public Property Let Item(ByVal sKey As String, ByVal vValue As Variant)
Attribute Item.VB_UserMemId = 0
    If m_ScopeType(m_Depth) = 1 Then
        Add sKey, vValue
    Else
        AddValue vValue
    End If
End Property

' Supports object values (such as Scripting.Dictionary) via Set jw("key") = dict
Public Property Set Item(ByVal sKey As String, ByVal objValue As Object)
    If m_ScopeType(m_Depth) = 1 Then
        Add sKey, objValue
    Else
        AddValue objValue
    End If
End Property

Private Sub EnsureCapacity(ByVal extraChars As Long)
    If m_Pos + extraChars >= m_Cap Then
        Do While m_Pos + extraChars >= m_Cap
            m_Cap = m_Cap * 2&
        Loop
        ReDim Preserve m_Buf(m_Cap - 1&)
    End If
End Sub

Private Sub WriteRawString(ByRef sText As String)
    Dim lLen As Long
    lLen = Len(sText)
    If lLen = 0& Then Exit Sub
    If m_Pos + lLen >= m_Cap Then EnsureCapacity lLen
    CopyMemory ByVal VarPtr(m_Buf(m_Pos)), ByVal StrPtr(sText), lLen * 2&
    m_Pos = m_Pos + lLen
End Sub


' --- STREAMING BUILDER API ---

Public Sub BeginObject()
    If m_Pos + 2& >= m_Cap Then EnsureCapacity 2&
    If m_NeedComma(m_Depth) Then
        m_Buf(m_Pos) = 44    ' ','
        m_Pos = m_Pos + 1&
    End If
    m_Buf(m_Pos) = 123       ' '{'
    m_Pos = m_Pos + 1&
    m_NeedComma(m_Depth) = True

    m_Depth = m_Depth + 1&
    If m_Depth >= m_DepthCap Then
        m_DepthCap = m_DepthCap * 2&
        ReDim Preserve m_NeedComma(m_DepthCap - 1&)
        ReDim Preserve m_ScopeType(m_DepthCap - 1&)
    End If
    m_NeedComma(m_Depth) = False
    m_ScopeType(m_Depth) = 1
    m_ExpectValue = False
End Sub

Public Sub EndObject()
    If m_ExpectValue Then Err.Raise vbObjectError + 5024, "JsonWrite", "EndObject called before writing value for key"
    If m_Depth > 0& Then m_Depth = m_Depth - 1&
    If m_Pos >= m_Cap Then EnsureCapacity 1&
    m_Buf(m_Pos) = 125       ' '}'
    m_Pos = m_Pos + 1&
End Sub

Public Sub BeginArray()
    If m_Pos + 2& >= m_Cap Then EnsureCapacity 2&
    If m_NeedComma(m_Depth) Then
        m_Buf(m_Pos) = 44    ' ','
        m_Pos = m_Pos + 1&
    End If
    m_Buf(m_Pos) = 91        ' '['
    m_Pos = m_Pos + 1&
    m_NeedComma(m_Depth) = True

    m_Depth = m_Depth + 1&
    If m_Depth >= m_DepthCap Then
        m_DepthCap = m_DepthCap * 2&
        ReDim Preserve m_NeedComma(m_DepthCap - 1&)
        ReDim Preserve m_ScopeType(m_DepthCap - 1&)
    End If
    m_NeedComma(m_Depth) = False
    m_ScopeType(m_Depth) = 2
    m_ExpectValue = False
End Sub

Public Sub EndArray()
    If m_ExpectValue Then Err.Raise vbObjectError + 5024, "JsonWrite", "EndArray called before writing value for key"
    If m_Depth > 0& Then m_Depth = m_Depth - 1&
    If m_Pos >= m_Cap Then EnsureCapacity 1&
    m_Buf(m_Pos) = 93        ' ']'
    m_Pos = m_Pos + 1&
End Sub

' Writes optional leading comma + escaped "key": prefix
Private Sub WriteKeyPrefix(ByRef sKey As String)
    If m_ExpectValue Then Err.Raise vbObjectError + 5024, "JsonWrite", "Missing value for previous key"
    If m_NeedComma(m_Depth) Then
        m_Buf(m_Pos) = 44    ' ','
        m_Pos = m_Pos + 1&
    End If
    WriteEscapedString sKey
    m_Buf(m_Pos) = 58        ' ':'
    m_Pos = m_Pos + 1&
End Sub

Public Sub WriteKey(ByRef sKey As String)
    WriteKeyPrefix sKey
    m_NeedComma(m_Depth) = False
    m_ExpectValue = True
End Sub

Public Sub Add(ByRef sKey As String, ByRef vValue As Variant)
    WriteKeyPrefix sKey
    WriteVariantValue vValue
    m_NeedComma(m_Depth) = True
    m_ExpectValue = False
End Sub

Public Sub AddString(ByRef sKey As String, ByRef sValue As String)
    WriteKeyPrefix sKey
    WriteEscapedString sValue
    m_NeedComma(m_Depth) = True
End Sub

Public Sub AddLong(ByRef sKey As String, ByVal lValue As Long)
    WriteKeyPrefix sKey
    WriteLongDirect lValue
    m_NeedComma(m_Depth) = True
End Sub

Public Sub AddBoolean(ByRef sKey As String, ByVal bValue As Boolean)
    WriteKeyPrefix sKey
    WriteBoolean bValue
    m_NeedComma(m_Depth) = True
End Sub

Public Sub AddNull(ByRef sKey As String)
    WriteKeyPrefix sKey
    WriteNull
    m_NeedComma(m_Depth) = True
End Sub

' Adds a standalone value inside an active array (or root)
Public Sub AddValue(ByRef vValue As Variant)
    If m_NeedComma(m_Depth) Then
        If m_Pos >= m_Cap Then EnsureCapacity 1&
        m_Buf(m_Pos) = 44    ' ','
        m_Pos = m_Pos + 1&
    End If
    WriteVariantValue vValue
    m_NeedComma(m_Depth) = True
End Sub


' --- DIRECT-BUFFER STRING & SCALAR WRITERS ---

Private Sub WriteEscapedStringByVal(ByVal sText As String)
    WriteEscapedString sText
End Sub

' Writes a double-quoted, JSON-escaped string directly into m_Buf without temp strings
Private Sub WriteEscapedString(ByRef sText As String)
    Dim lLen As Long
    Dim i As Long
    Dim charCode As Long, nextCode As Long

    lLen = Len(sText)
    If lLen = 0& Then
        ' +4 reserves 2 quotes + 2 spare chars for delimiter/colon
        If m_Pos + 4& >= m_Cap Then EnsureCapacity 4&
        m_Buf(m_Pos) = 34
        m_Buf(m_Pos + 1&) = 34
        m_Pos = m_Pos + 2&
        Exit Sub
    End If

    ' Worst case: 2 quotes + 6 chars per input char + 2 spare chars
    If m_Pos + (lLen * 6&) + 4& >= m_Cap Then EnsureCapacity (lLen * 6&) + 4&
    If lLen > m_CharInCap Then
        Do While lLen > m_CharInCap
            m_CharInCap = m_CharInCap * 2&
        Loop
        ReDim m_CharIn(m_CharInCap - 1&)
    End If
    CopyMemory ByVal VarPtr(m_CharIn(0)), ByVal StrPtr(sText), lLen * 2&
    m_Buf(m_Pos) = 34    ' '"'
    m_Pos = m_Pos + 1&
    For i = 0& To lLen - 1&
        charCode = m_CharIn(i)
        If charCode < 0& Then charCode = charCode + 65536
        ' Pass valid UTF-16 surrogate pairs through directly
        If charCode >= &HD800& And charCode <= &HDBFF& And i < lLen - 1& Then
            nextCode = m_CharIn(i + 1&)
            If nextCode < 0& Then nextCode = nextCode + 65536
            If nextCode >= &HDC00& And nextCode <= &HDFFF& Then
                m_Buf(m_Pos) = m_CharIn(i)
                m_Buf(m_Pos + 1&) = m_CharIn(i + 1&)
                m_Pos = m_Pos + 2&
                i = i + 1&
                GoTo NextChar
            End If
        End If
        Select Case charCode
            Case 34    ' "
                m_Buf(m_Pos) = 92
                m_Buf(m_Pos + 1&) = 34
                m_Pos = m_Pos + 2&
            Case 92    ' \
                m_Buf(m_Pos) = 92
                m_Buf(m_Pos + 1&) = 92
                m_Pos = m_Pos + 2&
            Case 47    ' /
                m_Buf(m_Pos) = 92
                m_Buf(m_Pos + 1&) = 47
                m_Pos = m_Pos + 2&
            Case 8     ' \b
                m_Buf(m_Pos) = 92
                m_Buf(m_Pos + 1&) = 98
                m_Pos = m_Pos + 2&
            Case 12    ' \f
                m_Buf(m_Pos) = 92
                m_Buf(m_Pos + 1&) = 102
                m_Pos = m_Pos + 2&
            Case 10    ' \n
                m_Buf(m_Pos) = 92
                m_Buf(m_Pos + 1&) = 110
                m_Pos = m_Pos + 2&
            Case 13    ' \r
                m_Buf(m_Pos) = 92
                m_Buf(m_Pos + 1&) = 114
                m_Pos = m_Pos + 2&
            Case 9     ' \t
                m_Buf(m_Pos) = 92
                m_Buf(m_Pos + 1&) = 116
                m_Pos = m_Pos + 2&
            Case 0 To 31, 127 To 159, &HD800& To &HDFFF&
                m_Buf(m_Pos) = 92       ' ''
                m_Buf(m_Pos + 1&) = 117    ' 'u'
                m_Buf(m_Pos + 2&) = m_HexChars((charCode And &HF000&) \ &H1000&)
                m_Buf(m_Pos + 3&) = m_HexChars((charCode And &HF00&) \ &H100&)
                m_Buf(m_Pos + 4&) = m_HexChars((charCode And &HF0&) \ &H10&)
                m_Buf(m_Pos + 5&) = m_HexChars(charCode And &HF&)
                m_Pos = m_Pos + 6&
            Case Else
                m_Buf(m_Pos) = m_CharIn(i)
                m_Pos = m_Pos + 1&
        End Select
NextChar:
    Next i
    m_Buf(m_Pos) = 34    ' '"'
    m_Pos = m_Pos + 1&
End Sub


Private Sub WriteBoolean(ByVal bVal As Boolean)
    If m_Pos + 5& >= m_Cap Then EnsureCapacity 5&
    If bVal Then
        m_Buf(m_Pos) = 116      ' t
        m_Buf(m_Pos + 1&) = 114    ' r
        m_Buf(m_Pos + 2&) = 117    ' u
        m_Buf(m_Pos + 3&) = 101    ' e
        m_Pos = m_Pos + 4&
    Else
        m_Buf(m_Pos) = 102      ' f
        m_Buf(m_Pos + 1&) = 97  ' a
        m_Buf(m_Pos + 2&) = 108    ' l
        m_Buf(m_Pos + 3&) = 115    ' s
        m_Buf(m_Pos + 4&) = 101    ' e
        m_Pos = m_Pos + 5&
    End If
End Sub

Private Sub WriteNull()
    If m_Pos + 4& >= m_Cap Then EnsureCapacity 4&
    m_Buf(m_Pos) = 110          ' n
    m_Buf(m_Pos + 1&) = 117     ' u
    m_Buf(m_Pos + 2&) = 108     ' l
    m_Buf(m_Pos + 3&) = 108     ' l
    m_Pos = m_Pos + 4&
End Sub

Private Sub WriteLongDirect(ByVal lVal As Long)
    Dim idx As Long
    Dim bNeg As Boolean
    If m_Pos + 12& >= m_Cap Then EnsureCapacity 12&
    If lVal = 0& Then
        m_Buf(m_Pos) = 48    ' '0'
        m_Pos = m_Pos + 1&
        Exit Sub
    End If
    If lVal = &H80000000 Then
        WriteRawString "-2147483648"
        Exit Sub
    End If
    If lVal < 0& Then
        bNeg = True
        lVal = -lVal
    End If
    idx = 0&
    Do While lVal > 0&
        m_NumBuf(idx) = 48 + (lVal Mod 10&)
        lVal = lVal \ 10&
        idx = idx + 1&
    Loop
    If bNeg Then
        m_Buf(m_Pos) = 45    ' '-'
        m_Pos = m_Pos + 1&
    End If
    Do While idx > 0&
        idx = idx - 1&
        m_Buf(m_Pos) = m_NumBuf(idx)
        m_Pos = m_Pos + 1&
    Loop
End Sub


Private Sub WriteNumber(ByRef vNum As Variant)
    Dim vt As Integer
    Dim sNum As String
    vt = VarType(vNum)
    If vt = vbLong Or vt = vbInteger Or vt = vbByte Then
        WriteLongDirect CLng(vNum)
        Exit Sub
    End If
    ' Guard against IEEE 754 NaN (v <> v) or Infinity
    If vt = vbDouble Or vt = vbSingle Then
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


Public Sub RefreshTimeZone()
    Dim tzi As TIME_ZONE_INFORMATION
    Dim lngResult As Long
    Dim lAbsBias As Long
    Dim sSign As String
    lngResult = GetTimeZoneInformation(tzi)
    If lngResult <> 0& Then
        m_TzBias = tzi.Bias
        If lngResult = 2& Then m_TzBias = m_TzBias + tzi.DaylightBias
    Else
        m_TzBias = 0&
    End If
    If m_TzBias <= 0& Then
        sSign = "+"
    Else
        sSign = "-"
    End If
    lAbsBias = Abs(m_TzBias)
    m_TzOffsetSuffix = sSign & Right$("00" & CStr(lAbsBias \ 60&), 2) & ":" & Right$("00" & CStr(lAbsBias Mod 60&), 2)
    m_TzCached = True
End Sub


' Converts a local VB6 Date to an ISO 8601 string (defaults to UTC "YYYY-MM-DDTHH:MM:SSZ")
Public Function LocalToIsoDate(ByVal dtLocalDate As Date, Optional ByVal blnUseLocalOffset As Boolean = False) As String
    Dim dtUtc As Date
    If Not m_TzCached Then RefreshTimeZone
    If blnUseLocalOffset Then
        LocalToIsoDate = Format$(dtLocalDate, "yyyy-mm-ddThh:nn:ss") & m_TzOffsetSuffix
    Else
        dtUtc = DateAdd("n", m_TzBias, dtLocalDate)
        LocalToIsoDate = Format$(dtUtc, "yyyy-mm-ddThh:nn:ss") & "Z"
    End If
End Function


Private Sub WriteIsoDate(ByVal dtLocalDate As Date)
    Dim dtUtc As Date
    Dim lYear As Long, lMonth As Long, lDay As Long
    Dim lHour As Long, lMin As Long, lSec As Long
    If Not m_TzCached Then RefreshTimeZone
    dtUtc = DateAdd("n", m_TzBias, dtLocalDate)
    lYear = Year(dtUtc)
    lMonth = Month(dtUtc)
    lDay = Day(dtUtc)
    lHour = Hour(dtUtc)
    lMin = Minute(dtUtc)
    lSec = Second(dtUtc)
    ' Reserve 22 chars: "YYYY-MM-DDTHH:MM:SSZ" + 2 spare chars for delimiters
    If m_Pos + 24& >= m_Cap Then EnsureCapacity 24&
    m_Buf(m_Pos) = 34                                       ' '"'
    m_Buf(m_Pos + 1&) = 48 + (lYear \ 1000&)
    m_Buf(m_Pos + 2&) = 48 + ((lYear \ 100&) Mod 10&)
    m_Buf(m_Pos + 3&) = 48 + ((lYear \ 10&) Mod 10&)
    m_Buf(m_Pos + 4&) = 48 + (lYear Mod 10&)
    m_Buf(m_Pos + 5&) = 45                                  ' '-'
    m_Buf(m_Pos + 6&) = 48 + (lMonth \ 10&)
    m_Buf(m_Pos + 7&) = 48 + (lMonth Mod 10&)
    m_Buf(m_Pos + 8&) = 45                                  ' '-'
    m_Buf(m_Pos + 9&) = 48 + (lDay \ 10&)
    m_Buf(m_Pos + 10&) = 48 + (lDay Mod 10&)
    m_Buf(m_Pos + 11&) = 84                                 ' 'T'
    m_Buf(m_Pos + 12&) = 48 + (lHour \ 10&)
    m_Buf(m_Pos + 13&) = 48 + (lHour Mod 10&)
    m_Buf(m_Pos + 14&) = 58                                 ' ':'
    m_Buf(m_Pos + 15&) = 48 + (lMin \ 10&)
    m_Buf(m_Pos + 16&) = 48 + (lMin Mod 10&)
    m_Buf(m_Pos + 17&) = 58                                 ' ':'
    m_Buf(m_Pos + 18&) = 48 + (lSec \ 10&)
    m_Buf(m_Pos + 19&) = 48 + (lSec Mod 10&)
    m_Buf(m_Pos + 20&) = 90                                 ' 'Z'
    m_Buf(m_Pos + 21&) = 34                                 ' '"'
    m_Pos = m_Pos + 22&
End Sub


' --- HYBRID VARIANT SERIALIZER & OUTPUT EXTRACTORS ---
' Serializes a Scripting.Dictionary, 2D key/value array, 1D array, or scalar into the buffer
Public Function Serialize(vData As Variant) As String
    Clear
    WriteVariantValue vData
    Serialize = ToString()
End Function


' Appends any supported Variant type into the active write buffer
Public Sub WriteVariantValue(vData As Variant)
    Dim vType As Integer
    Dim objRef As Object
    Dim pObj As Long
    Dim lVTable As Long
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
            WriteIsoDate CDate(vData)
        Case vbNull, vbEmpty
            WriteNull
        Case vbObject
            Set objRef = vData
            pObj = ObjPtr(objRef)
            If pObj = 0& Then
                WriteNull
            Else
                CopyMemory lVTable, ByVal pObj, 4&
                If lVTable = m_DictVTable And m_DictVTable <> 0& Then
                    WriteDictionary pObj
                ElseIf TypeName(objRef) = "Dictionary" Then
                    m_DictVTable = lVTable
                    WriteDictionary pObj
                Else
                    Err.Raise vbObjectError + 5020, "JsonWrite", "Unsupported object type: " & TypeName(objRef)
                End If
            End If
        Case Else
            Err.Raise vbObjectError + 5021, "JsonWrite", "Unsupported Variant type: " & vType
    End Select
    m_ExpectValue = False
End Sub


' Serializes a late-bound Scripting.Dictionary via cached vtable + DispCallFunc
Private Sub WriteDictionary(ByVal pDict As Long)
    Dim lCount As Long
    Dim vKeys As Variant
    Dim vItems As Variant
    Dim vArg As Variant
    Dim pArgVar As Long
    Dim vtArg As Integer
    Dim hr As Long
    Dim vRet As Variant
    pArgVar = VarPtr(vArg)
    ' 1. Call IDictionary::get_Count(&lCount) via vtable offset 44
    vtArg = vbLong Or &H4000
    CopyMemory ByVal pArgVar, vtArg, 2&
    CopyMemory ByVal pArgVar + 8&, VarPtr(lCount), 4&
    hr = DispCallFunc(pDict, VTBL_DICT_COUNT, CC_STDCALL, vbLong, 1&, vtArg, pArgVar, vRet)
    CopyMemory ByVal pArgVar, 0, 2&
    If hr < 0& Then Err.Raise hr, "JsonWrite", "Dictionary get_Count failed"
    If lCount <= 0& Then
        If m_Pos + 2& >= m_Cap Then EnsureCapacity 2&
        m_Buf(m_Pos) = 123        ' '{'
        m_Buf(m_Pos + 1&) = 125   ' '}'
        m_Pos = m_Pos + 2&
        Exit Sub
    End If
    ' 2. Call IDictionary::Keys(&vKeys) and IDictionary::Items(&vItems) via vtable
    vtArg = vbVariant Or &H4000
    CopyMemory ByVal pArgVar, vtArg, 2&
    CopyMemory ByVal pArgVar + 8&, VarPtr(vKeys), 4&
    hr = DispCallFunc(pDict, VTBL_DICT_KEYS, CC_STDCALL, vbLong, 1&, vtArg, pArgVar, vRet)
    If hr < 0& Then
        CopyMemory ByVal pArgVar, 0, 2&
        Err.Raise hr, "JsonWrite", "Dictionary Keys failed"
    End If
    CopyMemory ByVal pArgVar + 8&, VarPtr(vItems), 4&
    hr = DispCallFunc(pDict, VTBL_DICT_ITEMS, CC_STDCALL, vbLong, 1&, vtArg, pArgVar, vRet)
    CopyMemory ByVal pArgVar, 0, 2&
    If hr < 0& Then Err.Raise hr, "JsonWrite", "Dictionary Items failed"
    ' 3. Walk returned Variant() SAFEARRAYs directly without SafeArrayGetElement
    Dim pSAKeys As Long, pSAItems As Long
    Dim pvKeys As Long, pvItems As Long
    Dim prevPvData As Long, prevCount As Long
    Dim i As Long
    CopyMemory pSAKeys, ByVal VarPtr(vKeys) + 8&, 4&
    CopyMemory pSAItems, ByVal VarPtr(vItems) + 8&, 4&
    CopyMemory pvKeys, ByVal pSAKeys + 12&, 4&
    CopyMemory pvItems, ByVal pSAItems + 12&, 4&
    If m_Pos + (lCount * 2&) + 2& >= m_Cap Then EnsureCapacity (lCount * 2&) + 2&
    m_Buf(m_Pos) = 123   ' '{'
    m_Pos = m_Pos + 1&
    prevPvData = m_ViewSA.pvData
    prevCount = m_ViewSA.cElements
    For i = 0& To lCount - 1&
        If i > 0& Then
            m_Buf(m_Pos) = 44    ' ','
            m_Pos = m_Pos + 1&
        End If
        m_ViewSA.pvData = pvKeys
        m_ViewSA.cElements = lCount
        If VarType(m_ViewArr(i)) = vbString Then
            WriteEscapedStringByVal m_ViewArr(i)
        Else
            WriteEscapedString CStr(m_ViewArr(i))
        End If
        m_Buf(m_Pos) = 58        ' ':'
        m_Pos = m_Pos + 1&
        m_ViewSA.pvData = pvItems
        m_ViewSA.cElements = lCount
        WriteVariantValue m_ViewArr(i)
    Next i
    m_ViewSA.pvData = prevPvData
    m_ViewSA.cElements = prevCount
    If m_Pos >= m_Cap Then EnsureCapacity 1&
    m_Buf(m_Pos) = 125   ' '}'
    m_Pos = m_Pos + 1&
End Sub


Private Sub WriteArrayValue(ByRef vArr As Variant)
    Dim vt As Integer
    Dim pSA As Long
    Dim cDims As Integer
    Dim cbElements As Long
    Dim pvData As Long
    Dim cElements1 As Long
    Dim cElements2 As Long
    Dim prevPvData As Long
    Dim prevCount As Long
    Dim totalVars As Long
    Dim i As Long
    CopyMemory vt, ByVal VarPtr(vArr), 2&
    CopyMemory pSA, ByVal VarPtr(vArr) + 8&, 4&
    If (vt And &H4000) <> 0 And pSA <> 0& Then    ' VT_BYREF
        CopyMemory pSA, ByVal pSA, 4&
    End If
    If pSA = 0& Then
        If m_Pos + 2& >= m_Cap Then EnsureCapacity 2&
        m_Buf(m_Pos) = 91        ' '['
        m_Buf(m_Pos + 1&) = 93   ' ']'
        m_Pos = m_Pos + 2&
        m_ExpectValue = False
        Exit Sub
    End If
    CopyMemory cDims, ByVal pSA, 2&
    CopyMemory cbElements, ByVal pSA + 4&, 4&
    CopyMemory pvData, ByVal pSA + 12&, 4&
    ' Fast-path for Variant arrays (vt = vbArray Or vbVariant)
    If (vt And &H1FFF) = (vbArray Or vbVariant) And cbElements = 16& Then
        prevPvData = m_ViewSA.pvData
        prevCount = m_ViewSA.cElements
        If cDims = 1 Then
            CopyMemory cElements1, ByVal pSA + 16&, 4&
            If cElements1 > 0& Then
                If m_Pos + cElements1 + 2& >= m_Cap Then EnsureCapacity cElements1 + 2&
            Else
                If m_Pos + 2& >= m_Cap Then EnsureCapacity 2&
            End If
            m_Buf(m_Pos) = 91    ' '['
            m_Pos = m_Pos + 1&
            If cElements1 > 0& And pvData <> 0& Then
                m_ViewSA.pvData = pvData
                m_ViewSA.cElements = cElements1
                WriteVariantValue m_ViewArr(0&)
                For i = 1& To cElements1 - 1&
                    m_Buf(m_Pos) = 44    ' ','
                    m_Pos = m_Pos + 1&
                    WriteVariantValue m_ViewArr(i)
                Next i
            End If
            If m_Pos >= m_Cap Then EnsureCapacity 1&
            m_Buf(m_Pos) = 93    ' ']'
            m_Pos = m_Pos + 1&
            m_ViewSA.pvData = prevPvData
            m_ViewSA.cElements = prevCount
            m_ExpectValue = False
            Exit Sub
        ElseIf cDims = 2 Then
            CopyMemory cElements2, ByVal pSA + 16&, 4&   ' Dim 2 count (N pairs)
            CopyMemory cElements1, ByVal pSA + 24&, 4&   ' Dim 1 count (must be 2: key, value)
            If cElements1 <> 2& Then
                Err.Raise vbObjectError + 5022, "JsonWrite", "2D object array first dimension must have 2 elements (key, value)"
            End If
            If cElements2 > 0& Then
                If m_Pos + (cElements2 * 2&) + 2& >= m_Cap Then EnsureCapacity (cElements2 * 2&) + 2&
            Else
                If m_Pos + 2& >= m_Cap Then EnsureCapacity 2&
            End If
            m_Buf(m_Pos) = 123   ' '{'
            m_Pos = m_Pos + 1&
            If cElements2 > 0& And pvData <> 0& Then
                totalVars = cElements2 * 2&
                m_ViewSA.pvData = pvData
                m_ViewSA.cElements = totalVars
                For i = 0& To totalVars - 2& Step 2&
                    If i > 0& Then
                        m_Buf(m_Pos) = 44    ' ','
                        m_Pos = m_Pos + 1&
                    End If
                    If VarType(m_ViewArr(i)) = vbString Then
                        WriteEscapedStringByVal m_ViewArr(i)
                    Else
                        WriteEscapedString CStr(m_ViewArr(i))
                    End If
                    m_Buf(m_Pos) = 58        ' ':'
                    m_Pos = m_Pos + 1&
                    WriteVariantValue m_ViewArr(i + 1&)
                Next i
            End If
            If m_Pos >= m_Cap Then EnsureCapacity 1&
            m_Buf(m_Pos) = 125   ' '}'
            m_Pos = m_Pos + 1&
            m_ViewSA.pvData = prevPvData
            m_ViewSA.cElements = prevCount
            m_ExpectValue = False
            Exit Sub
        End If
    End If
    ' Fallback for typed arrays (e.g., Long(), String(), Byte(), Double())
    WriteTypedArrayFallback vArr, cDims
    m_ExpectValue = False
End Sub


Private Sub WriteTypedArrayFallback(ByRef vArr As Variant, ByVal cDims As Integer)
    Dim i As Long
    Dim lLow1 As Long, lHigh1 As Long
    Dim lLow2 As Long, lHigh2 As Long
    Dim lValIdx As Long
    If cDims = 1 Then
        lLow1 = LBound(vArr, 1)
        lHigh1 = UBound(vArr, 1)
        If m_Pos + (lHigh1 - lLow1 + 3&) >= m_Cap Then EnsureCapacity (lHigh1 - lLow1 + 3&)
        m_Buf(m_Pos) = 91    ' '['
        m_Pos = m_Pos + 1&
        For i = lLow1 To lHigh1
            If i > lLow1 Then
                m_Buf(m_Pos) = 44
                m_Pos = m_Pos + 1&
            End If
            WriteVariantValue vArr(i)
        Next i
        If m_Pos >= m_Cap Then EnsureCapacity 1&
        m_Buf(m_Pos) = 93    ' ']'
        m_Pos = m_Pos + 1&
    ElseIf cDims = 2 Then
        lLow1 = LBound(vArr, 1)
        lHigh1 = UBound(vArr, 1)
        If (lHigh1 - lLow1) <> 1 Then
            Err.Raise vbObjectError + 5022, "JsonWrite", "2D object array first dimension must have 2 elements (key, value)"
        End If
        lLow2 = LBound(vArr, 2)
        lHigh2 = UBound(vArr, 2)
        lValIdx = lLow1 + 1&
        If m_Pos + ((lHigh2 - lLow2 + 1&) * 2&) + 2& >= m_Cap Then EnsureCapacity ((lHigh2 - lLow2 + 1&) * 2&) + 2&
        m_Buf(m_Pos) = 123   ' '{'
        m_Pos = m_Pos + 1&
        For i = lLow2 To lHigh2
            If i > lLow2 Then
                m_Buf(m_Pos) = 44
                m_Pos = m_Pos + 1&
            End If
            WriteEscapedString CStr(vArr(lLow1, i))
            m_Buf(m_Pos) = 58
            m_Pos = m_Pos + 1&
            WriteVariantValue vArr(lValIdx, i)
        Next i
        If m_Pos >= m_Cap Then EnsureCapacity 1&
        m_Buf(m_Pos) = 125   ' '}'
        m_Pos = m_Pos + 1&
    Else
        Err.Raise vbObjectError + 5023, "JsonWrite", "Arrays with more than 2 dimensions are not supported"
    End If
End Sub


' Extracts the accumulated JSON buffer as a VB6 String in a single CopyMemory call
Public Function ToString() As String
    If m_Pos <= 0& Then Exit Function
    ToString = Space$(m_Pos)
    CopyMemory ByVal StrPtr(ToString), ByVal VarPtr(m_Buf(0)), m_Pos * 2&
End Function


' Converts the accumulated UTF-16 buffer directly into a UTF-8 Byte() array
Public Function ToUtf8Bytes() As Byte()
    Dim lSize As Long
    Dim outBytes() As Byte
    If m_Pos <= 0& Then Exit Function
    lSize = WideCharToMultiByte(CP_UTF8, 0&, ByVal VarPtr(m_Buf(0)), m_Pos, 0&, 0&, 0&, 0&)
    If lSize > 0& Then
        ReDim outBytes(lSize - 1&)
        Call WideCharToMultiByte(CP_UTF8, 0&, ByVal VarPtr(m_Buf(0)), m_Pos, ByVal VarPtr(outBytes(0)), lSize, 0&, 0&)
        ToUtf8Bytes = outBytes
    End If
End Function

' Writes the accumulated JSON buffer directly to a UTF-8 file (supports Unicode paths)
Public Sub SaveToFile(ByRef sFilePath As String, Optional ByVal blnWriteBom As Boolean = False)
    Dim hFile As Long
    Dim lSize As Long
    Dim lWritten As Long
    Dim outBytes() As Byte
    Dim bBom(0 To 2) As Byte

    hFile = CreateFileW(StrPtr(sFilePath), GENERIC_WRITE, 0&, 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0&)
    If hFile = INVALID_HANDLE_VALUE Then
        Err.Raise vbObjectError + 5025, "JsonWrite", "Failed to open file for writing: " & sFilePath
    End If

    If blnWriteBom Then
        bBom(0) = &HEF
        bBom(1) = &HBB
        bBom(2) = &HBF
        If WriteFile(hFile, bBom(0), 3&, lWritten, 0&) = 0& Then
            CloseHandle hFile
            Err.Raise vbObjectError + 5026, "JsonWrite", "Failed writing UTF-8 BOM to file"
        End If
    End If

    If m_Pos > 0& Then
        lSize = WideCharToMultiByte(CP_UTF8, 0&, ByVal VarPtr(m_Buf(0)), m_Pos, 0&, 0&, 0&, 0&)
        If lSize > 0& Then
            ReDim outBytes(lSize - 1&)
            Call WideCharToMultiByte(CP_UTF8, 0&, ByVal VarPtr(m_Buf(0)), m_Pos, ByVal VarPtr(outBytes(0)), lSize, 0&, 0&)
            If WriteFile(hFile, outBytes(0), lSize, lWritten, 0&) = 0& Or lWritten <> lSize Then
                CloseHandle hFile
                Err.Raise vbObjectError + 5026, "JsonWrite", "Failed writing UTF-8 payload to file"
            End If
        End If
    End If

    CloseHandle hFile
End Sub


