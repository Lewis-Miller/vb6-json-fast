VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSonFast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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

Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

' --- MICROSOFT C-RUNTIME VECTORIZED MEMORY SCANNERS ---
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpWideCharStrDst As Long, ByVal cchWideCharDst As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

' --- CONSTANTS ---
Private Const CP_UTF8 As Long = 65001

' Token Type Enums
Public Enum TokenType
    TOK_NONE = 0
    TOK_OBJECT_OPEN = 1
    TOK_OBJECT_CLOSE = 2
    TOK_ARRAY_OPEN = 3
    TOK_ARRAY_CLOSE = 4
    TOK_KEY = 5
    TOK_STRING = 6
    TOK_NUMBER = 7
    TOK_TRUE = 8
    TOK_FALSE = 9
    TOK_NULL = 10
End Enum

' --- PARALLEL SCALAR VECTOR STORAGE ---
Private TokType() As Byte
Private TokOffset() As Long
Private TokLength() As Long
Private TokNext() As Long
Private TokChild() As Long
Private TokTail() As Long

' --- TRACKING REGISTERS ---
Private TokenCount As Long
Private TokensAllocated As Long
Private TokenBatchSize As Long

' --- SOURCE BUFFER REGISTERS ---
Private JsonBytes() As Byte
Private JsonByteLength As Long

'property vars
Private m_Strict As Boolean

Private Sub Class_Initialize()
    TokenBatchSize = 100000
    TokensAllocated = TokenBatchSize

    ReDim TokType(TokensAllocated - 1&)
    ReDim TokOffset(TokensAllocated - 1&)
    ReDim TokLength(TokensAllocated - 1&)
    ReDim TokNext(TokensAllocated - 1&)
    ReDim TokChild(TokensAllocated - 1&)
    ReDim TokTail(TokensAllocated - 1&)

End Sub

Private Sub EnsureTokenCapacity()
    If TokenCount >= TokensAllocated Then
        TokensAllocated = TokensAllocated * 2&
        ReDim Preserve TokType(TokensAllocated - 1&)
        ReDim Preserve TokOffset(TokensAllocated - 1&)
        ReDim Preserve TokLength(TokensAllocated - 1&)
        ReDim Preserve TokNext(TokensAllocated - 1&)
        ReDim Preserve TokChild(TokensAllocated - 1&)
        ReDim Preserve TokTail(TokensAllocated - 1&)
    End If
End Sub

Public Property Get Count() As Long
    Count = TokenCount
End Property

Public Property Let Strict(ByVal value As Boolean)
    m_Strict = value
End Property

Public Property Get Strict() As Boolean
    Strict = m_Strict
End Property



Public Sub Parse(ByRef vntJson As Variant)
    Dim lBytePos As Long
    Dim lLen As Long
    Dim lSize As Long
    Dim rootByte As Byte

    If VarType(vntJson) = vbString Then
        lLen = Len(vntJson)
        If lLen = 0 Then GoTo EmptyPayload

        lSize = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(vntJson), lLen, 0&, 0&, 0&, 0&)
        If lSize <= 0 Then GoTo EmptyPayload
        
        ReDim JsonBytes(lSize - 1) As Byte
        lSize = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(vntJson), lLen, ByVal VarPtr(JsonBytes(0)), lSize, 0&, 0&)
    ElseIf VarType(vntJson) = (vbByte Or vbArray) Then
        JsonBytes = vntJson
    Else
        Debug.Print "CRASH CONTEXT:" & vbCrLf & GetContextChunk(lBytePos)
        Err.Raise vbObjectError + 5001, "JSonFast", "Pos " & lBytePos & " - Invalid data source passed to Parse()"
    End If

    On Error GoTo ErrorHandle
    JsonByteLength = UBound(JsonBytes) + 1
    On Error GoTo 0

    TokenCount = 0

    If JsonByteLength >= 3 Then
        If JsonBytes(0) = 239 And JsonBytes(1) = 187 And JsonBytes(2) = 191 Then
            lBytePos = 3
        End If
    End If

    SkipWhitespace lBytePos
    If lBytePos >= JsonByteLength Then GoTo EmptyPayload

    rootByte = JsonBytes(lBytePos)

    If rootByte = 123 Then
        ParseObject lBytePos, -1&
    ElseIf rootByte = 91 Then
        ParseArray lBytePos, -1&
    Else
        Debug.Print "CRASH CONTEXT:" & vbCrLf & GetContextChunk(lBytePos)
        Err.Raise vbObjectError + 5002, "JSonFast", "Pos " & lBytePos & " - JSON payload must start with '{' or '['"
    End If

    If m_Strict Then
        SkipWhitespace lBytePos
        If lBytePos < JsonByteLength Then
            Debug.Print "CRASH CONTEXT:" & vbCrLf & GetContextChunk(lBytePos)
            Err.Raise vbObjectError + 5011, "JSonFast", "Pos " & lBytePos & " - Trailing content after root value"
        End If
    End If

    If TokenCount > 0 Then
        TokensAllocated = TokenCount
        ReDim Preserve TokType(TokensAllocated - 1&)
        ReDim Preserve TokOffset(TokensAllocated - 1&)
        ReDim Preserve TokLength(TokensAllocated - 1&)
        ReDim Preserve TokNext(TokensAllocated - 1&)
        ReDim Preserve TokChild(TokensAllocated - 1&)
        ReDim Preserve TokTail(TokensAllocated - 1&)
    End If

    Exit Sub

EmptyPayload:
    TokenCount = 0
    Erase JsonBytes
    JsonByteLength = 0
    Exit Sub

ErrorHandle:
    Debug.Print "CRASH CONTEXT:" & vbCrLf & GetContextChunk(lBytePos)
    Erase JsonBytes
    JsonByteLength = 0
    Err.Raise vbObjectError + 5003, "JSonFast", "Pos " & lBytePos & " - Parse() error occurred. " & Err.Description
End Sub

Private Function ParseObject(ByRef lBytePos As Long, ByVal parentIdx As Long) As Long
    Dim currentIdx As Long
    Dim lastElementIdx As Long

    EnsureTokenCapacity
    currentIdx = TokenCount
    TokenCount = TokenCount + 1

    TokType(currentIdx) = TOK_OBJECT_OPEN
    TokOffset(currentIdx) = lBytePos
    TokLength(currentIdx) = 1
    TokNext(currentIdx) = -1
    TokChild(currentIdx) = -1
    TokTail(currentIdx) = -1

    ' Link container to parent container vertically only
    If parentIdx <> -1 Then
        If TokChild(parentIdx) = -1 Then TokChild(parentIdx) = currentIdx
    End If

    lBytePos = lBytePos + 1    ' Skip '{'
    lastElementIdx = -1

    Do
        SkipWhitespace lBytePos
        If lBytePos >= JsonByteLength Then Exit Do

        If JsonBytes(lBytePos) = 125 Then    ' '}'
            lBytePos = lBytePos + 1
            EnsureTokenCapacity
            Dim closeIdx As Long
            closeIdx = TokenCount
            TokenCount = TokenCount + 1
            TokType(closeIdx) = TOK_OBJECT_CLOSE
            TokOffset(closeIdx) = lBytePos - 1
            TokLength(closeIdx) = 1
            TokNext(closeIdx) = -1
            TokChild(closeIdx) = currentIdx
            TokNext(currentIdx) = closeIdx
            ParseObject = currentIdx
            Exit Function
        End If

        ' Parse Key and register it as an immediate sibling of the object container
        Dim keyIdx As Long
        keyIdx = ConsumeString(lBytePos, currentIdx, TOK_KEY)

        If lastElementIdx = -1 Then
            TokChild(currentIdx) = keyIdx
            TokTail(currentIdx) = keyIdx
        Else
            TokNext(lastElementIdx) = keyIdx
            TokTail(currentIdx) = keyIdx
        End If
        lastElementIdx = keyIdx

        SkipWhitespace lBytePos
        If lBytePos >= JsonByteLength Then Err.Raise vbObjectError + 5004, "JSonFast", "Unexpected EOF waiting for ':'"
        If JsonBytes(lBytePos) <> 58 Then Err.Raise vbObjectError + 5004, "JSonFast", "Expected ':' delimiter"
        lBytePos = lBytePos + 1    ' Skip ':'

        ' Parse Value and explicitly map it as a child of the matching key
        Dim valIdx As Long
        valIdx = ParseValue(lBytePos, keyIdx)
        TokChild(keyIdx) = valIdx

        SkipWhitespace lBytePos
        If lBytePos >= JsonByteLength Then Exit Do

        Dim nextByte As Byte
        nextByte = JsonBytes(lBytePos)

        If nextByte = 44 Then
            lBytePos = lBytePos + 1
        ElseIf nextByte = 125 Then
            ' Handled on loop iteration boundary check
        Else
            Err.Raise vbObjectError + 5005, "JSonFast", "Expected ',' or '}' structural separator (Found ASCII " & nextByte & " at pos " & lBytePos & ")"
        End If
    Loop

    ParseObject = currentIdx
End Function

Private Function ParseArray(ByRef lBytePos As Long, ByVal parentIdx As Long) As Long
    Dim currentIdx As Long
    Dim lastElementIdx As Long

    EnsureTokenCapacity
    currentIdx = TokenCount
    TokenCount = TokenCount + 1

    TokType(currentIdx) = TOK_ARRAY_OPEN
    TokOffset(currentIdx) = lBytePos
    TokLength(currentIdx) = 1
    TokNext(currentIdx) = -1
    TokChild(currentIdx) = -1
    TokTail(currentIdx) = -1

    ' Link container to parent container vertically only
    If parentIdx <> -1 Then
        If TokChild(parentIdx) = -1 Then TokChild(parentIdx) = currentIdx
    End If

    lBytePos = lBytePos + 1    ' Skip '['
    lastElementIdx = -1

    Do
        SkipWhitespace lBytePos
        If lBytePos >= JsonByteLength Then Exit Do

        If JsonBytes(lBytePos) = 93 Then    ' ']'
            lBytePos = lBytePos + 1
            EnsureTokenCapacity
            Dim closeIdx As Long
            closeIdx = TokenCount
            TokenCount = TokenCount + 1
            TokType(closeIdx) = TOK_ARRAY_CLOSE
            TokOffset(closeIdx) = lBytePos - 1
            TokLength(closeIdx) = 1
            TokNext(closeIdx) = -1
            TokChild(closeIdx) = currentIdx
            TokNext(currentIdx) = closeIdx
            ParseArray = currentIdx
            Exit Function
        End If

        Dim valIdx As Long
        valIdx = ParseValue(lBytePos, currentIdx)

        ' Array values are tracked horizontally straight inside the array container
        If lastElementIdx = -1 Then
            TokChild(currentIdx) = valIdx
            TokTail(currentIdx) = valIdx
        Else
            TokNext(lastElementIdx) = valIdx
            TokTail(currentIdx) = valIdx
        End If
        lastElementIdx = valIdx

        SkipWhitespace lBytePos
        If lBytePos >= JsonByteLength Then Exit Do

        Dim nextByte As Byte
        nextByte = JsonBytes(lBytePos)

        If nextByte = 44 Then
            lBytePos = lBytePos + 1
        ElseIf nextByte = 93 Then
            ' Handled on loop iteration boundary check
        Else
            Err.Raise vbObjectError + 5006, "JSonFast", "Expected ',' or ']' sequence block separator (Found ASCII " & nextByte & " at pos " & lBytePos & ")"
        End If
    Loop

    ParseArray = currentIdx
End Function

Private Function ParseValue(ByRef lBytePos As Long, ByVal parentIdx As Long) As Long
    SkipWhitespace lBytePos
    If lBytePos >= JsonByteLength Then Exit Function

    Dim targetByte As Byte
    Dim thisValIdx As Long

    targetByte = JsonBytes(lBytePos)
    thisValIdx = TokenCount

    Select Case targetByte
        Case 34:
            thisValIdx = ConsumeString(lBytePos, parentIdx, TOK_STRING)
        Case 123:
            thisValIdx = ParseObject(lBytePos, parentIdx)    ' Correctly maps object token index
        Case 91:
            thisValIdx = ParseArray(lBytePos, parentIdx)  ' Correctly maps array token index
        Case 116:
            thisValIdx = ConsumeLiteral(lBytePos, parentIdx, TOK_TRUE, 4&)
        Case 102:
            thisValIdx = ConsumeLiteral(lBytePos, parentIdx, TOK_FALSE, 5&)
        Case 110:
            thisValIdx = ConsumeLiteral(lBytePos, parentIdx, TOK_NULL, 4&)
        Case 45, 48 To 57:
            thisValIdx = ConsumeNumber(lBytePos, parentIdx)
        Case Else
            Err.Raise vbObjectError + 5007, "JSonFast", "Invalid lexical character (ASCII " & targetByte & ") at pos " & lBytePos
    End Select

    ParseValue = thisValIdx
End Function





Private Function ConsumeString(ByRef lBytePos As Long, ByVal parentIdx As Long, ByVal tType As Byte) As Long
    Dim lStart As Long
    Dim lEnd As Long
    Dim thisTokenIdx As Long
    Dim lSlashCheck As Long
    Dim lSlashCount As Long

    lBytePos = lBytePos + 1
    lStart = lBytePos

    lEnd = lStart
    Do While lEnd < JsonByteLength
        If JsonBytes(lEnd) = 34 Then
            lSlashCount = 0
            lSlashCheck = lEnd - 1
            Do While lSlashCheck >= lStart
                If JsonBytes(lSlashCheck) = 92 Then
                    lSlashCount = lSlashCount + 1
                    lSlashCheck = lSlashCheck - 1
                Else
                    Exit Do
                End If
            Loop
            If (lSlashCount And 1&) = 0& Then Exit Do
        End If
        lEnd = lEnd + 1
    Loop
    lBytePos = lEnd + 1

    If lBytePos > JsonByteLength + 1 Then
        Err.Raise vbObjectError + 5008, "JSonFast", "Unterminated string sequence in structural payload"
    End If

    EnsureTokenCapacity
    thisTokenIdx = TokenCount
    TokenCount = TokenCount + 1

    TokType(thisTokenIdx) = tType
    TokOffset(thisTokenIdx) = lStart
    TokLength(thisTokenIdx) = lEnd - lStart
    TokNext(thisTokenIdx) = -1
    TokChild(thisTokenIdx) = -1
    TokTail(thisTokenIdx) = -1

    ConsumeString = thisTokenIdx
End Function

Private Function ConsumeNumber(ByRef lBytePos As Long, ByVal parentIdx As Long) As Long
    Dim lStart As Long
    Dim thisTokenIdx As Long

    lStart = lBytePos
    Do While lBytePos < JsonByteLength
        Select Case JsonBytes(lBytePos)
            Case 48 To 57, 46, 69, 101, 43, 45
                lBytePos = lBytePos + 1
            Case Else
                Exit Do
        End Select
    Loop

    EnsureTokenCapacity
    thisTokenIdx = TokenCount
    TokenCount = TokenCount + 1

    TokType(thisTokenIdx) = TOK_NUMBER
    TokOffset(thisTokenIdx) = lStart
    TokLength(thisTokenIdx) = lBytePos - lStart
    TokNext(thisTokenIdx) = -1
    TokChild(thisTokenIdx) = -1
    TokTail(thisTokenIdx) = -1

    ConsumeNumber = thisTokenIdx
End Function

Private Function ConsumeLiteral(ByRef lBytePos As Long, ByVal parentIdx As Long, ByVal tType As Byte, ByVal lLength As Long) As Long
    Dim thisTokenIdx As Long
    Dim bBadLiteral As Boolean

    EnsureTokenCapacity
    thisTokenIdx = TokenCount
    TokenCount = TokenCount + 1

    TokType(thisTokenIdx) = tType
    TokOffset(thisTokenIdx) = lBytePos
    TokLength(thisTokenIdx) = lLength
    TokNext(thisTokenIdx) = -1
    TokChild(thisTokenIdx) = -1
    TokTail(thisTokenIdx) = -1


    If m_Strict Then

        If ((lLength <> 4) And (lLength <> 5)) Or ((lBytePos + lLength) > JsonByteLength) Then
            bBadLiteral = True
        Else
            Select Case tType
                Case TOK_TRUE
                    If JsonBytes(lBytePos) <> 116 Or JsonBytes(lBytePos + 1) <> 114 Or _
                       JsonBytes(lBytePos + 2) <> 117 Or JsonBytes(lBytePos + 3) <> 101 Then
                        bBadLiteral = True
                    End If

                Case TOK_FALSE
                    If JsonBytes(lBytePos) <> 102 Or JsonBytes(lBytePos + 1) <> 97 Or _
                       JsonBytes(lBytePos + 2) <> 108 Or JsonBytes(lBytePos + 3) <> 115 Or _
                       JsonBytes(lBytePos + 4) <> 101 Then
                        bBadLiteral = True
                    End If

                Case TOK_NULL
                    If JsonBytes(lBytePos) <> 110 Or JsonBytes(lBytePos + 1) <> 117 Or _
                       JsonBytes(lBytePos + 2) <> 108 Or JsonBytes(lBytePos + 3) <> 108 Then
                        bBadLiteral = True
                    End If

                Case Else
                    'sum ting wong

            End Select
        End If
        If bBadLiteral Then Err.Raise vbObjectError + 5010, "JSonFast", "Invalid or malformed literal"
    End If


    lBytePos = lBytePos + lLength
    ConsumeLiteral = thisTokenIdx
End Function


Private Sub SkipWhitespace(ByRef lBytePos As Long)
    Do While lBytePos < JsonByteLength
        Select Case JsonBytes(lBytePos)
            Case 32, 9, 10, 13
                lBytePos = lBytePos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub


' --- CORE INTERFACE ---
Public Property Get Item(ByVal sPath As String) As String
Attribute Item.VB_UserMemId = 0
    Dim targetIdx As Long

    If TokenCount = 0 Then Exit Property

    ' Normalize path boundaries cleanly
    If Left$(sPath, 1) = "/" Then sPath = Mid$(sPath, 2)
    If Right$(sPath, 1) = "/" Then sPath = Left$(sPath, Len(sPath) - 1)

    ' Core fix: If requesting a top-level root element value, search right from index 0
    targetIdx = NavigateToPath(0&, sPath)

    If targetIdx <> -1 Then
        Item = ResolveValueString(targetIdx)
    End If
End Property

Public Property Get ItemAsDate(ByVal sPath As String, Optional ByVal AdjustToLocal As Boolean = True) As Date
    Dim targetIdx As Long

    If TokenCount = 0 Then Exit Property

    If Left$(sPath, 1) = "/" Then sPath = Mid$(sPath, 2)
    If Right$(sPath, 1) = "/" Then sPath = Left$(sPath, Len(sPath) - 1)

    targetIdx = NavigateToPath(0&, sPath)

    If targetIdx <> -1 Then
        If TokType(targetIdx) = TOK_STRING And TokLength(targetIdx) >= 10 Then
            ' Explicitly pass our main JsonBytes database register reference down to convert
            ItemAsDate = ParseIsoDateFromBytes(JsonBytes, TokOffset(targetIdx), TokLength(targetIdx))

            If AdjustToLocal Then
                ItemAsDate = UtcToLocalDate(ItemAsDate)
            End If
        End If
    End If
End Property

Public Function ItemAsLong(ByVal sPath As String) As Long
    Dim targetIdx As Long
    Dim sVal As String
    Dim dVal As Double

    If TokenCount = 0 Then Exit Function

    If Left$(sPath, 1) = "/" Then sPath = Mid$(sPath, 2)
    If Right$(sPath, 1) = "/" Then sPath = Left$(sPath, Len(sPath) - 1)

    targetIdx = NavigateToPath(0&, sPath)
    If targetIdx = -1 Then Exit Function

    Select Case TokType(targetIdx)
        Case TOK_NUMBER
            sVal = StringFromBytes(TokOffset(targetIdx), TokLength(targetIdx))
            dVal = Val(sVal)
            If dVal >= -2147483648# And dVal <= 2147483647# Then
                ItemAsLong = CLng(dVal)
            Else
                Err.Raise vbObjectError + 5012, "JSonFast", "Value out of Long range at path: " & sPath
            End If
        Case TOK_TRUE
            ItemAsLong = 1
        Case TOK_FALSE, TOK_NULL
            ItemAsLong = 0
    End Select
End Function

Public Function ItemAsDouble(ByVal sPath As String) As Double
    Dim targetIdx As Long
    Dim sVal As String

    If TokenCount = 0 Then Exit Function

    If Left$(sPath, 1) = "/" Then sPath = Mid$(sPath, 2)
    If Right$(sPath, 1) = "/" Then sPath = Left$(sPath, Len(sPath) - 1)

    targetIdx = NavigateToPath(0&, sPath)
    If targetIdx = -1 Then Exit Function

    Select Case TokType(targetIdx)
        Case TOK_NUMBER
            sVal = StringFromBytes(TokOffset(targetIdx), TokLength(targetIdx))
            ItemAsDouble = CDbl(Val(sVal))
        Case TOK_TRUE
            ItemAsDouble = 1
        Case TOK_FALSE, TOK_NULL
            ItemAsDouble = 0
    End Select
End Function

Public Function ItemAsBoolean(ByVal sPath As String) As Boolean
    Dim targetIdx As Long
    Dim sVal As String

    If TokenCount = 0 Then Exit Function

    If Left$(sPath, 1) = "/" Then sPath = Mid$(sPath, 2)
    If Right$(sPath, 1) = "/" Then sPath = Left$(sPath, Len(sPath) - 1)

    targetIdx = NavigateToPath(0&, sPath)
    If targetIdx = -1 Then Exit Function

    Select Case TokType(targetIdx)
        Case TOK_TRUE
            ItemAsBoolean = True
        Case TOK_FALSE, TOK_NULL
            ItemAsBoolean = False
        Case TOK_NUMBER
            sVal = StringFromBytes(TokOffset(targetIdx), TokLength(targetIdx))
             ItemAsBoolean = (CDbl(Val(sVal)) <> 0#)
        Case TOK_STRING
            sVal = ResolveValueString(targetIdx)
            ItemAsBoolean = (StrComp(sVal, "true", vbTextCompare) = 0)
    End Select
End Function

Public Function GetRaw(ByVal sPath As String) As String
    Dim targetIdx As Long
    Dim closeIdx As Long
    Dim lStart As Long
    Dim lEnd As Long

    If TokenCount = 0 Then Exit Function

    If Left$(sPath, 1) = "/" Then sPath = Mid$(sPath, 2)
    If Right$(sPath, 1) = "/" Then sPath = Left$(sPath, Len(sPath) - 1)

    targetIdx = NavigateToPath(0&, sPath)
    If targetIdx = -1 Then Exit Function

    ' for containers use the matching close token so the slice includes the full structure
    If TokType(targetIdx) = TOK_OBJECT_OPEN Or TokType(targetIdx) = TOK_ARRAY_OPEN Then
        closeIdx = TokNext(targetIdx)
        If closeIdx = -1 Then Exit Function
        lStart = TokOffset(targetIdx)
        lEnd = TokOffset(closeIdx) + TokLength(closeIdx)
    Else
        lStart = TokOffset(targetIdx)
        lEnd = lStart + TokLength(targetIdx)
    End If

    If lEnd > lStart And lEnd <= JsonByteLength Then
        GetRaw = StringFromBytes(lStart, lEnd - lStart)
    End If
End Function


Public Function Exists(ByVal sPath As String) As Boolean
    Dim targetIdx As Long

    If TokenCount = 0 Then Exit Function

    If Left$(sPath, 1) = "/" Then sPath = Mid$(sPath, 2)
    If Right$(sPath, 1) = "/" Then sPath = Left$(sPath, Len(sPath) - 1)

    targetIdx = NavigateToPath(0&, sPath)
    Exists = (targetIdx <> -1)
End Function

Private Function NavigateToPath(ByVal startIdx As Long, ByVal sPath As String) As Long
    Dim segments() As String
    Dim currentIdx As Long
    Dim i As Long
    Dim sTargetSegment As String
    Dim bFound As Boolean

    NavigateToPath = -1
    If TokenCount = 0 Then Exit Function
    If LenB(sPath) = 0 Then
        NavigateToPath = startIdx
        Exit Function
    End If

    segments = Split(sPath, "/")
    currentIdx = startIdx

    For i = LBound(segments) To UBound(segments)
        ' Safety verification check: prevent subscript out of range errors
        If currentIdx < 0 Or currentIdx >= TokenCount Then
            NavigateToPath = -1
            Exit Function
        End If

        sTargetSegment = segments(i)
        bFound = False

        ' 1. If currently pointing directly to a KEY token, jump inside its value container
        If TokType(currentIdx) = TOK_KEY Then
            currentIdx = TokChild(currentIdx)
            If currentIdx = -1 Then Exit Function
        End If

        ' 2. If current node is an Object Container, scan horizontally across child keys
        If TokType(currentIdx) = TOK_OBJECT_OPEN Then
            Dim childIdx As Long
            childIdx = TokChild(currentIdx)

            Do While childIdx <> -1
                If TokType(childIdx) = TOK_KEY Then
                    ' Lazy evaluation: extract key text value on-the-fly to compare
                    If StrComp(ResolveValueString(childIdx), sTargetSegment, vbTextCompare) = 0 Then
                        currentIdx = childIdx
                        bFound = True
                        Exit Do
                    End If
                End If
                childIdx = TokNext(childIdx)
            Loop

            ' 3. If current node is an Array Container, evaluate the numerical segment index position
        ElseIf TokType(currentIdx) = TOK_ARRAY_OPEN Then
            If IsNumeric(sTargetSegment) Then
                Dim targetArrIdx As Long
                Dim arrayWalker As Long
                Dim currentArrPos As Long

                targetArrIdx = CLng(sTargetSegment)
                arrayWalker = TokChild(currentIdx)
                currentArrPos = 0

                Do While arrayWalker <> -1
                    If currentArrPos = targetArrIdx Then
                        currentIdx = arrayWalker
                        bFound = True
                        Exit Do
                    End If
                    currentArrPos = currentArrPos + 1
                    arrayWalker = TokNext(arrayWalker)
                Loop
            End If
        End If

        If Not bFound Then
            NavigateToPath = -1
            Exit Function
        End If
    Next i

    ' Resolve structural value mappings cleanly
    If currentIdx >= 0 And currentIdx < TokenCount Then
        If TokType(currentIdx) = TOK_KEY Then
            NavigateToPath = TokChild(currentIdx)
        Else
            NavigateToPath = currentIdx
        End If
    Else
        NavigateToPath = -1
    End If
End Function

Public Function GetArrayCount(ByVal sPath As String) As Long
    Dim targetIdx As Long

    If TokenCount = 0 Then Exit Function

    If Left$(sPath, 1) = "/" Then sPath = Mid$(sPath, 2)
    If Right$(sPath, 1) = "/" Then sPath = Left$(sPath, Len(sPath) - 1)

    targetIdx = NavigateToPath(0&, sPath)

    If targetIdx <> -1 Then
        If TokType(targetIdx) = TOK_ARRAY_OPEN Then
            Dim childIdx As Long
            childIdx = TokChild(targetIdx)
            Do While childIdx <> -1
                GetArrayCount = GetArrayCount + 1
                childIdx = TokNext(childIdx)
            Loop
        End If
    End If
End Function


'Low-level token access

'===========================================================
''example useage (note: this will require TOK_* consts public access)
'Dim root As Long
'root = 0   ' after Parse, root is always index 0 for a top-level object/array

'If parser.GetTokenType(root) = TOK_ARRAY_OPEN Then
    'Dim child As Long
    'child = parser.GetFirstChild(root)
    'Do While child <> -1
        ' 'child is the value token
        'Debug.Print parser.GetValueString(child)
        'child = parser.GetNextSibling(child)
   ' Loop
'End If
'=========================================================

Public Function GetTokenType(ByVal Index As Long) As Byte
    If Index < 0 Or Index >= TokenCount Then Exit Function
    GetTokenType = TokType(Index)
End Function

Public Function GetFirstChild(ByVal Index As Long) As Long
    GetFirstChild = -1
    If Index < 0 Or Index >= TokenCount Then Exit Function
    GetFirstChild = TokChild(Index)
End Function

Public Function GetNextSibling(ByVal Index As Long) As Long
    Dim nextIdx As Long
    GetNextSibling = -1
    If Index < 0 Or Index >= TokenCount Then Exit Function
    nextIdx = TokNext(Index)
    If nextIdx <> -1 Then
        If TokType(nextIdx) <> TOK_OBJECT_CLOSE And TokType(nextIdx) <> TOK_ARRAY_CLOSE Then
            GetNextSibling = nextIdx
        End If
    End If
End Function



Public Function GetValueString(ByVal Index As Long) As String
    If Index < 0 Or Index >= TokenCount Then Exit Function
    GetValueString = ResolveValueString(Index)
End Function

' Convenience Walkers
'GetArrayItem walks the sibling chain n steps from TokChild. GetObjectValue walks the keys and does a single ResolveValueString + compare per
'key (same work NavigateToPath already does, but without the path-splitting overhead when the caller already holds the object token).
Public Function GetArrayItem(ByVal arrayIndex As Long, ByVal Index As Long) As Long
    Dim child As Long
    Dim n As Long

    GetArrayItem = -1
    If arrayIndex < 0 Or arrayIndex >= TokenCount Then Exit Function
    If TokType(arrayIndex) <> TOK_ARRAY_OPEN Then Exit Function
    If Index < 0 Then Exit Function

    child = TokChild(arrayIndex)
    n = 0
    Do While child <> -1
        If n = Index Then
            GetArrayItem = child
            Exit Function
        End If
        n = n + 1
        child = TokNext(child)
    Loop
End Function

Public Function GetObjectValue(ByVal objectIndex As Long, ByVal sKey As String) As Long
    Dim child As Long

    GetObjectValue = -1
    If objectIndex < 0 Or objectIndex >= TokenCount Then Exit Function
    If TokType(objectIndex) <> TOK_OBJECT_OPEN Then Exit Function
    If LenB(sKey) = 0 Then Exit Function

    child = TokChild(objectIndex)
    Do While child <> -1
        If TokType(child) = TOK_KEY Then
            If StrComp(ResolveValueString(child), sKey, vbTextCompare) = 0 Then
                GetObjectValue = TokChild(child)
                Exit Function
            End If
        End If
        child = TokNext(child)
    Loop
End Function

'returns the first child index of the array at that path (or -1). Caller then uses GetNextSibling.
Public Function EnumerateArray(ByVal path As String) As Long
    Dim idx As Long

    EnumerateArray = -1
    If TokenCount = 0 Then Exit Function

    If Left$(path, 1) = "/" Then path = Mid$(path, 2)
    If Right$(path, 1) = "/" Then path = Left$(path, Len(path) - 1)

    idx = NavigateToPath(0&, path)
    If idx = -1 Then Exit Function
    If TokType(idx) <> TOK_ARRAY_OPEN Then Exit Function

    EnumerateArray = TokChild(idx)
End Function

Public Function EnumerateObject(ByVal path As String) As Long
    Dim idx As Long

    EnumerateObject = -1
    If TokenCount = 0 Then Exit Function

    If Left$(path, 1) = "/" Then path = Mid$(path, 2)
    If Right$(path, 1) = "/" Then path = Left$(path, Len(path) - 1)

    idx = NavigateToPath(0&, path)
    If idx = -1 Then Exit Function
    If TokType(idx) <> TOK_OBJECT_OPEN Then Exit Function

    EnumerateObject = TokChild(idx)
End Function


Private Function ResolveValueString(ByVal tokenIdx As Long) As String
    Dim lOffset As Long
    Dim lLength As Long
    Dim tType As Byte

    tType = TokType(tokenIdx)
    lOffset = TokOffset(tokenIdx)
    lLength = TokLength(tokenIdx)

    Select Case tType
        Case TOK_KEY, TOK_STRING, TOK_NUMBER
            If lLength > 0 Then
                ' Extract the raw string from bytes first
                Dim sRawValue As String
                sRawValue = StringFromBytes(lOffset, lLength)

                ' AUTOMATIC UNESCAPING PASS:
                ' Check if this specific string literal contains a backslash escape marker.
                ' If it does, run it through UnEscape
                If tType = TOK_STRING Then
                    If InStr(sRawValue, "\") > 0 Then
                        ResolveValueString = UnEscape(sRawValue)
                        Exit Function
                    End If
                End If

                ' value path: Return the raw decoded text
                ResolveValueString = sRawValue
            Else
                ResolveValueString = ""
            End If
        Case TOK_TRUE
            ResolveValueString = "true"
        Case TOK_FALSE
            ResolveValueString = "false"
        Case TOK_NULL
            ResolveValueString = "null"
    End Select
End Function


Private Function StringFromBytes(ByVal lOffset As Long, ByVal lLength As Long) As String
    Dim lSize As Long
    lSize = MultiByteToWideChar(CP_UTF8, 0&, ByVal VarPtr(JsonBytes(lOffset)), lLength, 0&, 0&)
    If lSize > 0 Then
        StringFromBytes = String$(lSize, 0)
        Call MultiByteToWideChar(CP_UTF8, 0&, ByVal VarPtr(JsonBytes(lOffset)), lLength, ByVal StrPtr(StringFromBytes), lSize)
    End If
End Function


Private Sub Class_Terminate()
    Erase TokType
    Erase TokOffset
    Erase TokLength
    Erase TokNext
    Erase TokChild
    Erase TokTail
    Erase JsonBytes
End Sub


Public Function GetContextChunk(ByVal CurrentPos As Long) As String
    Dim lStart As Long
    Dim lEnd As Long
    Dim lLength As Long
    Dim lSize As Long

    If CurrentPos > 150 Then
        lStart = CurrentPos - 150
    Else
        lStart = 0
    End If

    If CurrentPos + 150 < JsonByteLength Then
        lEnd = CurrentPos + 150
    Else
        lEnd = JsonByteLength - 1
    End If

    lLength = lEnd - lStart + 1
    If lLength <= 0 Then Exit Function

    ' convert bytes
    lSize = MultiByteToWideChar(CP_UTF8, 0&, ByVal VarPtr(JsonBytes(lStart)), lLength, 0&, 0&)
    If lSize > 0 Then
        GetContextChunk = String$(lSize, 0)
        lSize = MultiByteToWideChar(CP_UTF8, 0&, ByVal VarPtr(JsonBytes(lStart)), lLength, ByVal StrPtr(GetContextChunk), lSize)

        'leave a visual marker indicating exactly where the pointer was sitting
        Dim lMarkerPos As Long
        lMarkerPos = (CurrentPos - lStart) + 1
        If lMarkerPos <= Len(GetContextChunk) Then
            GetContextChunk = Left$(GetContextChunk, lMarkerPos - 1) & "??" & Mid$(GetContextChunk, lMarkerPos)
        End If
    End If
End Function



' =======================================================================================
' JSON FORMATTING AND TRANSFORMATION UTILITIES
' =======================================================================================

Public Function Escape(ByVal strInput As String) As String
    Dim lInputLen As Long
    Dim inArr() As Integer
    Dim outArr() As Integer
    Dim lSrcIdx As Long, lDstIdx As Long, lMaxOutLen As Long
    Dim charCode As Long, nextCharCode As Long
    Dim sHex As String, i As Long
    
    If LenB(strInput) = 0 Then Exit Function

    lInputLen = Len(strInput)

    ' map the string to a 2-byte integer array
    ReDim inArr(lInputLen - 1)
    CopyMemory ByVal VarPtr(inArr(0)), ByVal StrPtr(strInput), lInputLen * 2&

    ' pre-allocate a cushion buffer (Max 6x expansion for \uXXXX states)
    lMaxOutLen = lInputLen * 6&
    ReDim outArr(lMaxOutLen - 1)

    lDstIdx = 0
    For lSrcIdx = 0 To lInputLen - 1
        'signed-to-unsigned conversion for high-order plane characters
        charCode = inArr(lSrcIdx)
        If charCode < 0 Then charCode = charCode + 65536

        ' handle UTF-16 Surrogate Pairs (e.g emojis)
        If charCode >= &HD800 And charCode <= &HDBFF And lSrcIdx < lInputLen - 1 Then
            nextCharCode = inArr(lSrcIdx + 1)
            If nextCharCode < 0 Then nextCharCode = nextCharCode + 65536
            If nextCharCode >= &HDC00 And nextCharCode <= &HDFFF Then
                outArr(lDstIdx) = inArr(lSrcIdx)
                lDstIdx = lDstIdx + 1
                outArr(lDstIdx) = inArr(lSrcIdx + 1)
                lDstIdx = lDstIdx + 1
                lSrcIdx = lSrcIdx + 1
                GoTo NextChar
            End If
        End If

        ' process Standard Escaping Rules
        Select Case charCode
            Case 34    ' \"
                outArr(lDstIdx) = 92
                lDstIdx = lDstIdx + 1
                outArr(lDstIdx) = 34
                lDstIdx = lDstIdx + 1
                
            Case 92    ' \\
                outArr(lDstIdx) = 92
                lDstIdx = lDstIdx + 1
                outArr(lDstIdx) = 92
                lDstIdx = lDstIdx + 1
                
            Case 47    ' \/
                outArr(lDstIdx) = 92
                lDstIdx = lDstIdx + 1
                outArr(lDstIdx) = 47
                lDstIdx = lDstIdx + 1
                
            Case 8    ' \b
                outArr(lDstIdx) = 92
                lDstIdx = lDstIdx + 1
                outArr(lDstIdx) = 98
                lDstIdx = lDstIdx + 1
                
            Case 12    ' \f
                outArr(lDstIdx) = 92
                lDstIdx = lDstIdx + 1
                outArr(lDstIdx) = 102
                lDstIdx = lDstIdx + 1
                
            Case 10    ' \n
                outArr(lDstIdx) = 92
                lDstIdx = lDstIdx + 1
                outArr(lDstIdx) = 110
                lDstIdx = lDstIdx + 1
                
            Case 13    ' \r
                outArr(lDstIdx) = 92
                lDstIdx = lDstIdx + 1
                outArr(lDstIdx) = 114
                lDstIdx = lDstIdx + 1
                
            Case 9    ' \t
                outArr(lDstIdx) = 92
                lDstIdx = lDstIdx + 1
                outArr(lDstIdx) = 116
                lDstIdx = lDstIdx + 1
                
            Case 0 To 31, 127 To 159, &HD800 To &HDFFF    ' control characters -> \uXXXX
                outArr(lDstIdx) = 92
                lDstIdx = lDstIdx + 1
                outArr(lDstIdx) = 117
                lDstIdx = lDstIdx + 1
                sHex = Right$("0000" & Hex$(charCode), 4)
                
                For i = 1 To 4
                    outArr(lDstIdx) = AscW(Mid$(sHex, i, 1))
                    lDstIdx = lDstIdx + 1
                Next i
                
            Case Else
                outArr(lDstIdx) = inArr(lSrcIdx)
                lDstIdx = lDstIdx + 1
                
        End Select
NextChar:
    Next lSrcIdx

    ' truncate and compile array back into a BSTR
    If lDstIdx > 0 Then
        Escape = Space$(lDstIdx)
        CopyMemory ByVal StrPtr(Escape), ByVal VarPtr(outArr(0)), lDstIdx * 2&
    End If
End Function

Public Function Minify(ByVal strInput As String) As String
    Dim lInputLen As Long
    Dim inArr() As Integer
    Dim outArr() As Integer
    Dim lSrcIdx As Long, lDstIdx As Long
    Dim bInString As Boolean, bEscape As Boolean
    Dim charCode As Integer

    If LenB(strInput) = 0 Then Exit Function

    lInputLen = Len(strInput)

    ReDim inArr(lInputLen - 1)
    CopyMemory ByVal VarPtr(inArr(0)), ByVal StrPtr(strInput), lInputLen * 2&
    ReDim outArr(lInputLen - 1)

    lDstIdx = 0
    bInString = False
    bEscape = False

    For lSrcIdx = 0 To lInputLen - 1
        charCode = inArr(lSrcIdx)

        If bInString Then
            ' inside a string: copy everything exactly as is
            outArr(lDstIdx) = charCode
            lDstIdx = lDstIdx + 1

            If bEscape Then
                bEscape = False
            ElseIf charCode = 92 Then    ' backslash (\) escape marker
                bEscape = True
            ElseIf charCode = 34 Then    ' closing quote (") found
                bInString = False
            End If
        Else
            ' outside a string: filter out all whitespace
            Select Case charCode
                Case 34    ' opening quote (")
                    bInString = True
                    outArr(lDstIdx) = charCode
                    lDstIdx = lDstIdx + 1
                Case 32, 13, 10, 9    ' strip spaces, CR, LF, Tabs, and control characters
                    ' skip
                Case Else
                    'keep structural characters, numbers, and keys
                    outArr(lDstIdx) = charCode
                    lDstIdx = lDstIdx + 1
            End Select
        End If
    Next lSrcIdx

    If lDstIdx > 0 Then
        Minify = Space$(lDstIdx)
        CopyMemory ByVal StrPtr(Minify), ByVal VarPtr(outArr(0)), lDstIdx * 2&
    End If
End Function


Public Function UnEscape(ByVal strInput As String) As String
    Dim lInputLen As Long
    Dim inArr() As Integer
    Dim outArr() As Integer
    Dim lSrcIdx As Long, lDstIdx As Long
    Dim charCode As Long, hexVal As Long, i As Long
    Dim hByte As Integer

    If LenB(strInput) = 0 Then Exit Function

    lInputLen = Len(strInput)

    ReDim inArr(lInputLen - 1)
    CopyMemory ByVal VarPtr(inArr(0)), ByVal StrPtr(strInput), lInputLen * 2&
    ReDim outArr(lInputLen - 1)

    Do While lSrcIdx < lInputLen
        charCode = inArr(lSrcIdx)

        If charCode = 92 And lSrcIdx < lInputLen - 1 Then    ' backslash escape (\)
            lSrcIdx = lSrcIdx + 1
            Select Case inArr(lSrcIdx)
                Case 34  ' \"
                    outArr(lDstIdx) = 34
                    lDstIdx = lDstIdx + 1

                Case 92  ' \\
                    outArr(lDstIdx) = 92
                    lDstIdx = lDstIdx + 1

                Case 47  ' \/
                    outArr(lDstIdx) = 47
                    lDstIdx = lDstIdx + 1

                Case 98   ' \b
                    outArr(lDstIdx) = 8
                    lDstIdx = lDstIdx + 1

                Case 102  ' \f
                    outArr(lDstIdx) = 12
                    lDstIdx = lDstIdx + 1

                Case 110    ' \n
                    outArr(lDstIdx) = 10
                    lDstIdx = lDstIdx + 1

                Case 114  ' \r
                    outArr(lDstIdx) = 13
                    lDstIdx = lDstIdx + 1

                Case 116   ' \t
                    outArr(lDstIdx) = 9
                    lDstIdx = lDstIdx + 1

                Case 117    ' \uXXXX hex evaluation
                    ' make sure 4 hex digits exist
                    If lSrcIdx + 4 < lInputLen Then
                        hexVal = 0
                        For i = 1 To 4
                            lSrcIdx = lSrcIdx + 1
                            hByte = inArr(lSrcIdx)
                            Select Case hByte
                                Case 48 To 57    ' 0-9
                                    hexVal = hexVal * 16& + (hByte - 48)
                                Case 65 To 70   ' A-F
                                    hexVal = hexVal * 16& + (hByte - 55)
                                Case 97 To 102     ' a-f
                                    hexVal = hexVal * 16& + (hByte - 87)
                            End Select
                        Next i

                        If hexVal > 32767 Then hexVal = hexVal - 65536
                        outArr(lDstIdx) = CInt(hexVal)
                        lDstIdx = lDstIdx + 1
                    Else
                        ' out of bounds fallback, write literal markers
                        outArr(lDstIdx) = 92
                        lDstIdx = lDstIdx + 1
                        outArr(lDstIdx) = 117
                        lDstIdx = lDstIdx + 1
                    End If
                Case Else
                    outArr(lDstIdx) = 92
                    lDstIdx = lDstIdx + 1
                    outArr(lDstIdx) = inArr(lSrcIdx)
                    lDstIdx = lDstIdx + 1
            End Select
        Else
            outArr(lDstIdx) = inArr(lSrcIdx)
            lDstIdx = lDstIdx + 1
        End If
        lSrcIdx = lSrcIdx + 1
    Loop

    If lDstIdx > 0 Then
        UnEscape = Space$(lDstIdx)
        CopyMemory ByVal StrPtr(UnEscape), ByVal VarPtr(outArr(0)), lDstIdx * 2&
    End If
End Function

Public Function PrettyPrint(ByVal sInput As String) As String

    Dim inArr() As Integer, outArr() As Integer
    Dim lSrcIdx As Long, lDstIdx As Long, lMaxOutLen As Long
    Dim lIndentLevel As Long, lInputLen As Long
    Dim bInString As Boolean, bEscape As Boolean, bEmptyPair As Boolean
    Dim charCode As Integer, NextChar As Integer
    Dim k As Long
    Dim targetClose As Integer
    Dim lLookIdx As Long

    If LenB(sInput) = 0 Then Exit Function

    lInputLen = Len(sInput)

    ReDim inArr(lInputLen - 1)
    CopyMemory ByVal VarPtr(inArr(0)), ByVal StrPtr(sInput), lInputLen * 2&

    lMaxOutLen = lInputLen * 4&
    ReDim outArr(lMaxOutLen - 1)

    For lSrcIdx = 0 To lInputLen - 1
        charCode = inArr(lSrcIdx)

        If lDstIdx >= lMaxOutLen - (lIndentLevel * 4& + 10&) Then
            lMaxOutLen = lMaxOutLen + lInputLen
            ReDim Preserve outArr(lMaxOutLen - 1)
        End If

        If bInString Then
            outArr(lDstIdx) = charCode
            lDstIdx = lDstIdx + 1
            If bEscape Then
                bEscape = False
            ElseIf charCode = 92 Then
                bEscape = True
            ElseIf charCode = 34 Then
                bInString = False
            End If
        Else
            Select Case charCode
                Case 34
                    bInString = True
                    outArr(lDstIdx) = charCode
                    lDstIdx = lDstIdx + 1

                Case 123, 91
                    outArr(lDstIdx) = charCode
                    lDstIdx = lDstIdx + 1

                    bEmptyPair = False
                    If lSrcIdx < lInputLen - 1 Then
                        If charCode = 123 Then
                            targetClose = 125
                        Else
                            targetClose = 93
                        End If

                        lLookIdx = lSrcIdx + 1

                        Do While lLookIdx < lInputLen
                            NextChar = inArr(lLookIdx)
                            If NextChar = 32 Or NextChar = 13 Or NextChar = 10 Or NextChar = 9 Then
                                lLookIdx = lLookIdx + 1
                            Else
                                If NextChar = targetClose Then
                                    bEmptyPair = True
                                    lSrcIdx = lLookIdx - 1
                                End If
                                Exit Do
                            End If
                        Loop
                    End If

                    If Not bEmptyPair Then
                        lIndentLevel = lIndentLevel + 1
                        outArr(lDstIdx) = 13
                        lDstIdx = lDstIdx + 1
                        outArr(lDstIdx) = 10
                        lDstIdx = lDstIdx + 1

                        For k = 1 To lIndentLevel * 2&
                            outArr(lDstIdx) = 32: lDstIdx = lDstIdx + 1
                        Next k
                    End If

                Case 125, 93
                    If lIndentLevel > 0 Then lIndentLevel = lIndentLevel - 1
                    outArr(lDstIdx) = 13
                    lDstIdx = lDstIdx + 1
                    outArr(lDstIdx) = 10
                    lDstIdx = lDstIdx + 1

                    For k = 1 To lIndentLevel * 2&
                        outArr(lDstIdx) = 32: lDstIdx = lDstIdx + 1
                    Next k

                    outArr(lDstIdx) = charCode
                    lDstIdx = lDstIdx + 1

                Case 44
                    outArr(lDstIdx) = charCode
                    lDstIdx = lDstIdx + 1
                    outArr(lDstIdx) = 13
                    lDstIdx = lDstIdx + 1
                    outArr(lDstIdx) = 10
                    lDstIdx = lDstIdx + 1

                    For k = 1 To lIndentLevel * 2&
                        outArr(lDstIdx) = 32: lDstIdx = lDstIdx + 1
                    Next k

                Case 58
                    outArr(lDstIdx) = charCode
                    lDstIdx = lDstIdx + 1
                    outArr(lDstIdx) = 32
                    lDstIdx = lDstIdx + 1

                Case 32, 13, 10, 9
                    ' skip  whitespace

                Case Else
                    outArr(lDstIdx) = charCode
                    lDstIdx = lDstIdx + 1
            End Select
        End If
    Next lSrcIdx

    If lDstIdx > 0 Then
        PrettyPrint = Space$(lDstIdx)
        CopyMemory ByVal StrPtr(PrettyPrint), ByVal VarPtr(outArr(0)), lDstIdx * 2&
    End If
End Function



' parses an ISO 8601 date string from JsonBytes
Private Function ParseIsoDateFromBytes(RefBytes() As Byte, ByVal lOffset As Long, ByVal lLength As Long) As Date
    Dim lngYear As Long, lngMonth As Long, lngDay As Long
    Dim lngHour As Long, lngMinute As Long, lngSecond As Long
    Dim lngMilli As Long, lngOffsetMinutes As Long
    Dim lIdx As Long
    Dim b As Byte
    Dim lMilliStart As Long
    Dim sMilliBuf As String
    Dim oHr As Long, oMin As Long

    On Error GoTo ErrorHandling

    ' ASCII digit conversion: (Character - 48)
    ' Format: YYYY-MM-DD
    lngYear = (RefBytes(lOffset) - 48&) * 1000& + (RefBytes(lOffset + 1) - 48&) * 100& + (RefBytes(lOffset + 2) - 48&) * 10& + (RefBytes(lOffset + 3) - 48&)
    lngMonth = (RefBytes(lOffset + 5) - 48&) * 10& + (RefBytes(lOffset + 6) - 48&)
    lngDay = (RefBytes(lOffset + 8) - 48&) * 10& + (RefBytes(lOffset + 9) - 48&)

    ' time sections
    If lLength >= 16 Then
        If RefBytes(lOffset + 10) = 84 Then    ' 'T'
            lngHour = (RefBytes(lOffset + 11) - 48&) * 10& + (RefBytes(lOffset + 12) - 48&)
            lngMinute = (RefBytes(lOffset + 14) - 48&) * 10& + (RefBytes(lOffset + 15) - 48&)

            If lLength >= 19 Then
                If RefBytes(lOffset + 16) = 58 Then    ' ':'
                    lngSecond = (RefBytes(lOffset + 17) - 48&) * 10& + (RefBytes(lOffset + 18) - 48&)
                End If
            End If
        End If
    End If

    ' scan trailing character offsets for milliseconds or custom timezone offsets
    lIdx = 19
    Do While lIdx < lLength
        b = RefBytes(lOffset + lIdx)

        If b = 46 Then    ' '.' found decimal point for milliseconds
            lIdx = lIdx + 1
            sMilliBuf = ""

            ' extract up to 3 decimal digits (ex: .27, .3, or .355 layout)
            Do While lIdx < lLength
                b = RefBytes(lOffset + lIdx)
                If b >= 48 And b <= 57 Then
                    sMilliBuf = sMilliBuf & Chr$(b)
                    lIdx = lIdx + 1
                Else
                    Exit Do
                End If
            Loop

            If Len(sMilliBuf) > 0 Then
                ' normalize variable width up to standard millisecond scale
                lngMilli = CLng(Left$(sMilliBuf & "000", 3))
            End If
            lIdx = lIdx - 1    ' Re-align pointer increment

        ElseIf b = 43 Or b = 45 Then    ' '+' or '-' Zone Offset
            If lIdx + 5 < lLength Then
                oHr = (RefBytes(lOffset + lIdx + 1) - 48&) * 10& + (RefBytes(lOffset + lIdx + 2) - 48&)
                oMin = (RefBytes(lOffset + lIdx + 4) - 48&) * 10& + (RefBytes(lOffset + lIdx + 5) - 48&)
                lngOffsetMinutes = oMin + (oHr * 60&)
                If b = 43 Then lngOffsetMinutes = -lngOffsetMinutes    ' invert to establish UTC baseline
                Exit Do
            End If
        End If
        lIdx = lIdx + 1
    Loop

    ' convert into a vb Date
    ParseIsoDateFromBytes = DateSerial(lngYear, lngMonth, lngDay) + TimeSerial(lngHour, lngMinute, lngSecond)
    If lngMilli > 0 Then ParseIsoDateFromBytes = ParseIsoDateFromBytes + (lngMilli / 86400000#)

    If lngOffsetMinutes <> 0 Then
        ParseIsoDateFromBytes = DateAdd("n", lngOffsetMinutes, ParseIsoDateFromBytes)
    End If

    Exit Function
ErrorHandling:
    ParseIsoDateFromBytes = 0
End Function


'' convert an iso 8601 string to a local date
Public Function IsoToLocalDate(ByVal strIsoDate As String) As Date
    Dim lLen As Long
    Dim tempBytes() As Byte
    Dim dtUtc As Date
    
    lLen = Len(strIsoDate)
    If lLen = 0 Then Exit Function

    tempBytes = StrConv(strIsoDate, vbFromUnicode)

    dtUtc = ParseIsoDateFromBytes(tempBytes, 0&, UBound(tempBytes) + 1)
    IsoToLocalDate = UtcToLocalDate(dtUtc)
End Function

'convert a utc date to local time using the current time zone
Public Function UtcToLocalDate(ByVal dtUtcDate As Date) As Date
    Dim tzi As TIME_ZONE_INFORMATION
    Dim lngBias As Long
    Dim lngResult As Long

    If dtUtcDate = 0 Then Exit Function

    lngResult = GetTimeZoneInformation(tzi)
    If lngResult <> 0 Then
        lngBias = tzi.Bias
        If lngResult = 2 Then lngBias = lngBias + tzi.DaylightBias
        UtcToLocalDate = DateAdd("n", -lngBias, dtUtcDate)
    Else
        UtcToLocalDate = dtUtcDate
    End If
End Function
