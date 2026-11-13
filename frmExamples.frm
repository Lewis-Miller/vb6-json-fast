VERSION 5.00
Begin VB.Form frmExamples 
   Caption         =   "JsonFast & JsonWrite - Complete Usage Examples"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSaveFile 
      Caption         =   "10. Save To UTF-8 File"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   6840
      Width           =   2895
   End
   Begin VB.CommandButton cmdRoundTripDynamic 
      Caption         =   "9. UTF-8 Round-Trip && Discovery"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   6120
      Width           =   2895
   End
   Begin VB.CommandButton cmdPipeline 
      Caption         =   "8. Filter && Transform Pipeline"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CommandButton cmdWriteSerialize 
      Caption         =   "7. Write: Arrays && Dictionaries"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton cmdWriteStream 
      Caption         =   "6. Write: Streaming Builder"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   2895
   End
   Begin VB.CommandButton cmdStrictAndErrors 
      Caption         =   "5. Strict Mode && Error Context"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   2895
   End
   Begin VB.CommandButton cmdRawAndDates 
      Caption         =   "4. Dates && Raw Sub-JSON Slices"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2340
      Width           =   2895
   End
   Begin VB.CommandButton cmdFastWalkers 
      Caption         =   "3. Low-Level Token Walkers"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2895
   End
   Begin VB.CommandButton cmdArrays 
      Caption         =   "2. Arrays && Nested Objects"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdBasic 
      Caption         =   "1. Basic Path && Typed Lookups"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "frmExamples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' sample payload covering nested objects, arrays, booleans, numbers, dates, and escapes
Private Function GetSamplePayload() As String
    Dim s As String
    s = "{" & _
        "  ""service"": ""DomainRegistry""," & _
        "  ""version"": 2.5," & _
        "  ""active"": true," & _
        "  ""retryLimit"": 3," & _
        "  ""createdAt"": ""2026-11-11T14:30:00.25Z""," & _
        "  ""notes"": ""Line 1\nLine 2\tTabbed \""Quoted\""""," & _
        "  ""config"": {" & _
        "    ""timeoutMs"": 5000," & _
        "    ""endpoint"": ""https://example.com""" & _
        "  }," & _
        "  ""records"": [" & _
        "    {""id"": 101, ""domain"": ""alpha.org"", ""verified"": true}," & _
        "    {""id"": 102, ""domain"": ""beta.net"", ""verified"": false}," & _
        "    {""id"": 103, ""domain"": ""gamma.io"", ""verified"": true}" & _
        "  ]" & _
        "}"
    GetSamplePayload = s
End Function

' ---------------------------------------------------------------------------------------
' 1. basic path lookups, default property usage, and typed getters
' ---------------------------------------------------------------------------------------
Private Sub cmdBasic_Click()
    Dim parser As JSonFast
    Dim sNotes As String
    Dim sOut As String

    Set parser = New JSonFast
    parser.Parse GetSamplePayload()

    ' vb6 multiline textboxes require CrLf rather than bare Lf
    sNotes = Replace(parser("notes"), vbCrLf, vbLf)
    sNotes = Replace(sNotes, vbLf, vbCrLf)

    ' Item is the default property, so parser("path") and parser.Item("path") are identical
    sOut = "=== 1. BASIC & TYPED LOOKUPS ===" & vbCrLf & vbCrLf & _
           "Default property : " & parser("service") & vbCrLf & _
           "Explicit .Item   : " & parser.Item("config/endpoint") & vbCrLf & _
           "Auto-unescaped   : " & vbCrLf & sNotes & vbCrLf & vbCrLf & _
           "--- Typed Accessors ---" & vbCrLf & _
           "ItemAsLong       : " & parser.ItemAsLong("retryLimit") & vbCrLf & _
           "ItemAsDouble     : " & parser.ItemAsDouble("version") & vbCrLf & _
           "ItemAsBoolean    : " & parser.ItemAsBoolean("active") & vbCrLf & _
           "Exists (valid)   : " & parser.Exists("config/timeoutMs") & vbCrLf & _
           "Exists (missing) : " & parser.Exists("config/nonExistentKey")

    txtOutput.Text = sOut
End Sub

' ---------------------------------------------------------------------------------------
' 2. iterating through arrays using path indexing
' ---------------------------------------------------------------------------------------
Private Sub cmdArrays_Click()
    Dim parser As JSonFast
    Dim lCount As Long
    Dim i As Long
    Dim sOut As String

    Set parser = New JSonFast
    parser.Parse GetSamplePayload()

    ' GetArrayCount returns the element count for any array path
    lCount = parser.GetArrayCount("records")

    sOut = "=== 2. ARRAYS & NESTED PATHS ===" & vbCrLf & vbCrLf & _
           "Total records found: " & lCount & vbCrLf & vbCrLf

    For i = 0 To lCount - 1
        sOut = sOut & "Record #" & i & ":" & vbCrLf & _
               "  ID       : " & parser.ItemAsLong("records/" & i & "/id") & vbCrLf & _
               "  Domain   : " & parser("records/" & i & "/domain") & vbCrLf & _
               "  Verified : " & parser.ItemAsBoolean("records/" & i & "/verified") & vbCrLf
    Next i

    txtOutput.Text = sOut
End Sub



' ---------------------------------------------------------------------------------------
' 3. high-speed low-level token walkers (avoids path string splitting in tight loops)
' ---------------------------------------------------------------------------------------
Private Sub cmdFastWalkers_Click()
    Dim parser As JSonFast
    Dim itemIdx As Long
    Dim idIdx As Long
    Dim domIdx As Long
    Dim keyIdx As Long
    Dim valIdx As Long
    Dim secondRecIdx As Long
    Dim sOut As String

    Set parser = New JSonFast
    parser.Parse GetSamplePayload()

    sOut = "=== 3. LOW-LEVEL TOKEN WALKERS ===" & vbCrLf & _
           "(best for large arrays/objects where Split() overhead matters)" & vbCrLf & vbCrLf

    ' A. EnumerateArray jumps directly to the first child token of the array
    sOut = sOut & "--- EnumerateArray(""records"") ---" & vbCrLf
    itemIdx = parser.EnumerateArray("records")

    Do While itemIdx <> -1
        idIdx = parser.GetObjectValue(itemIdx, "id")
        domIdx = parser.GetObjectValue(itemIdx, "domain")

        If idIdx <> -1 And domIdx <> -1 Then
            sOut = sOut & "Token[" & itemIdx & "] -> ID=" & parser.GetValueString(idIdx) & _
                   ", Domain=" & parser.GetValueString(domIdx) & vbCrLf
        End If

        itemIdx = parser.GetNextSibling(itemIdx)
    Loop

    ' B. GetArrayItem grabs a specific element by index directly from an array token
    sOut = sOut & vbCrLf & "--- GetArrayItem (direct index 1) ---" & vbCrLf
    secondRecIdx = parser.GetArrayItem(GetRecordsArrayToken(parser), 1&)
    If secondRecIdx <> -1 Then
        domIdx = parser.GetObjectValue(secondRecIdx, "domain")
        sOut = sOut & "Record #1 domain: " & parser.GetValueString(domIdx) & vbCrLf
    End If

    ' C. EnumerateObject walks all key/value pairs inside an object path
    sOut = sOut & vbCrLf & "--- EnumerateObject(""config"") ---" & vbCrLf
    keyIdx = parser.EnumerateObject("config")
    Do While keyIdx <> -1
        valIdx = parser.GetFirstChild(keyIdx)
        sOut = sOut & parser.GetValueString(keyIdx) & " = " & parser.GetValueString(valIdx) & vbCrLf
        keyIdx = parser.GetNextSibling(keyIdx)
    Loop

    txtOutput.Text = sOut
End Sub

' helper to locate the "records" array container token from root
Private Function GetRecordsArrayToken(ByRef parser As JSonFast) As Long
    Dim keyIdx As Long
    GetRecordsArrayToken = -1
    keyIdx = parser.GetFirstChild(0&)
    Do While keyIdx <> -1
        If parser.GetValueString(keyIdx) = "records" Then
            GetRecordsArrayToken = parser.GetFirstChild(keyIdx)
            Exit Function
        End If
        keyIdx = parser.GetNextSibling(keyIdx)
    Loop
End Function

' ---------------------------------------------------------------------------------------
' 4. zero-allocation ISO 8601 date parsing & raw sub-tree slicing
' ---------------------------------------------------------------------------------------
Private Sub cmdRawAndDates_Click()
    Dim parser As JSonFast
    Dim dtUtc As Date
    Dim dtLocal As Date
    Dim sRawObj As String
    Dim sOut As String

    Set parser = New JSonFast
    parser.Parse GetSamplePayload()

    ' ItemAsDate reads raw bytes directly without intermediate string allocations
    dtUtc = parser.ItemAsDate("createdAt", False)
    dtLocal = parser.ItemAsDate("createdAt", True)

    ' GetRaw extracts the exact unparsed JSON slice for an object, array, or scalar
    sRawObj = parser.GetRaw("config")

    sOut = "=== 4. DATES & RAW SUB-JSON SLICES ===" & vbCrLf & vbCrLf & _
           "Raw ISO string   : " & parser("createdAt") & vbCrLf & _
           "Parsed UTC Date  : " & Format$(dtUtc, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
           "Local Timezone   : " & Format$(dtLocal, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf & _
           "Raw 'config' JSON slice:" & vbCrLf & _
           sRawObj

    txtOutput.Text = sOut
End Sub



' ---------------------------------------------------------------------------------------
' 5. strict mode validation & byte-stream ingestion
' ---------------------------------------------------------------------------------------
Private Sub cmdStrictAndErrors_Click()
    Dim parser As JSonFast
    Dim bUtf8() As Byte
    Dim sOut As String

    Set parser = New JSonFast

    ' scenario A: parsing a raw UTF-8 byte array directly (skips WideCharToMultiByte)
    bUtf8 = StrConv("{""directBytes"": true}", vbFromUnicode)
    parser.Parse bUtf8
    sOut = "=== 5. STRICT MODE & BYTE ARRAYS ===" & vbCrLf & vbCrLf & _
           "Direct Byte() parse result: " & parser.ItemAsBoolean("directBytes") & vbCrLf & vbCrLf

    ' scenario B: strict mode catches trailing garbage after root close
    parser.Strict = True
    sOut = sOut & "Test 1 - Trailing garbage after root:" & vbCrLf & _
           TryParseStrict(parser, "{""valid"": 123} trailing_garbage") & vbCrLf & vbCrLf

    ' scenario C: strict mode rejects malformed literals (e.g. "truX" instead of "true")
    sOut = sOut & "Test 2 - Malformed literal (""truX""):" & vbCrLf & _
           TryParseStrict(parser, "{""flag"": truX}")

    txtOutput.Text = sOut
End Sub

' helper to test strict parsing errors cleanly without halting subsequent test cases
Private Function TryParseStrict(ByRef parser As JSonFast, ByVal sJson As String) As String
    On Error GoTo ParseError
    parser.Parse sJson
    TryParseStrict = "  Passed (unexpected)"
    Exit Function

ParseError:
    TryParseStrict = "  Caught Error #" & Err.Number & ": " & Err.Description
    Err.Clear
End Function

' ---------------------------------------------------------------------------------------
' 6. JsonWrite streaming builder API + default Item property + Clear() reuse
' ---------------------------------------------------------------------------------------
Private Sub cmdWriteStream_Click()
    Dim writer As JSonWrite
    Dim formatter As JSonFast
    Dim sFirstCompact As String
    Dim sSecondCompact As String
    Dim sPrettyPrint As String
    Dim sThirdCompact As String
    Dim sOut As String

    Set writer = New JSonWrite
    Set formatter = New JSonFast

    'this demonstrates 2 different methods for writing json data

    ' A. explicit streaming methods (.Add / .AddValue)
    writer.BeginObject
    writer.Add "app", "RegistrySync"
    writer.Add "uptimeSec", 8420
    writer.Add "healthy", True
    writer.Add "timestamp", Now
    writer.Add "optionalTag", Null

    writer.WriteKey "hosts"
    writer.BeginArray
    writer.AddValue "http://example.com"
    writer.AddValue "http://example.com"
    writer.EndArray
    writer.EndObject

    sFirstCompact = writer.ToString()

    ' B. reset cursor with .Clear (reuses pre-allocated buffer) & use default Item property
    writer.Clear

    writer.BeginObject
    writer("service") = "DomainRegistry"
    writer("version") = 2.5
    writer("active") = True

    writer.WriteKey "hosts"
    writer.BeginArray
    writer("") = "http://example.com"
    writer("") = "http://example.com"
    writer.EndArray
    writer.EndObject

    sSecondCompact = writer.ToString()

    ' C. zero-allocation typed streaming methods (bypasses Variant boxing)
    writer.Clear
    writer.RefreshTimeZone   ' optional: refreshes cached OS time zone bias after DST change

    writer.BeginObject
        writer.AddString "service", "DomainRegistry"
        writer.AddLong "uptimeSec", 8420&
        writer.AddBoolean "healthy", True
        writer.AddNull "lastError"
    writer.EndObject

    sPrettyPrint = formatter.PrettyPrint(writer.ToString())
    sThirdCompact = formatter.Minify(sPrettyPrint)
    
    sOut = "=== 6. JSONWRITE STREAMING BUILDER ===" & vbCrLf & vbCrLf & _
           "--- Style A: Explicit .Add / .AddValue ---" & vbCrLf & _
           sFirstCompact & vbCrLf & vbCrLf & _
           "--- Style B: Default Item Property (after .Clear) ---" & vbCrLf & _
           "Compact (" & writer.Length & " chars):" & vbCrLf & _
           sSecondCompact & vbCrLf & vbCrLf & _
           "Formatted via JsonFast.PrettyPrint:" & vbCrLf & _
           formatter.PrettyPrint(sSecondCompact) & vbCrLf & vbCrLf & _
           "--- Style C: zero-allocation typed streaming builder ---" & vbCrLf & _
           "PrettyPrint (" & Len(sPrettyPrint) & " chars):" & vbCrLf & _
           sPrettyPrint & vbCrLf & vbCrLf & _
           "Compacted via JsonFast.Minify: " & Len(sThirdCompact) & " Chars." & vbCrLf & _
           sThirdCompact
           
           

    txtOutput.Text = sOut
End Sub


' ---------------------------------------------------------------------------------------
' 7. JsonWrite hybrid Serialize() with 2D arrays, 1D arrays, and Scripting.Dictionary
' ---------------------------------------------------------------------------------------
Private Sub cmdWriteSerialize_Click()
    Dim writer As JSonWrite
    Dim formatter As JSonFast
    Dim arrObj(0 To 1, 0 To 2) As Variant
    Dim dict As Object
    Dim sOut As String

    Set writer = New JSonWrite
    Set formatter = New JSonFast

    ' A. 2D Variant array (row 0 = keys, row 1 = values) serializes as a JSON object
    arrObj(0, 0) = "user": arrObj(1, 0) = "Jane Doe"
    arrObj(0, 1) = "roles": arrObj(1, 1) = Array("admin", "auditor")
    arrObj(0, 2) = "active": arrObj(1, 2) = True

    sOut = "=== 7. SERIALIZE ARRAYS & DICTIONARIES ===" & vbCrLf & vbCrLf & _
           "--- 2D Key/Value Array -> JSON Object ---" & vbCrLf & _
           formatter.PrettyPrint(writer.Serialize(arrObj)) & vbCrLf & vbCrLf

    ' B. late-bound Scripting.Dictionary (no scrrun.dll reference needed)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "environment", "production"
    dict.Add "owner", arrObj

    sOut = sOut & "--- Scripting.Dictionary + Nested 2D Array ---" & vbCrLf & _
           formatter.PrettyPrint(writer.Serialize(dict))

    txtOutput.Text = sOut
End Sub

' ---------------------------------------------------------------------------------------
' 8. parse with JsonFast, filter records, write new JSON with JsonWrite
' ---------------------------------------------------------------------------------------
Private Sub cmdPipeline_Click()
    Dim parser As JSonFast
    Dim writer As JSonWrite
    Dim itemIdx As Long
    Dim verIdx As Long
    Dim idIdx As Long
    Dim domIdx As Long
    Dim lVerifiedCount As Long
    Dim sSummaryJson As String
    Dim sOut As String

    Set parser = New JSonFast
    Set writer = New JSonWrite

    parser.Parse GetSamplePayload()

    ' build a new filtered summary payload while walking the parsed token tree
    writer.BeginObject
    writer.Add "sourceService", parser("service")
    writer.WriteKey "verifiedDomains"
    writer.BeginArray

    itemIdx = parser.EnumerateArray("records")
    Do While itemIdx <> -1
        verIdx = parser.GetObjectValue(itemIdx, "verified")
        If verIdx <> -1 Then
            If parser.GetTokenType(verIdx) = TOK_TRUE Then
                idIdx = parser.GetObjectValue(itemIdx, "id")
                domIdx = parser.GetObjectValue(itemIdx, "domain")

                writer.BeginObject
                writer.Add "id", CLng(Val(parser.GetValueString(idIdx)))
                writer.Add "domain", parser.GetValueString(domIdx)
                writer.EndObject
                lVerifiedCount = lVerifiedCount + 1
            End If
        End If
        itemIdx = parser.GetNextSibling(itemIdx)
    Loop

    writer.EndArray
    writer.Add "verifiedCount", lVerifiedCount
    writer.EndObject

    sSummaryJson = writer.ToString()

    sOut = "=== 8. FILTER & TRANSFORM PIPELINE ===" & vbCrLf & _
           "(JsonFast reads -> filters verified=true -> JsonWrite builds output)" & vbCrLf & vbCrLf & _
           parser.PrettyPrint(sSummaryJson)

    txtOutput.Text = sOut
End Sub

' ---------------------------------------------------------------------------------------
' 9. zero-copy UTF-8 Byte() round-trip + dynamic schema discovery
' ---------------------------------------------------------------------------------------
Private Sub cmdRoundTripDynamic_Click()
    Dim writer As JSonWrite
    Dim parser As JSonFast
    Dim bUtf8() As Byte
    Dim sOut As String

    Set writer = New JSonWrite
    Set parser = New JSonFast

    ' build payload and export directly to UTF-8 Byte() (skips BSTR -> UTF-8 re-encoding)
    writer.BeginObject
    writer.Add "batchId", 9001
    writer.Add "status", "ok"
    writer.WriteKey "metrics"
    writer.BeginArray
    writer.AddValue 12.5
    writer.AddValue 48.75
    writer.AddValue True
    writer.EndArray
    writer.EndObject

    bUtf8 = writer.ToUtf8Bytes()

    ' ingest raw UTF-8 bytes directly into JsonFast
    parser.Strict = True
    parser.Parse bUtf8

    sOut = "=== 9. UTF-8 ROUND-TRIP & DYNAMIC DISCOVERY ===" & vbCrLf & _
           "(JsonWrite.ToUtf8Bytes -> JsonFast.Parse Byte() -> WalkTokenTree)" & vbCrLf & vbCrLf & _
           "UTF-8 byte payload size: " & (UBound(bUtf8) + 1) & " bytes" & vbCrLf & vbCrLf

    WalkTokenTree parser, 0&, 0&, sOut
    txtOutput.Text = sOut
End Sub

' walks container nodes (objects and arrays) horizontally across their sibling chains.
' needed because JsonFast stores tokens in flat parallel arrays rather than nested objects,
' so discovering unknown schemas requires stepping from child to sibling via token indices.
Private Sub WalkTokenTree(ByRef parser As JSonFast, ByVal tokenIdx As Long, ByVal depth As Long, ByRef sOut As String)
    Dim tType As Byte
    Dim childIdx As Long
    Dim valIdx As Long
    Dim sPad As String
    Dim i As Long

    If tokenIdx = -1 Then Exit Sub
    sPad = String$(depth * 2&, " ")
    tType = parser.GetTokenType(tokenIdx)

    Select Case tType
        Case TOK_OBJECT_OPEN
            childIdx = parser.GetFirstChild(tokenIdx)
            Do While childIdx <> -1
                valIdx = parser.GetFirstChild(childIdx)
                sOut = sOut & sPad & "[Key] " & parser.GetValueString(childIdx) & " -> "
                DescribeValueNode parser, valIdx, depth, sOut
                childIdx = parser.GetNextSibling(childIdx)
            Loop

        Case TOK_ARRAY_OPEN
            i = 0
            childIdx = parser.GetFirstChild(tokenIdx)
            Do While childIdx <> -1
                sOut = sOut & sPad & "[" & i & "] -> "
                DescribeValueNode parser, childIdx, depth, sOut
                i = i + 1
                childIdx = parser.GetNextSibling(childIdx)
            Loop
    End Select
End Sub

' inspects a single value token, prints its type/content, and recurses if it is a container.
' separated from WalkTokenTree so both object key-values and array elements share the same
' mutual-recursion logic without duplicating type checks.
Private Sub DescribeValueNode(ByRef parser As JSonFast, ByVal valIdx As Long, ByVal depth As Long, ByRef sOut As String)
    Select Case parser.GetTokenType(valIdx)
        Case TOK_OBJECT_OPEN
            sOut = sOut & "{Object}" & vbCrLf
            WalkTokenTree parser, valIdx, depth + 1&, sOut
        Case TOK_ARRAY_OPEN
            sOut = sOut & "[Array]" & vbCrLf
            WalkTokenTree parser, valIdx, depth + 1&, sOut
        Case TOK_STRING
            sOut = sOut & "(String) " & Replace(parser.GetValueString(valIdx), vbLf, "\n") & vbCrLf
        Case TOK_NUMBER
            sOut = sOut & "(Number) " & parser.GetValueString(valIdx) & vbCrLf
        Case TOK_TRUE, TOK_FALSE
            sOut = sOut & "(Bool) " & parser.GetValueString(valIdx) & vbCrLf
        Case TOK_NULL
            sOut = sOut & "(Null)" & vbCrLf
    End Select
End Sub

' ---------------------------------------------------------------------------------------
' 10. Save to UTF-8 File
' ---------------------------------------------------------------------------------------

Private Sub cmdSaveFile_Click()
    Dim writer As JSonWrite
    Dim sPath As String
    Dim sOut As String

    Set writer = New JSonWrite
    sPath = App.path & "\export_demo.json"

    writer.BeginObject
        writer.AddString "exportedBy", "JSonWrite"
        writer.AddLong "recordCount", 3&
        writer.AddBoolean "utf8NoBom", True
        writer.Add "generatedAt", Now
    writer.EndObject

    ' Write UTF-8 directly to disk (supports full Unicode paths, defaults to no BOM)
    writer.SaveToFile sPath, False

    sOut = "=== 10. DIRECT UTF-8 FILE EXPORT ===" & vbCrLf & vbCrLf & _
           "Saved to  : " & sPath & vbCrLf & _
           "File size : " & FileLen(sPath) & " bytes" & vbCrLf & vbCrLf & _
           "Payload:" & vbCrLf & _
           writer.ToString()

    ' Clean up temporary demo file
    If Len(Dir$(sPath)) > 0 Then Kill sPath

    txtOutput.Text = sOut
End Sub


