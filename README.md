# JSonFast & JSonWrite for Visual Basic 6

A high-performance, zero-dependency JSON parser and serializer for Visual Basic 6. Designed for people who need to parse JSON in VB6 without pulling in COM wrappers or building a full object tree.

Most VB6 JSON options either wrap an external library or allocate a Dictionary/Collection for every object and array. JSonFast instead tokenizes the raw UTF-8 bytes into flat parallel arrays and only materializes strings when you ask for them, reaching **330–360 MB/s** parse throughput in native compiled mode. JSonWrite does the reverse, writing directly into a pre-allocated UTF-16 buffer.

## Features

### `JSonFast.cls` (Parser & Query Engine)
* **Low-overhead token stream:** Tokenizes raw UTF-8 into parallel arrays (`TokType`, `TokOffset`, `TokLength`, `TokNext`, `TokChild`, `TokTail`) and decodes strings only when queried.
* **Path & Typed Lookups:** Default `Item` property (`parser("user/name")`) with automatic string unescaping, plus `ItemAsLong`, `ItemAsDouble`, `ItemAsBoolean`, `ItemAsDate` (direct-byte ISO 8601 parsing with UTC/local timezone conversion), `GetRaw`, and `Exists`.
* **Low-Level Token Walkers:** `EnumerateArray`, `EnumerateObject`, `GetArrayItem`, `GetObjectValue`, `GetFirstChild`, and `GetNextSibling` for O(1) sibling traversal without path-splitting overhead.
* **Strict Validation Mode:** Optional `Strict = True` flag to validate literals (`true`, `false`, `null`) and reject trailing data after the root container.
* **Formatting Utilities:** Fast 2-byte UTF-16 `Integer()` array implementations of `PrettyPrint`, `Minify`, `Escape`, and `UnEscape`.

### `JSonWrite.cls` (Serializer)
* **Streaming Builder API:** `BeginObject`, `BeginArray`, `Add`, `AddValue`, `WriteKey`, `EndObject`, and `EndArray` write directly into a pre-allocated UTF-16 `Integer()` buffer with no intermediate objects.
* **Default `Item` property:** Assign with `writer("key") = value` inside an object or `writer("") = value` inside an array. `Property Set` supports object values such as `Scripting.Dictionary`.
* **Hybrid `Serialize(vData)`:** Automatically serializes 2D key-value `Variant` arrays (`{object}`), 1D `Variant` arrays (`[array]`), late-bound `Scripting.Dictionary` trees (no `scrrun.dll` reference required), and scalars (`String`, numbers, `Boolean`, `Date`, `Null`).
* **Direct UTF-8 Output:** `ToString()` for VB6 strings or `ToUtf8Bytes()` for direct file/socket output.

### `frmExamples.frm` (Interactive Showcase)
* Includes 9 runnable examples covering basic/typed lookups, array iteration, low-level token walkers, ISO dates, raw sub-tree slicing, strict mode error handling, streaming writes, array/dictionary serialization, and UTF-8 round-trip pipelines.

## Performance Benchmarks

Measured in native compiled mode (with Advanced Optimizations enabled: *Remove Array Bounds Checks*, *Assume No Aliasing*, *Remove Integer Overflow Checks*):

### `JSonFast` Multi-File Parse Throughput

| File Name | Size (MB) | Runs | Avg / Parse | Throughput |
| :--- | :---: | :---: | :---: | :---: |
| `small.json` | 0.0109 | 500 | 0.04 ms | **280.52 MB/s** |
| `1mb.json` | 1.0969 | 25 | 3.63 ms | **302.27 MB/s** |
| `15mb.json` | 14.6867 | 5 | 49.88 ms | **294.43 MB/s** |
| `large.json` | 24.9303 | 5 | 69.15 ms | **360.54 MB/s** |
| `40mb.json` | 39.3201 | 2 | 130.09 ms | **302.24 MB/s** |
| `100mb.json` | 102.9777 | 2 | 308.74 ms | **333.54 MB/s** |

### `JSonWrite` Serialization Throughput (10,000 Records / ~0.95 MB)

| Write Mode | Throughput |
| :--- | :---: |
| Streaming Builder (`BeginObject` / `Add`) | **12.06 MB/s** |
| 2D Key-Value `Variant` Array (`Serialize`) | **7.81 MB/s** |
| Late-Bound `Scripting.Dictionary` (`Serialize`) | **1.06 MB/s** |

## Quick Start

### Reading JSON (`JSonFast`)
```vb
Dim parser As New JSonFast
parser.Parse "{""user"":""Jane Doe"",""active"":true,""created"":""2026-11-11T14:30:00Z""}"

Debug.Print parser("user")                        ' "Jane Doe" (Item is default property)
Debug.Print parser.ItemAsBoolean("active")        ' True
Debug.Print parser.ItemAsDate("created", True)    ' Local VB6 Date
```

### Walking Large Arrays Without Path Overhead
```vb
Dim itemIdx As Long
Dim nameIdx As Long

itemIdx = parser.EnumerateArray("records")
Do While itemIdx <> -1
    nameIdx = parser.GetObjectValue(itemIdx, "domain")
    If nameIdx <> -1 Then Debug.Print parser.GetValueString(nameIdx)
    itemIdx = parser.GetNextSibling(itemIdx)
Loop
```

### Writing JSON (`JSonWrite`)
```vb
Dim writer As New JSonWrite

writer.BeginObject
    writer("service") = "DomainRegistry"      ' default Item property
    writer("version") = 2.5
    writer("active") = True
    writer.WriteKey "hosts"
    writer.BeginArray
        writer("") = "http://example.com"    ' empty key = array value
        writer("") = "http://example.com"
    writer.EndArray
writer.EndObject

Debug.Print writer.ToString()
```

## Installation

1. Copy `JSonFast.cls` and `JSonWrite.cls` into your VB6 project directory and add them via **Project -> Add File...**
2. *(Optional)* Add `frmExamples.frm` and open `Project1.vbp` to explore all 9 interactive usage scenarios.
3. No external DLLs, OCXs, or project references are required—only standard `kernel32` APIs (`CopyMemory`, `MultiByteToWideChar`, `WideCharToMultiByte`, `GetTimeZoneInformation`).
