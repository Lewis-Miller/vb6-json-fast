# JSonFast & JSonWrite for Visual Basic 6

A high-performance, zero-dependency JSON parser, query engine, and serializer written in pure Visual Basic 6.

Designed using data-oriented parallel scalar arrays (`Byte()` and `Long()` vectors), direct UTF-16 integer buffer writes via `CopyMemory`, and zero-copy `SAFEARRAY` views, **JsonFast** and **JsonWrite** eliminate COM object overhead, `Variant` tree allocations, and `BSTR` string concatenation churn—achieving **330–360 MB/s** parsing throughput and **~50 MB/s** serialization throughput in native compiled mode.

## Features

### `JSonFast.cls` (Parser & Query Engine)
* **Zero-Allocation Lazy Evaluation:** Tokenizes raw UTF-8 byte streams into flat parallel arrays and decodes strings on demand only when queried.
* **Path & Typed Lookups:** Default `Item` property (`parser("user/name")`) with automatic string unescaping, plus `ItemAsLong`, `ItemAsDouble`, `ItemAsBoolean`, `ItemAsDate` (direct-byte ISO 8601 parsing with UTC/local timezone conversion), `GetRaw`, and `Exists`.
* **Low-Level Token Walkers:** `EnumerateArray`, `EnumerateObject`, `GetArrayItem`, `GetObjectValue`, `GetFirstChild`, and `GetNextSibling` for `O(1)` sibling traversal without path-splitting overhead.
* **Strict Validation Mode:** Optional `Strict = True` flag to validate literals (`true`, `false`, `null`) and reject trailing data after the root container.
* **Formatting Utilities:** Fast 2-byte UTF-16 `Integer()` array implementations of `PrettyPrint`, `Minify`, `Escape`, and `UnEscape`.

### `JSonWrite.cls` (Serializer)
* **Streaming Builder API:** `BeginObject`, `BeginArray`, `Add`, `AddValue`, `WriteKey`, `EndObject`, and `EndArray` write directly into a pre-allocated UTF-16 `Integer()` buffer with zero intermediate objects.
* **Typed Streaming Fast-Paths:** `AddString`, `AddLong`, `AddBoolean`, and `AddNull` bypass `Variant` boxing and runtime dispatch for maximum throughput on hot loops.
* **Zero-Copy Hybrid `Serialize(vData)`:** Uses direct `SAFEARRAY` `pvData` views and cached COM `vtable` calls (`DispCallFunc`) to serialize 2D key-value `Variant` arrays (`{object}`), 1D `Variant` arrays (`[array]`), late-bound `Scripting.Dictionary` trees (no `scrrun.dll` reference required), and scalars (`String`, numbers, `Boolean`, `Date`, `Null`).
* **Cached Time Zone & Direct Date Formatting:** Caches OS time zone bias (`RefreshTimeZone`) and formats ISO 8601 UTC timestamps directly into the output buffer without `Format$` string allocations.
* **Direct UTF-8 Output & File Export:** `ToString()` for VB6 strings, `ToUtf8Bytes()` for socket/memory buffers, and `SaveToFile(path, [writeBom])` for direct Unicode-safe UTF-8 disk writes via Win32 `CreateFileW`.

### `frmExamples.frm` (Interactive Showcase)
* Includes 10 runnable examples covering basic/typed lookups, array iteration, low-level token walkers, ISO dates, raw sub-tree slicing, strict mode error handling, streaming writes (including typed fast-paths), array/dictionary serialization, zero-copy UTF-8 round-trip pipelines, and direct UTF-8 file export.

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
| Streaming Builder — Typed (`AddString` / `AddLong` / `AddBoolean`) | **50.70 MB/s** |
| Streaming Builder — Variant (`BeginObject` / `Add`) | **44.48 MB/s** |
| 2D Key-Value `Variant` Array (`Serialize`) | **42.38 MB/s** |
| Late-Bound `Scripting.Dictionary` (`Serialize`) | **20.00 MB/s** |

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
    writer.AddString "service", "DomainRegistry"
    writer.Add "version", 2.5
    writer.AddBoolean "active", True
    writer.WriteKey "hosts"
    writer.BeginArray
        writer.AddValue "https://example.com"
        writer.AddValue "https://example.com"
    writer.EndArray
writer.EndObject

Debug.Print writer.ToString()
writer.SaveToFile App.Path & "\output.json"   ' Direct UTF-8 file export (no BOM)
```

## Installation

1. Copy `JSonFast.cls` and `JSonWrite.cls` into your VB6 project directory and add them via **Project -> Add File...**
2. *(Optional)* Add `frmExamples.frm` to explore all 10 interactive usage scenarios.
3. No external DLLs, OCXs, or project references are required—only standard Windows APIs (`kernel32`, `oleaut32`, `msvbvm60`).
