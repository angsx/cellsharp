# Security model

CellSharp treats workbook contents as untrusted data. It does not execute VBA macros, formulas, external links, or embedded objects. Formula cells are imported only from their cached scalar value; formulas without a cache are data errors. Normal strings are always written as literal inline text, including strings beginning with `=`, `+`, `-`, or `@`.

## Resource limits

Every read uses bounded defaults to reduce denial-of-service risk from oversized or highly compressed Open XML packages:

- compressed package: 64 MiB;
- XML characters in any one Open XML part: 32 MiB;
- physical rows inspected in one worksheet: 100,000;
- collected data errors in one worksheet: 1,000;
- physical columns: Excel's fixed maximum of 16,384 (`XFD`).

Exceeding a package, row, error, or column limit throws a structural input exception. The Open XML SDK rejects an oversized XML part while loading it. These failures are not returned as row diagnostics because the input cannot be processed safely or completely.

Applications with a legitimate larger workload can opt in explicitly:

```csharp
var options = new ExcelReadOptions(
    maxRows: 250_000,
    maxErrors: 5_000,
    maxPackageBytes: 128L * 1024 * 1024,
    maxCharactersInPart: 64L * 1024 * 1024);

var result = Excel.Read<Order>(stream, options);
```

For multi-sheet reads, pass the limits when opening the package so the Open XML part limit is established before any part is loaded:

```csharp
using var workbook = Excel.Open(stream, options);
var result = workbook.Read(orderSchema);
```

Raise limits only for inputs whose provenance and expected size justify the additional memory and CPU exposure. A cancellation token is complementary; it is not a substitute for resource limits.

## Trust boundaries

- File paths and streams are caller-selected capabilities. CellSharp does not authorize or sandbox filesystem access; an application must ensure an untrusted user cannot choose arbitrary server paths.
- `ExcelReadError.RawValue`, worksheet headers, and sheet names can originate in an uploaded workbook. Encode them for the destination before rendering in HTML, SQL, shell commands, logs, or another interpreter.
- Custom converters, validation delegates, schema overlays, write row enumerables, and formula callbacks are application code and execute with the host process's permissions.
- `Formula(...)` deliberately writes executable spreadsheet expressions. Use only constant/application-authored formulas and never concatenate untrusted text into a formula. Put untrusted values in ordinary cells and reference those cells from a fixed formula.
- Import validates explicit physical cell references, row bounds, duplicate physical columns, and cell/row metadata conflicts. It does not trust a worksheet dimension declaration when choosing rows to inspect.
- `Hyperlink(...)` writes only HTTP(S) external relationships or non-empty internal locations. CellSharp never follows workbook relationships or downloads external workbook resources.
- Shared strings are loaded eagerly because cell references can address them by index; their XML part remains subject to `MaxCharactersInPart`, and missing or invalid shared-string references are rejected as malformed input. Keep the default limits for public uploads unless a larger trusted feed requires an explicit override.
- CellSharp validates the structures it consumes, but it is not malware detection. If an application stores or redistributes original uploads, scan and govern those original files independently.

## Deployment guidance

Keep CellSharp and its transitive dependencies updated, retain NuGet auditing in CI, use the committed lock files, and process public uploads with an application-level request-size limit and a cancellation deadline. Run especially exposed document-processing workloads with least privilege and without unnecessary filesystem or network access.

Maintainers should follow the lightweight [release checklist](release-security.md) and verify the package contents before publication.
