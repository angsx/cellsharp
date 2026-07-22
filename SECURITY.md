# Security Policy

## Supported versions

CellSharp is pre-1.0. Security fixes are considered for the current maintained version only. There are no supported historical release lines at this time; this policy can be expanded when multiple release lines exist.

## Reporting a vulnerability

Do not open a public issue for an undisclosed vulnerability.

Use [GitHub Private Vulnerability Reporting](https://github.com/angsx/cellsharp/security/advisories/new) if it is enabled for the repository. If that option is unavailable, contact Angelo Collura through a private contact method published on the project or maintainer GitHub profile. No security email address is published by this repository.

Please include:

- A clear description of the vulnerability and its impact.
- Affected CellSharp versions and target frameworks.
- Steps to reproduce, including a proof of concept when safe to share.
- Any known mitigations or workarounds.

## Disclosure

CellSharp follows coordinated, responsible disclosure. Please allow time to investigate and prepare a fix before public disclosure. No response-time commitment is made.

## Security guarantees and trust boundaries

The maintained library uses finite resource limits for XLSX package size, Open XML part size, physical rows, collected errors, and columns. It writes ordinary strings as literal cells and never evaluates workbook formulas, macros, external links, or embedded objects. See the detailed [security model](docs/security.md), including the explicit trust boundaries around caller-selected paths, custom delegates, raw diagnostics, and native `Formula(...)` output.
