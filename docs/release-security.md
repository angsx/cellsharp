# Release checklist

CellSharp releases are intentionally lightweight. Before publishing a package:

- build Release;
- run tests;
- check the package version;
- pack the NuGet package;
- check the NuGet vulnerability report;
- inspect package contents;
- update the changelog or release notes;
- publish the verified package.

For an important release, also verify the symbol package and Source Link when available, and perform an Excel or LibreOffice smoke test.

## Package provenance

Release credentials must not be committed to the repository. Prefer the standard publishing method offered by the package feed and keep any publishing credentials limited to the release action or maintainer environment. Verify the package version and contents immediately before publication.
