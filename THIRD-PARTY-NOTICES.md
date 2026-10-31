# Third-party notices

XlDuck is MPL-2.0 software. Its dependencies retain their own licences.

The distribution includes these notice sets:

| Components | Notices |
| --- | --- |
| Excel-DNA, DuckDB.NET, Apache Arrow and WebView2 | `licenses/`: licence files and original NuGet metadata copied from the locked packages |
| DuckDB and its bundled algorithms, codecs, ICU and Parquet libraries | `third-party/notices/duckdb-1.5.6/` |
| DuckDB HTTPFS and AWS credential-provider extensions | `third-party/notices/duckdb-httpfs/` and `duckdb-aws/` |
| Native reader dependencies: curl, OpenSSL, zlib, AWS SDK and AWS CRT libraries | Other component directories under `third-party/notices/` |
| Excel-DNA and its LZMA resources | `third-party/notices/ExcelDna-1.9.0/` |
| Vega, Vega-Lite, Vega Embed, Vega Interpreter, Leaflet and SQL Formatter | `preview/licenses/`, alongside the pinned renderer assets |

`third-party/notices.json` records the exact upstream URL and SHA-256 of each additional notice. These files are checked into the repository and included in every package; packaging verifies them before building. NuGet hashes appear in `licenses/nuget-packages.json`, preview package versions/integrity in `preview/versions.json`, and reader binary hashes in `scripts/Get-ReaderExtensions.ps1` in the source repository.

The native notice inventory follows [DuckDB v1.5.6](https://github.com/duckdb/duckdb/tree/v1.5.6), including its vendored notices and the HTTPFS and AWS revisions pinned by that release. The reader dependency inventory follows the AWS extension's [vcpkg baseline](https://github.com/microsoft/vcpkg/tree/84bab45d415d22042bd0b9081aea57f362da3f35/ports); it includes a conservative superset of native dependencies. It is a notice inventory, not a claim that every listed library is linked into every binary. For dual-licensed components offering Apache-2.0 or MIT as alternatives, this distribution uses the permissive option. The complete upstream texts are preserved, including their alternative terms.

The Windows/.NET Desktop and WebView2 runtimes are separate prerequisites supplied by Microsoft; XlDuck does not redistribute their installers. WebView2 SDK notices accompany the host libraries. Excel's interop types are embedded at build time; the original ExcelDna.Interop NuGet metadata is retained for provenance.

When upgrading a native dependency, update the corresponding notice files and inventory as well as the binary pins. Review newly introduced dependencies and upstream NOTICE files; do not treat a package's top-level licence alone as the entire dependency inventory.
