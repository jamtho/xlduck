# Read and compute policy and connections

The new engine accepts research queries and supported data readers while rejecting SQL that changes data, installs capabilities, or changes the session configuration. This contract applies to the production Excel host, its workbook worker and all output/preview paths. The production host provides Excel Protected View enforcement and connection settings.

## Access model

An ordinary workbook has broad read access under the user's account and configured credentials. There are no per-query, folder or bucket permissions. The host creates the session with `ConnectionSettings`; worksheet SQL cannot change these settings. The default session has external access disabled and can compute using values, captures and supplied handles. Enabling external reads also enables access to local files; adding the installed HTTP reader enables remote files.

This policy prevents supported SQL mutation paths. It does not make an unfamiliar workbook safe to run without a trust decision: readable data can appear in results or be sent in an outgoing URL. A remote service can also attach side effects to a GET request. The worker runs under the user's account for fault isolation; it is not an OS security sandbox. Read-only source credentials remain useful defence in depth.

The production host creates sessions automatically for ordinary workbooks and blocks execution in Protected View. Excel's Enable Editing decision admits a protected workbook. Ordinary copies and externally modified workbooks run without a second permission decision. Connections are user-local settings; explicit keys are loaded from Windows Credential Manager and passed into temporary scoped secrets.

## Query contract

Every query is parsed with DuckDB's SQL serializer before binding or execution. Exactly one SELECT query is accepted, including supported CTEs, joins, windows, aggregates, lambdas and explicit-column pivots. A SELECT prefix alone is insufficient: functions can execute more SQL, change settings or create auxiliary tables.

The policy inspects the entire parsed tree. Scalar and aggregate functions come from the reviewed, pinned built-in catalog with explicit exclusions for dynamic SQL/binding and internal functions. Random values and UUIDs remain available and follow normal result reuse. Scalar macros are admitted only when every overload expands into supported capabilities. Table functions have an explicit allowlist. Extensions cannot silently expand this surface.

| Capability | Current behaviour |
| --- | --- |
| CSV, JSON, newline-delimited JSON and Parquet | Explicit `read_csv`, `read_json`, `read_ndjson`, `read_parquet` and their listed aliases |
| Raw text, blobs and file discovery | `read_text`, `read_blob`, `glob`, `sniff_csv` |
| Parquet metadata | Supported metadata, schema and bloom-probe functions |
| In-memory computation | Supplied table/fragment handles, captures, CTEs, range/series, unnest and other admitted built-ins |
| Catalog inspection | Selected table, column, type, function, setting and extension metadata |
| Worksheet writes and configuration | No DDL, DML, COPY/export, ATTACH, PRAGMA, SET, INSTALL, LOAD, macros, sequences or secret creation |
| Dynamic SQL and arbitrary readers | No `query`, `query_table`, serialized SQL execution, foreign database queries or unreviewed extensions |
| CSV reject tables | Reject-table options are unavailable because they create mutable auxiliary state outside result ownership |

The exact function lists and exclusions are maintained in [ReadComputePolicy.cs](ReadComputePolicy.cs). Internal result creation, dropping tables and connection setup use separate, host-owned commands. User SQL is never treated as one of those commands.

For an Excel coding agent:

- Use positional `?` arguments for values and table/fragment handles. Do not interpolate values into SQL.
- Use an explicit reader: `SELECT * FROM read_parquet(?)`, rather than `SELECT * FROM 'file.parquet'`. This also makes the source visible in the query.
- Refer to computed tables through supplied handles, never a guessed internal table name. CTE definitions can use earlier CTEs; a recursive term can use its working table. Self-references cannot expose an unowned physical table.
- Use explicit pivot values, for example `PIVOT source ON category IN ('a','b') USING sum(amount)`. Automatic discovery of pivot columns generates an extra engine statement and is unavailable.
- Materialize a table handle when several downstream computations should share one source snapshot. A fragment is a recipe: a changed downstream query can execute its source reads again.
- Reentering an unchanged formula or pressing F9/Shift F9 reuses a retained result. To fetch changed external data, restart that workbook's XlDuck session. Opening a trusted workbook starts a fresh session.

Fragments pass the same capability validation when acquired, without reading their external source. Output and plot materialization then run through the same checked query path. There is no privileged preview bypass.

## Host connection settings

Settings contain routing information, never access keys or session tokens. For example, host code can construct:

```csharp
var settings = new ConnectionSettings(
    EnableExternalReads: true,
    ReaderExtensionDirectory: installedReaderDirectory,
    S3: [new S3Connection(
        Scope: "s3://research-bucket/",
        Credentials: S3CredentialSource.AwsProfile,
        Region: "eu-west-2",
        Profile: "research")]);
```

Settings are sent over the private worker startup pipe, not command-line arguments or cells, and are fixed for the session. A settings change requires a fresh session. Ordinary restart preserves the host's chosen settings.

| Credential source | Meaning |
| --- | --- |
| `Anonymous` | Public S3-compatible sources without credentials |
| `Environment` | AWS credentials supplied to the host process and inherited by its worker |
| `AwsProfile` | Explicit named AWS configuration/credentials profile |
| `AwsSso` | Explicit named SSO profile with authentication managed outside worksheet SQL; live SSO validation is still pending |

The host creates temporary scoped DuckDB secrets through the supported AWS provider, then clears legacy credential settings populated by the reader extension. There is no fallback to a general provider chain: each connection chooses `env`, `config` or `sso`. Named profile configuration can itself contain user-managed credential providers. S3 scope selects credentials; it is not a read permission. Include the slash after the bucket, and use a prefix only when routing requires it. Region, optional endpoint, TLS and path-style addressing support S3-compatible services. TLS defaults to enabled.

Authenticate or renew credentials outside the sheet, then restart the workbook session. Refresh and expiry behaviour for live SSO and provider-specific profiles still needs integration validation. Setup failures identify the connection without returning provider exception details.

The new engine does not load persistent DuckDB secrets and does not offer arbitrary `DuckConfig` SQL. Existing users will configure environment credentials or named profiles through the host settings instead. No compatibility bridge to the legacy secret store is planned.

## Pinned engine and extensions

The policy is pinned to native DuckDB **v1.5.6**, supplied by DuckDB.NET **1.5.6.1**. A different native version fails closed until its function catalog and policy are reviewed. This is separate from the still-open public .NET runtime and deployment decision.

Only the matching official Windows AMD64 `httpfs` and, when credentials require it, `aws` extension binaries are loaded by privileged setup. SHA-256 checks in [ConnectionSetup.cs](ConnectionSetup.cs) pin the exact binaries; DuckDB signature checks remain enabled. Automatic extension installation/loading, community extensions, unsigned extensions, persistent secrets and engine/HTTP logging are disabled. Configuration is then locked.

For development, [Get-ReaderExtensions.ps1](../scripts/Get-ReaderExtensions.ps1) downloads those two pinned binaries from DuckDB's official distribution and verifies them before placing them in the ignored artifacts directory. Worksheet execution never downloads extensions. Public installation must ship these verified dependencies and provide a repair path; that packaging work remains open.

The transport does not log SQL, query arguments or worker stderr. Native error presentation removes SQL excerpts, quoted values and remote URLs, and avoids suggesting forbidden extension installation. This is diagnostic hygiene, not a complete secret detector: a query can deliberately return sensitive data, and applications should not log returned errors indiscriminately.

## Validation and remaining work

The policy suite runs against both the direct engine and the workbook worker. It covers real local CSV/JSON/Parquet reads, result reuse after a source changes, nested prohibited functions, CTE ownership, actual absence of write effects, locked settings, fragments/plots and extension integrity. The Excel suite covers coarse host configuration, blocked SQL and preserving settings across restart.

Optional reader tests use an owned loopback HTTP/S3 fixture with dummy credentials. They exercise HTTP reads, S3 Parquet reads and wildcard listing, named profiles, request methods, errors and signed-URL redaction. The fixture checks authentication headers but does not cryptographically verify AWS signatures. These tests do not establish compatibility with a live AWS account, SSO login, enterprise proxy or all S3-compatible services.

Production lifecycle, trust persistence, settings and preview integration are implemented. Before public release, complete live source/SSO checks and packaged dependency validation. See [the modernization plan](../PLAN-modernization.md) and [test commands](../tests/README.md).

The implementation uses DuckDB's documented [SQL serialization](https://duckdb.org/docs/current/data/json/sql_to_and_from_json), [AWS credential providers](https://duckdb.org/docs/1.5/core_extensions/aws) and [S3 reader](https://duckdb.org/docs/1.5/core_extensions/httpfs/s3api). Engine settings are supplied through [DuckDB.NET connection configuration](https://duckdb.net/docs/connection-string.html); these mechanisms support, but do not replace, XlDuck's capability checks.
