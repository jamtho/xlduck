# Agent Guidelines

## Commit Messages

- Write commit messages as a human would - concise, descriptive, imperative mood
- Do not add "Created with Claude", "Co-Authored-By: Claude", or similar attribution
- Do not add emoji to commit messages
- Follow conventional commits style when appropriate (feat:, fix:, refactor:, etc.)

## Code Style

- Target .NET 8+ with C# 12 features (file-scoped namespaces, nullable reference types)
- Keep Excel UDFs simple and focused - one function, one purpose
- Wrap all DuckDB calls in try/catch, returning `#ERROR: message` for Excel display
- Use static classes for Excel function containers

## Project Structure

```
XlDuck/
├── XlDuck.csproj          # Project file with Excel-DNA and DuckDB refs
├── DuckFunctions.cs       # Excel UDFs
├── RibbonController.cs    # XlDuck ribbon tab callbacks
├── Ribbon.xml             # Ribbon UI definition
├── Log.cs                 # File logger to %LOCALAPPDATA%\XlDuck
├── Preview/
│   ├── PreviewPane.cs           # WebView2 host (COM-visible UserControl)
│   ├── PreviewPaneManager.cs    # Singleton managing panes per window
│   ├── PreviewController.cs     # Debounce and serial queue
│   ├── PreviewDataProvider.cs   # Data access for preview
│   ├── PreviewModels.cs         # JSON models for WebView2
│   └── preview.html             # Embedded HTML UI
└── bin/Debug/net8.0-windows/
    ├── XlDuck-AddIn64.xll # Development add-in (use this)
    └── publish/
        └── XlDuck-AddIn64-packed.xll  # Standalone distributable
```

## Development Workflow

The 64-bit add-in includes native DuckDB; 32-bit does not.

### Development and integration testing

After code changes, build and run the isolated integration suite:

```powershell
powershell -ExecutionPolicy Bypass -File tests/Run-Tests.ps1
```

The suite builds the add-in, creates a new Excel process and temporary workbook, verifies process ownership, and closes only that workbook and instance. Use `-NoBuild` to test an existing build. Never attach tests to `ActiveWorkbook` or use the Running Object Table to choose an Excel instance.

For a visible development session:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Reload-AddIn.ps1
```

Despite its historical name, this script opens a new development instance and leaves existing Excel instances alone. It loads a private copy of the build so subsequent builds can proceed. Close development instances manually. Copies are under the ignored `artifacts/development` directory; remove them only after their Excel instances are closed.

## Testing

### Regression Tests

When fixing a bug, add a test that reproduces the bug unless you have a very good reason not to. This prevents the same bug from returning later. The test should fail before the fix and pass after.

### Key Testing Notes

1. **Use `.Formula2` not `.Formula`** when writing formulas via COM. `.Formula` adds the `@` implicit intersection operator which prevents dynamic array spilling.

2. Use `scripts/ExcelSession.ps1` for automation ownership and bounded cell waits. Keep `.Value2` assertions independent of cell width and display formatting.

3. Security dialogs can block add-in registration. Do not change Trust Center settings or dismiss unrelated prompts automatically.

4. Never close unrelated workbooks, suppress their save prompts, or kill all Excel processes. A failed test must release its own resources in `finally`.

## DuckDB Notes

- Single in-memory connection is shared across all function calls
- Connection is lazily initialized on first use
- Tables/data persist for the Excel session
