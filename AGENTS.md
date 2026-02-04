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

### TDD Cycle

After making code changes, use the reload script to rebuild and restart Excel with the add-in:

```powershell
# Closes Excel, rebuilds, relaunches with add-in loaded
powershell -ExecutionPolicy Bypass -File scripts/reload-addin.ps1

# Run the test suite
powershell -ExecutionPolicy Bypass -File tests/Run-Tests.ps1
```

The reload script:
1. Cleanly closes Excel via COM (no save prompts)
2. Runs `dotnet build`
3. Launches Excel with a blank workbook
4. Registers the XLL add-in

Use `-NoBuild` to skip the build step if you only need to restart Excel.

## Testing

### Automated Tests

Tests use PowerShell COM automation to interact with Excel. The test suite requires Excel to be running with the add-in loaded and a workbook open - use `scripts/reload-addin.ps1` to set this up.

### Regression Tests

When fixing a bug, add a test that reproduces the bug unless you have a very good reason not to. This prevents the same bug from returning later. The test should fail before the fix and pass after.

### Key Testing Notes

1. **Use `.Formula2` not `.Formula`** when writing formulas via COM. `.Formula` adds the `@` implicit intersection operator which prevents dynamic array spilling.

2. **xl.exe (xl-cli) may not find Excel via ROT** - use direct PowerShell COM instead:
   ```powershell
   $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
   $sheet = $excel.ActiveWorkbook.ActiveSheet
   ```

3. **Security dialogs block automation** - dismiss the unsigned add-in warning manually before running tests.

4. **Excel must have a workbook open** - create one via COM if needed:
   ```powershell
   $excel.Workbooks.Add() | Out-Null
   ```

## DuckDB Notes

- Single in-memory connection is shared across all function calls
- Connection is lazily initialized on first use
- Tables/data persist for the Excel session
