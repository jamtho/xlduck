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
└── bin/Debug/net8.0-windows/
    ├── XlDuck-AddIn64.xll # Development add-in (use this)
    └── publish/
        └── XlDuck-AddIn64-packed.xll  # Standalone distributable
```

## Development Workflow

- Build: `dotnet build` from XlDuck directory
- Test: Run `tests/Run-Tests.ps1` (requires Excel with add-in loaded)
- The 64-bit add-in includes native DuckDB; 32-bit does not

## Testing

### Automated Tests

Tests use PowerShell COM automation to interact with Excel. Run:

```powershell
# First, launch Excel with the add-in
Start-Process "XlDuck\bin\Debug\net8.0-windows\XlDuck-AddIn64.xll"

# Then run tests (after dismissing any security dialogs)
powershell -ExecutionPolicy Bypass -File tests/Run-Tests.ps1
```

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
