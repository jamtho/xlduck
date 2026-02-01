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
- Test: Open XlDuck-AddIn64.xll directly (launches Excel with add-in)
- The 64-bit add-in includes native DuckDB; 32-bit does not

## DuckDB Notes

- Single in-memory connection is shared across all function calls
- Connection is lazily initialized on first use
- Tables/data persist for the Excel session
