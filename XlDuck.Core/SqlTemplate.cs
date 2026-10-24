// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Collections.Immutable;
using System.Text;

namespace XlDuck.Core;

/// <summary>Resolves positional arguments without interpreting question marks inside SQL literals/comments.</summary>
public static class SqlTemplate
{
    public static string QuoteIdentifier(string name) => "\"" + name.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";

    public static PreparedQuery Prepare(ComputationRequest request, Func<ResultHandle, ComputationRequest?> resolveFragment, int maxExpansionCharacters = 8 * 1024 * 1024)
    {
        if (maxExpansionCharacters < 1) throw new ArgumentOutOfRangeException(nameof(maxExpansionCharacters));
        var remaining = maxExpansionCharacters;
        void Consume(int characters)
        {
            if (characters > remaining) throw new ArgumentException("Expanded SQL exceeds this session's size limit. Simplify the fragments or materialize a shared intermediate result.");
            remaining -= characters;
        }
        var parameters = ImmutableArray.CreateBuilder<BoundParameter>();
        var tables = new HashSet<ResultHandle>();
        var expanding = new HashSet<ResultHandle>();
        string Expand(ComputationRequest recipe, int depth)
        {
            if (depth > 64) throw new ArgumentException("Fragment nesting exceeds 64 levels.");
            var index = 0;
            var sql = Rewrite(recipe.Sql, () =>
            {
                if (index >= recipe.Arguments.Length) throw new ArgumentException("SQL has more placeholders than arguments.");
                var argument = recipe.Arguments[index++];
                argument.Validate();
                if (argument.Value is { } value)
                {
                    var name = "__xl" + parameters.Count;
                    parameters.Add(new(name, value));
                    return value.Kind == ValueKind.Decimal ? $"CAST(${name} AS DECIMAL(38,{value.DecimalScale()}))" : "$" + name;
                }
                var handle = argument.Handle!.Value;
                if (handle.Kind == ResultKind.Table) { tables.Add(handle); return QuoteIdentifier(handle.TableName); }
                if (handle.Kind != ResultKind.Fragment) throw new ArgumentException("SQL expects a table or fragment handle, not a plot.");
                if (!expanding.Add(handle)) throw new ArgumentException("Fragment dependency cycle.");
                try
                {
                    var fragment = resolveFragment(handle) ?? throw new ArgumentException("Fragment is unavailable in this session.");
                    return "(" + Expand(fragment, depth + 1) + "\n)";
                }
                finally { expanding.Remove(handle); }
            }, Consume);
            if (index != recipe.Arguments.Length) throw new ArgumentException("More arguments were supplied than SQL placeholders.");
            return sql;
        }
        return new(Expand(request, 0), parameters.ToImmutable(), [.. tables.OrderBy(handle => handle.Id)]);
    }

    private static string Rewrite(string sql, Func<string> placeholder, Action<int> consume)
    {
        var output = new StringBuilder();
        void AppendCharacter(char value) { consume(1); output.Append(value); }
        void AppendSlice(int start, int length) { consume(length); output.Append(sql, start, length); }
        void AppendText(string value) { consume(value.Length); output.Append(value); }
        bool terminated = false;
        for (var i = 0; i < sql.Length;)
        {
            var start = i;
            if (char.IsWhiteSpace(sql[i])) { AppendCharacter(sql[i++]); continue; }
            if (sql[i] == '-' && i + 1 < sql.Length && sql[i + 1] == '-')
            {
                i += 2;
                while (i < sql.Length && sql[i] is not ('\n' or '\r')) i++;
                AppendSlice(start, i - start); continue;
            }
            if (sql[i] == '/' && i + 1 < sql.Length && sql[i + 1] == '*')
            {
                i += 2;
                var nesting = 1;
                while (i < sql.Length && nesting > 0)
                {
                    if (i + 1 < sql.Length && sql[i] == '/' && sql[i + 1] == '*') { nesting++; i += 2; }
                    else if (i + 1 < sql.Length && sql[i] == '*' && sql[i + 1] == '/') { nesting--; i += 2; }
                    else i++;
                }
                if (nesting != 0) throw new ArgumentException("Unterminated SQL comment.");
                AppendSlice(start, i - start); continue;
            }
            if (terminated) throw new ArgumentException("Only one SQL statement is allowed.");
            if (sql[i] == ';') { terminated = true; i++; continue; }
            if (sql[i] is '\'' or '"')
            {
                var quote = sql[i++];
                var escaped = quote == '\'' && start > 0 && sql[start - 1] is 'E' or 'e' &&
                    (start < 2 || !IsIdentifier(sql[start - 2]));
                var closed = false;
                while (i < sql.Length)
                {
                    if (escaped && sql[i] == '\\') { i = Math.Min(i + 2, sql.Length); continue; }
                    if (sql[i++] != quote) continue;
                    if (i < sql.Length && sql[i] == quote) { i++; continue; }
                    closed = true; break;
                }
                if (!closed) throw new ArgumentException("Unterminated SQL literal or identifier.");
                AppendSlice(start, i - start); continue;
            }
            if (sql[i] == '$')
            {
                var endTag = i + 1;
                while (endTag < sql.Length && (char.IsAsciiLetterOrDigit(sql[endTag]) || sql[endTag] == '_')) endTag++;
                if (endTag < sql.Length && sql[endTag] == '$' && (endTag == i + 1 || !char.IsAsciiDigit(sql[i + 1])))
                {
                    var tag = sql[i..(endTag + 1)];
                    var end = sql.IndexOf(tag, endTag + 1, StringComparison.Ordinal);
                    if (end < 0) throw new ArgumentException("Unterminated dollar-quoted SQL literal.");
                    i = end + tag.Length; AppendSlice(start, i - start); continue;
                }
                throw new ArgumentException("Use ? placeholders for XlDuck arguments; named SQL parameters are reserved.");
            }
            if (sql[i] == '?') { AppendText(placeholder()); i++; continue; }
            AppendCharacter(sql[i++]);
        }
        // A trailing line comment must not consume the wrapping SQL added by a caller.
        AppendCharacter('\n');
        return output.ToString();
    }
    private static bool IsIdentifier(char value) => char.IsLetterOrDigit(value) || value == '_';
}
