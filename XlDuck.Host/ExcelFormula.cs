// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
namespace XlDuck.Host;

public static class ExcelFormula
{
    private static readonly HashSet<string> Functions = new(StringComparer.OrdinalIgnoreCase)
    {
        "DuckQuery", "DuckQueryOut", "DuckQueryOutScalar", "DuckFrag", "DuckCapture", "DuckOut", "DuckScalar",
        "DuckPlot", "DuckInspect", "DuckDate", "DuckDateTime", "DuckStatus"
    };
    public static bool UsesXlDuck(string formula)
    {
        for (var i = 0; i < formula.Length; i++)
        {
            if (formula[i] is '"' or '\'')
            {
                var quote = formula[i++];
                for (; i < formula.Length; i++)
                    if (formula[i] == quote) { if (i + 1 < formula.Length && formula[i + 1] == quote) i++; else break; }
            }
            else if (char.IsLetter(formula[i]) || formula[i] == '_')
            {
                var start = i;
                while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] is '_' or '.')) i++;
                var name = formula[start..i]; name = name[(name.LastIndexOf('.') + 1)..];
                while (i < formula.Length && char.IsWhiteSpace(formula[i])) i++;
                if (i < formula.Length && formula[i] == '(' && Functions.Contains(name)) return true;
                i--;
            }
        }
        return false;
    }
}
