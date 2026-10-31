// Copyright (c) 2026 James Thompson
// SPDX-License-Identifier: MPL-2.0
using System.Globalization;
using XlDuck.Core;

namespace XlDuck.Host;

public static class ExcelValues
{
    public static object? Display(CellValue value, bool date1904)
    {
        var display = value.ToDisplayValue();
        if (value.Kind is ValueKind.Date or ValueKind.Timestamp && display is double serial)
        {
            if (date1904) serial -= 1462;
            else if (serial < 61) serial -= 1;
            return serial < 0 ? value.Text : serial;
        }
        return display;
    }
    public static string DateText(double serial, bool date1904, bool includeTime)
    {
        if (!double.IsFinite(serial) || serial < 0 || !date1904 && serial >= 60 && serial < 61)
            throw new ArgumentException("This Excel serial does not represent a real date.");
        var oa = date1904 ? serial + 1462 : serial < 60 ? serial + 1 : serial;
        return DateTime.FromOADate(oa).ToString(includeTime ? "yyyy-MM-dd HH:mm:ss.fffffff" : "yyyy-MM-dd", CultureInfo.InvariantCulture);
    }
}
