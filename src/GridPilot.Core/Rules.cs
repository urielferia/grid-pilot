using System.Globalization;
using System.Text.RegularExpressions;

namespace GridPilot.Core;

public static class Rules
{
    public static decimal Number(string value) => decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) ? n : throw new InvalidOperationException("A valid number was expected.");
    public static void Formula(string formula)
    {
        if (!formula.StartsWith('=') || formula.Length > 1000 || Regex.IsMatch(formula, @"[\[\]\\|{}#;]|https?:|_xlfn|_xlpm", RegexOptions.IgnoreCase))
            throw new InvalidOperationException("Unsupported formula; use internal A1 references and allowed functions.");
        var stripped = Regex.Replace(formula, "\"[^\"]*\"|'[^']*'!", "");
        foreach (Match m in Regex.Matches(stripped, @"([A-Za-z_][A-Za-z0-9_.]*)\s*\("))
            if (!new[] { "SUM", "SUMIFS", "COUNTIFS", "IF", "ROUND", "INDEX", "MATCH" }.Contains(m.Groups[1].Value.ToUpperInvariant()))
                throw new InvalidOperationException("Unsupported function: " + m.Groups[1].Value);
        if (stripped.Contains(':') && Regex.IsMatch(stripped, @"\b[A-Z]+:[A-Z]+\b", RegexOptions.IgnoreCase))
            throw new InvalidOperationException("Use bounded cell ranges.");
        var tokens = Regex.Replace(stripped, @"(?:[A-Za-z_][A-Za-z0-9_]*!)|(?:\$?[A-Za-z]{1,3}\$?[0-9]+)|(?:\b(?:SUM|SUMIFS|COUNTIFS|IF|ROUND|INDEX|MATCH)\s*(?=\())", "", RegexOptions.IgnoreCase);
        if (Regex.IsMatch(tokens, @"[A-Za-z_!@]"))
            throw new InvalidOperationException("Unsupported reference or name; use internal A1 cells.");
    }
    public static void Validate(WorkbookSnapshot s, WorkbookProfile p)
    {
        foreach (var t in p.Tables)
        {
            var table = s.Tables.SingleOrDefault(x => x.Name == t.Name) ?? throw new InvalidOperationException("Missing table " + t.Name);
            foreach (var f in t.Fields) if (!table.Headers.Contains(f.Header)) throw new InvalidOperationException("Missing field " + f.Header);
            var tableStart = Regex.Match(table.Address, @"([A-Z]+)(\d+)");
            int startColumn = 0; foreach (char ch in tableStart.Groups[1].Value) startColumn = startColumn * 26 + ch - 'A' + 1;
            var startRow = int.Parse(tableStart.Groups[2].Value, CultureInfo.InvariantCulture);
            foreach (var f in t.Fields.Where(f => f.Calculated)) for (int r = 0; r < table.Rows.Count; r++)
            {
                var column = startColumn + table.Headers.IndexOf(f.Header); var letters = ""; while (column > 0) { column--; letters = (char)('A' + column % 26) + letters; column /= 26; }
                var cell = s.Cells.SingleOrDefault(c => c.Sheet == table.Sheet && c.Address == letters + (startRow + r + 1));
                if (cell is null || cell.Formula == "") throw new InvalidOperationException("Missing calculated formula: " + f.Header);
            }
            var id = t.Fields.Single(f => f.Key == "Id").Header;
            if (table.Rows.Select(r => r[id]).Distinct(StringComparer.OrdinalIgnoreCase).Count() != table.Rows.Count)
                throw new InvalidOperationException("Duplicate identifiers.");
            foreach (var r in table.Rows) foreach (var f in t.Fields)
            {
                var v = r[f.Header];
                if (f.Required && string.IsNullOrWhiteSpace(v)) throw new InvalidOperationException("Missing " + f.Header);
                if (f.Type is "number" or "date") { var n = Number(v); if (f.Type == "date" && (n < 1 || n > 2958465 || n != decimal.Truncate(n))) throw new InvalidOperationException("Invalid date."); }
            }
        }
        List<Dictionary<string, string>> Rows(string name) => s.Table(name).Rows.Select(r => p.Table(name).Fields.ToDictionary(f => f.Key, f => r[f.Header])).ToList();
        var materials = Rows("Inventario"); var visits = Rows("Visitas"); var movements = Rows("Movimientos");
        foreach (var v in visits) if (v["Estado"] is not ("Realizada" or "Programada")) throw new InvalidOperationException("Invalid visit status.");
        foreach (var m in movements)
        {
            if (m["Tipo"] is not ("Entrada" or "Consumo")) throw new InvalidOperationException("Invalid movement type.");
            if (!materials.Any(x => x["Id"] == m["Material"])) throw new InvalidOperationException("Non-existent material.");
            if (Number(m["Cantidad"]) <= 0 || Number(m["CostoUnitario"]) < 0) throw new InvalidOperationException("Invalid quantity or unit cost.");
            if (m["Tipo"] == "Consumo" && !visits.Any(x => x["Id"] == m["Visita"] && x["Estado"] == "Realizada")) throw new InvalidOperationException("Consumption requires a completed visit.");
            if (m["Visita"] != "" && !visits.Any(x => x["Id"] == m["Visita"])) throw new InvalidOperationException("Non-existent visit.");
            Equal(Number(m["Cantidad"]) * Number(m["CostoUnitario"]), Number(m["Importe"]), .01m);
        }
        foreach (var m in materials)
        {
            if (new[] { "SaldoInicial", "Minimo", "CostoReferencia" }.Any(f => Number(m[f]) < 0)) throw new InvalidOperationException("Catalog contains negative values.");
            var related = movements.Where(x => x["Material"] == m["Id"]).ToList();
            var incoming = related.Where(x => x["Tipo"] == "Entrada").Sum(x => Number(x["Cantidad"]));
            var outgoing = related.Where(x => x["Tipo"] == "Consumo").Sum(x => Number(x["Cantidad"]));
            var stock = Number(m["SaldoInicial"]) + incoming - outgoing;
            if (stock < 0) throw new InvalidOperationException("Insufficient stock: " + m["Id"]);
            Equal(incoming, Number(m["Entradas"]), .000001m); Equal(outgoing, Number(m["Consumos"]), .000001m); Equal(stock, Number(m["Existencia"]), .000001m);
        }
        foreach (var v in visits) Equal(movements.Where(x => x["Tipo"] == "Consumo" && x["Visita"] == v["Id"]).Sum(x => Number(x["Cantidad"]) * Number(x["CostoUnitario"])), Number(v["CostoMateriales"]), .01m);
        foreach (var c in s.Cells)
        {
            if (c.Value.StartsWith('#') || c.Formula.Contains("#REF!")) throw new InvalidOperationException("Excel error in " + c.Sheet + "!" + c.Address);
            if (c.Formula == "") continue;
            var clean = Regex.Replace(c.Formula, "\"[^\"]*\"|'[^']*'!", "");
            foreach (Match function in Regex.Matches(clean, @"([A-Za-z_][A-Za-z0-9_.]*)\s*\("))
                if (!new[] { "SUM", "SUMIFS", "COUNTIFS", "IF", "ROUND", "INDEX", "MATCH", "SUMPRODUCT" }.Contains(function.Groups[1].Value.ToUpperInvariant()))
                    throw new InvalidOperationException("Unsupported function in " + c.Sheet + "!" + c.Address);
        }
        string Cell(string address) => s.Cells.Single(c => c.Sheet == p.SummarySheet && c.Address == address).Value;
        var start = Number(Cell("B5")); var end = Number(Cell("B6"));
        if (start > end) throw new InvalidOperationException("Invalid summary period.");
        Equal(visits.Count(v => v["Estado"] == "Realizada" && Number(v["Fecha"]) >= start && Number(v["Fecha"]) <= end), Number(Cell("B9")), 0);
        Equal(movements.Where(m => m["Tipo"] == "Consumo" && Number(m["Fecha"]) >= start && Number(m["Fecha"]) <= end).Sum(m => Number(m["Cantidad"]) * Number(m["CostoUnitario"])), Number(Cell("B10")), .01m);
        Equal(materials.Count(m => Number(m["Existencia"]) < Number(m["Minimo"])), Number(Cell("B11")), 0);
    }
    static void Equal(decimal expected, decimal actual, decimal tolerance) { if (Math.Abs(expected - actual) > tolerance) throw new InvalidOperationException("Excel calculation does not match the independent verifier."); }
}
