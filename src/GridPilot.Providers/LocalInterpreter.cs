using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using GridPilot.Core;

namespace GridPilot.Providers;

/// <summary>
/// Resolves clear English and Spanish natural-language instructions against the current workbook snapshot
/// without requiring a cloud round-trip (or when the user operates locally), including row-based
/// operations, replacing/copying/swapping rows, cascading row deletions, bulk updates, and structural edits.
/// </summary>
public static class LocalInterpreter
{
    public static Interpretation? TryInterpret(string instruction, WorkbookSnapshot snapshot, WorkbookProfile profile)
    {
        if (string.IsNullOrWhiteSpace(instruction)) return null;
        var clauses = SplitClauses(instruction);
        var operations = new List<Operation>();
        string? lastTable = null;
        foreach (var clause in clauses)
        {
            var sub = TryClause(clause, snapshot, profile, ref lastTable);
            if (sub is null) return null;
            if (!string.IsNullOrEmpty(sub.Clarification)) return sub;
            operations.AddRange(sub.Operations);
        }
        return operations.Count > 0 ? new("", operations) : null;
    }

    static List<string> SplitClauses(string text)
    {
        var lines = text.Split(['\r', '\n', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var list = new List<string>();
        foreach (var line in lines)
        {
            var parts = Regex.Split(
                line,
                @"\s+(?:y\s+luego|luego|despu[eé]s|adem[aá]s|tambien|tambi[eé]n|and\s+then|then|afterwards|also|y\s+(?=(?:en\s+la|cambia|actualiza|modifica|establece|pon|fija|reemplaza|copia|intercambia|elimina|borra|quita|agrega|a[ñn]ade|crea|renombra))|and\s+(?=(?:in\s+the|on\s+the|change|update|modify|set|replace|copy|swap|delete|remove|add|create|rename|increase|increment|decrease|subtract)))\s+",
                RegexOptions.IgnoreCase);
            foreach (var p in parts) if (!string.IsNullOrWhiteSpace(p)) list.Add(p.Trim().TrimEnd('.'));
        }
        return list;
    }

    static Interpretation? TryClause(string text, WorkbookSnapshot snapshot, WorkbookProfile profile, ref string? contextTable)
    {
        var norm = Norm(text);
        var detected = DetectTable(norm, profile);
        if (detected is not null) contextTable = detected;
        var tableName = detected ?? contextTable;

        // 1. Structural: AddSheet / RenameSheet / DeleteSheet (Spanish & English)
        if (Regex.Match(text, @"(?:crea|agrega|a[ñn]ade|create|add)\s+(?:una\s+|a\s+)?(?:hoja|pesta[ñn]a|sheet|tab)\s+(?:llamada\s+|con\s+nombre\s+|called\s+|named\s+)?[""'«]?([A-Za-z0-9ÁÉÍÓÚáéíóúÑñ_\- ]+?)[""'»]?$", RegexOptions.IgnoreCase) is { Success: true } addSh)
            return new("", [new(OperationKind.AddSheet, Value: addSh.Groups[1].Value.Trim())]);
        if (Regex.Match(text, @"(?:renombra|rename)\s+(?:la\s+|the\s+)?(?:hoja|pesta[ñn]a|sheet|tab)\s+[""'«]?([A-Za-z0-9ÁÉÍÓÚáéíóúÑñ_\- ]+?)[""'»]?\s+(?:a|como|por|to|as)\s+[""'«]?([A-Za-z0-9ÁÉÍÓÚáéíóúÑñ_\- ]+?)[""'»]?$", RegexOptions.IgnoreCase) is { Success: true } renSh)
            return new("", [new(OperationKind.RenameSheet, Table: ResolveSheetName(renSh.Groups[1].Value.Trim(), snapshot, profile), Value: renSh.Groups[2].Value.Trim())]);
        if (Regex.Match(text, @"(?:elimina|borra|quita|delete|remove)\s+(?:la\s+|the\s+)?(?:hoja|pesta[ñn]a|sheet|tab)\s+[""'«]?([A-Za-z0-9ÁÉÍÓÚáéíóúÑñ_\- ]+?)[""'»]?$", RegexOptions.IgnoreCase) is { Success: true } delSh)
            return new("", [new(OperationKind.DeleteSheet, Table: ResolveSheetName(delSh.Groups[1].Value.Trim(), snapshot, profile))]);

        // 2. Structural: AddColumn / RenameColumn / DeleteColumn (Spanish & English)
        if (tableName is not null)
        {
            var tp = profile.Table(tableName);
            if (Regex.Match(text, @"(?:agrega|a[ñn]ade|crea|add|create)\s+(?:la\s+|una\s+|the\s+|a\s+)?(?:columna|column)(?:\s+calculada|\s+de\s+texto|\s+num[eé]rica|\s+calculated)?\s+(?:llamada\s+|called\s+|named\s+)?[""'«]?([A-Za-z0-9ÁÉÍÓÚáéíóúÑñ_ ]+?)[""'»]?(?:\s+(?:a|en|to|in)\s+(?:la\s+|the\s+)?(?:tabla\s+|hoja\s+|pesta[ñn]a\s+|table\s+|sheet\s+|tab\s+)?[A-Za-zÁÉÍÓÚáéíóúÑñ]+)?(?:\s*[,:]?\s*(?:con\s+f[oó]rmula|f[oó]rmula|igual\s+a|con|with\s+formula|formula|equal\s+to|with)\s*(=[^\s.]+.*))?$", RegexOptions.IgnoreCase) is { Success: true } addCol)
            {
                var colName = Regex.Replace(addCol.Groups[1].Value.Trim(), @"\s+(?:a|en|al|to|in)\s+(?:inventario|inventory|visitas?|visits?|movimientos?|movements?)$", "", RegexOptions.IgnoreCase).Trim();
                var formulaMatch = Regex.Match(text, @"(=[A-Za-z0-9_($:+\-*/.,""!<>=& ]+)");
                var formula = formulaMatch.Success ? formulaMatch.Groups[1].Value.Trim().TrimEnd('.') : "";
                return new("", [new(OperationKind.AddColumn, tp.Name, Field: colName, Value: formula)]);
            }
            if (Regex.Match(text, @"(?:renombra|rename)\s+(?:la\s+|the\s+)?(?:columna|campo|column|field)\s+[""'«]?([A-Za-z0-9ÁÉÍÓÚáéíóúÑñ_ ]+?)[""'»]?(?:\s+(?:de|in|of)\s+[A-Za-zÁÉÍÓÚáéíóúÑñ]+)?\s+(?:a|como|por|to|as)\s+[""'«]?([A-Za-z0-9ÁÉÍÓÚáéíóúÑñ_ ]+?)[""'»]?$", RegexOptions.IgnoreCase) is { Success: true } renCol)
            {
                var f = FindField(tp, renCol.Groups[1].Value.Trim());
                if (f is not null) return new("", [new(OperationKind.RenameColumn, tp.Name, Field: f.Key, Value: renCol.Groups[2].Value.Trim())]);
            }
            if (Regex.Match(text, @"(?:elimina|borra|quita|delete|remove)\s+(?:la\s+|the\s+)?(?:columna|campo|column|field)\s+(?:de\s+|of\s+)?[""'«]?([A-Za-z0-9ÁÉÍÓÚáéíóúÑñ_ ]+?)[""'»]?(?:\s+(?:de|en|del|in|from|of)\s+[A-Za-zÁÉÍÓÚáéíóúÑñ]+)?$", RegexOptions.IgnoreCase) is { Success: true } delCol)
            {
                var f = FindField(tp, delCol.Groups[1].Value.Trim());
                if (f is not null) return new("", [new(OperationKind.DeleteColumn, tp.Name, Field: f.Key)]);
            }
        }

        // 3. Replace / Copy row data (Spanish & English)
        if (tableName is not null && (norm.Contains("reemplaza") || norm.Contains("sustituye") || norm.Contains("copia") || norm.Contains("iguala") || norm.Contains("replace") || norm.Contains("copy")))
        {
            var tp = profile.Table(tableName);
            var ts = snapshot.Table(tp.Name);
            var rows = ExtractRowReferences(norm, tp, ts);
            if (rows.Count == 2)
            {
                bool copyOrder = (norm.Contains("copia") && Regex.IsMatch(norm, @"copia.*\b(?:a|en|hacia)\b"))
                              || (norm.Contains("copy") && Regex.IsMatch(norm, @"copy.*\b(?:to|into)\b"));
                var targetIdx = copyOrder ? rows[1] : rows[0];
                var sourceIdx = copyOrder ? rows[0] : rows[1];
                if (targetIdx >= 1 && targetIdx <= ts.Rows.Count && sourceIdx >= 1 && sourceIdx <= ts.Rows.Count)
                {
                    var idHeader = tp.Fields.Single(f => f.Key == "Id").Header;
                    var targetId = ts.Rows[targetIdx - 1][idHeader];
                    var sourceRow = ts.Rows[sourceIdx - 1];
                    var ops = new List<Operation>();
                    foreach (var f in tp.Fields.Where(f => !f.Calculated && f.Key != "Id"))
                        ops.Add(new(OperationKind.SetValue, tp.Name, targetId, f.Key, FormatValueForWrite(f, sourceRow[f.Header])));
                    return new("", ops);
                }
            }
        }

        // 4. Swap two rows (Spanish & English)
        if (tableName is not null && (norm.Contains("intercambia") || norm.Contains("permuta") || norm.Contains("invierte") || norm.Contains("swap") || norm.Contains("exchange")))
        {
            var tp = profile.Table(tableName);
            var ts = snapshot.Table(tp.Name);
            var rows = ExtractRowReferences(norm, tp, ts);
            if (rows.Count == 2 && rows[0] >= 1 && rows[0] <= ts.Rows.Count && rows[1] >= 1 && rows[1] <= ts.Rows.Count)
            {
                var idHeader = tp.Fields.Single(f => f.Key == "Id").Header;
                var idA = ts.Rows[rows[0] - 1][idHeader];
                var idB = ts.Rows[rows[1] - 1][idHeader];
                var rowA = ts.Rows[rows[0] - 1];
                var rowB = ts.Rows[rows[1] - 1];
                var ops = new List<Operation>();
                foreach (var f in tp.Fields.Where(f => !f.Calculated && f.Key != "Id"))
                    ops.Add(new(OperationKind.SetValue, tp.Name, idA, f.Key, FormatValueForWrite(f, rowB[f.Header])));
                foreach (var f in tp.Fields.Where(f => !f.Calculated && f.Key != "Id"))
                    ops.Add(new(OperationKind.SetValue, tp.Name, idB, f.Key, FormatValueForWrite(f, rowA[f.Header])));
                return new("", ops);
            }
        }

        // 5. Delete row(s) or record(s) (Spanish & English)
        if (Regex.IsMatch(norm, @"\b(?:elimina|borra|quita|suprime|delete|remove)\b") && !Regex.IsMatch(norm, @"\b(?:columna|campo|column|field|hoja|pesta[ñn]a\s+[a-z0-9_]+$|sheet\s+[a-z0-9_]+$|tab\s+[a-z0-9_]+$)\b"))
        {
            var explicitIds = ExtractExplicitIds(text);
            bool keepMovements = norm.Contains("pero deja") || norm.Contains("but leave") || norm.Contains("but keep");
            if (explicitIds.Count > 0 && !keepMovements)
            {
                var ops = new List<Operation>();
                // Delete movements first so foreign-key dependencies are satisfied when deleting a material/visit and its movements together.
                foreach (var (tbl, id) in explicitIds.OrderBy(x => x.Table == "Movimientos" ? 0 : 1))
                {
                    if (tbl is "Visitas" or "Inventario" && !explicitIds.Any(x => x.Table == "Movimientos"))
                        ops.AddRange(DependentMovementDeletes(snapshot, profile, tbl, id));
                    ops.Add(new(OperationKind.DeleteRecord, tbl, Id: id));
                }
                return new("", ops);
            }
            if (explicitIds.Count == 1 && keepMovements)
                return new("", [new(OperationKind.DeleteRecord, explicitIds[0].Table, Id: explicitIds[0].Id)]);

            if (tableName is not null)
            {
                var tp = profile.Table(tableName);
                var ts = snapshot.Table(tp.Name);
                var rowIndices = ExtractRowRangeOrList(norm, tp, ts);
                if (rowIndices.Count > 0)
                {
                    var idHeader = tp.Fields.Single(f => f.Key == "Id").Header;
                    var targetIds = rowIndices.Where(i => i >= 1 && i <= ts.Rows.Count).Select(i => ts.Rows[i - 1][idHeader]).Distinct().ToList();
                    if (targetIds.Count > 0)
                    {
                        var ops = new List<Operation>();
                        if (!keepMovements && tp.Name is "Visitas" or "Inventario")
                            foreach (var id in targetIds)
                                foreach (var dep in DependentMovementDeletes(snapshot, profile, tp.Name, id))
                                    if (!ops.Any(o => o.Kind == OperationKind.DeleteRecord && o.Table == dep.Table && o.Id == dep.Id))
                                        ops.Add(dep);
                        foreach (var id in targetIds)
                            ops.Add(new(OperationKind.DeleteRecord, tp.Name, Id: id));
                        return new("", ops);
                    }
                }
            }
        }

        // 6. Set field value(s) on a row or record (single or multi-field, or relative math)
        if (TrySetValues(text, norm, snapshot, profile, tableName, out var setOps))
            return new("", setOps);

        return null;
    }

    static List<Operation> DependentMovementDeletes(WorkbookSnapshot snapshot, WorkbookProfile profile, string parentTable, string parentId)
    {
        var movProfile = profile.Table("Movimientos");
        var movIdHeader = movProfile.Fields.Single(f => f.Key == "Id").Header;
        var fkHeader = movProfile.Fields.Single(f => f.Key == (parentTable == "Visitas" ? "Visita" : "Material")).Header;
        return snapshot.Table("Movimientos").Rows
            .Where(r => string.Equals(r[fkHeader], parentId, StringComparison.OrdinalIgnoreCase))
            .Select(r => new Operation(OperationKind.DeleteRecord, "Movimientos", Id: r[movIdHeader]))
            .ToList();
    }

    static bool TrySetValues(string text, string norm, WorkbookSnapshot snapshot, WorkbookProfile profile, string? tableName, out List<Operation> ops)
    {
        ops = [];
        var explicitIds = ExtractExplicitIds(text);
        TableProfile? tp = null;
        string? recordId = null;

        if (explicitIds.Count == 1 && !Regex.IsMatch(norm, @"\b(?:reemplaza|intercambia|copia|replace|swap|copy)\b"))
        {
            tp = profile.Table(explicitIds[0].Table);
            recordId = explicitIds[0].Id;
        }
        else if (tableName is not null)
        {
            tp = profile.Table(tableName);
            var ts = snapshot.Table(tp.Name);
            var rows = ExtractRowReferences(norm, tp, ts);
            if (rows.Count == 1 && rows[0] >= 1 && rows[0] <= ts.Rows.Count)
            {
                var idHeader = tp.Fields.Single(f => f.Key == "Id").Header;
                recordId = ts.Rows[rows[0] - 1][idHeader];
            }
        }
        if (tp is null || recordId is null) return false;

        var tsTarget = snapshot.Table(tp.Name);
        var idCol = tp.Fields.Single(f => f.Key == "Id").Header;
        var currentRow = tsTarget.Rows.SingleOrDefault(r => string.Equals(r[idCol], recordId, StringComparison.OrdinalIgnoreCase));
        if (currentRow is null) return false;

        // Relative arithmetic: Spanish & English ("incrementa/aumenta/suma/increase/increment/add" or "reduce/resta/disminuye/decrease/subtract")
        if (Regex.Match(norm, @"\b(incrementa|aumenta|suma|sube|decrementa|disminuye|reduce|resta|baja|increase|increment|add|decrease|decrement|subtract|lower)\b") is { Success: true } mathVerb)
        {
            var field = MatchFieldInText(norm, tp);
            var numMatch = Regex.Matches(norm, @"(?<![a-z0-9])(\d+(?:\.\d+)?)(?![a-z0-9])")
                .Select(m => m.Groups[1].Value)
                .LastOrDefault(n => !norm.Contains("fila " + n) && !norm.Contains("registro " + n) && !norm.Contains("row " + n) && !norm.Contains("record " + n));
            if (field is not null && !field.Calculated && field.Type == "number" && numMatch is not null && decimal.TryParse(numMatch, NumberStyles.Float, CultureInfo.InvariantCulture, out var delta))
            {
                var current = Rules.Number(currentRow[field.Header]);
                bool subtract = mathVerb.Groups[1].Value is "decrementa" or "disminuye" or "reduce" or "resta" or "baja" or "decrease" or "decrement" or "subtract" or "lower";
                var next = subtract ? current - delta : current + delta;
                ops.Add(new(OperationKind.SetValue, tp.Name, recordId, field.Key, next.ToString(CultureInfo.InvariantCulture)));
                return true;
            }
        }

        // Literal text assignment: "texto literal: ..." / "literal text: ..."
        if (Regex.Match(text, @"(?:texto\s+literal|literal\s+text)\s*:\s*(.+)$", RegexOptions.IgnoreCase) is { Success: true } lit)
        {
            var field = MatchFieldInText(norm, tp);
            if (field is not null && !field.Calculated && field.Key != "Id")
            {
                ops.Add(new(OperationKind.SetValue, tp.Name, recordId, field.Key, lit.Groups[1].Value.Trim()));
                return true;
            }
        }

        // Status shorthand (Spanish & English)
        if (tp.Name == "Visitas" && Regex.Match(norm, @"\b(realizada|programada|completed|scheduled)\b") is { Success: true } stMatch && (norm.Contains("estado") || norm.Contains("esta ") || norm.Contains("marca") || norm.Contains("pon") || norm.Contains("cambia") || norm.Contains("actualiza") || norm.Contains("status") || norm.Contains("mark") || norm.Contains("set") || norm.Contains("change") || norm.Contains("update") || norm.Contains("is ")))
        {
            if (!norm.Contains("consumo") && !norm.Contains("movimiento") && !norm.Contains("consumption") && !norm.Contains("movement"))
            {
                var val = stMatch.Groups[1].Value is "realizada" or "completed" ? "Realizada" : "Programada";
                ops.Add(new(OperationKind.SetValue, tp.Name, recordId, "Estado", val));
                return true;
            }
        }

        // General field change: "cambia el <campo> ... de X a Y" / "change <field> ... from X to Y" or "... to Y"
        var targetField = MatchFieldInText(norm, tp);
        if (targetField is not null && !targetField.Calculated && targetField.Key != "Id")
        {
            if (targetField.Type == "date")
            {
                if (TryExtractDate(text, out var isoDate))
                {
                    ops.Add(new(OperationKind.SetValue, tp.Name, recordId, targetField.Key, isoDate));
                    return true;
                }
                return false;
            }
            if (targetField.Type == "number")
            {
                var mFromTo = Regex.Match(norm, @"\b(?:de|from)\s+(-?\d+(?:\.\d+)?)\s*(?:mxn|pesos|piezas?|pieces?|units?)?\s+(?:a|to)\s+(-?\d+(?:\.\d+)?)\b");
                if (mFromTo.Success)
                {
                    ops.Add(new(OperationKind.SetValue, tp.Name, recordId, targetField.Key, mFromTo.Groups[2].Value));
                    return true;
                }
                var mTo = Regex.Matches(norm, @"\b(?:a|en|por|to|as|=)\s*(-?\d+(?:\.\d+)?)\s*(?:mxn|pesos|piezas?|pieces?|units?)?(?!\s*(?:fila|registro|row|record))\b");
                if (mTo.Count > 0)
                {
                    ops.Add(new(OperationKind.SetValue, tp.Name, recordId, targetField.Key, mTo[^1].Groups[1].Value));
                    return true;
                }
            }
            else
            {
                var mText = Regex.Match(text, @"\b(?:a|por|como|to|as)\s+[""'«]?([^""'».;]+?)[""'»]?\.?$", RegexOptions.IgnoreCase);
                if (mText.Success)
                {
                    var val = mText.Groups[1].Value.Trim();
                    if (targetField.Key == "Estado" && (val.Equals("realizada", StringComparison.OrdinalIgnoreCase) || val.Equals("completed", StringComparison.OrdinalIgnoreCase))) val = "Realizada";
                    if (targetField.Key == "Estado" && (val.Equals("programada", StringComparison.OrdinalIgnoreCase) || val.Equals("scheduled", StringComparison.OrdinalIgnoreCase))) val = "Programada";
                    ops.Add(new(OperationKind.SetValue, tp.Name, recordId, targetField.Key, val));
                    return true;
                }
            }
        }
        return false;
    }

    static FieldProfile? MatchFieldInText(string norm, TableProfile tp)
    {
        foreach (var f in tp.Fields.OrderByDescending(f => Math.Max(f.Header.Length, f.Key.Length)))
        {
            var h = Norm(f.Header);
            var k = Norm(f.Key);
            if (Regex.IsMatch(norm, $@"\b{Regex.Escape(h)}\b") || Regex.IsMatch(norm, $@"\b{Regex.Escape(k)}\b"))
                return f;
        }
        // English and Spanish field synonyms
        if (tp.Name == "Visitas")
        {
            if (norm.Contains("descripcion") || norm.Contains("description") || norm.Contains("trabajo") || norm.Contains("work")) return tp.Fields.SingleOrDefault(f => f.Key == "Descripcion");
            if (norm.Contains("cliente") || norm.Contains("client") || norm.Contains("customer")) return tp.Fields.SingleOrDefault(f => f.Key == "Cliente");
            if (norm.Contains("sitio") || norm.Contains("site") || norm.Contains("location")) return tp.Fields.SingleOrDefault(f => f.Key == "Sitio");
            if (norm.Contains("tecnico") || norm.Contains("technician")) return tp.Fields.SingleOrDefault(f => f.Key == "Tecnico");
            if (norm.Contains("estado") || norm.Contains("status")) return tp.Fields.SingleOrDefault(f => f.Key == "Estado");
            if (norm.Contains("fecha") || norm.Contains("date")) return tp.Fields.SingleOrDefault(f => f.Key == "Fecha");
        }
        if (tp.Name == "Inventario")
        {
            if (norm.Contains("saldo inicial") || norm.Contains("initial balance") || norm.Contains("opening balance") || norm.Contains("initial stock")) return tp.Fields.SingleOrDefault(f => f.Key == "SaldoInicial");
            if (norm.Contains("minimo") || norm.Contains("minimum")) return tp.Fields.SingleOrDefault(f => f.Key == "Minimo");
            if (norm.Contains("costo de referencia") || norm.Contains("costo referencia") || norm.Contains("reference cost") || norm.Contains("price") || norm.Contains("precio")) return tp.Fields.SingleOrDefault(f => f.Key == "CostoReferencia");
            if (norm.Contains("descripcion") || norm.Contains("description")) return tp.Fields.SingleOrDefault(f => f.Key == "Descripcion");
            if (norm.Contains("unidad") || norm.Contains("unit")) return tp.Fields.SingleOrDefault(f => f.Key == "Unidad");
        }
        if (tp.Name == "Movimientos")
        {
            if (norm.Contains("costo unitario") || norm.Contains("precio unitario") || norm.Contains("unit cost") || norm.Contains("unit price")) return tp.Fields.SingleOrDefault(f => f.Key == "CostoUnitario");
            if (norm.Contains("cantidad") || norm.Contains("quantity") || norm.Contains("amount")) return tp.Fields.SingleOrDefault(f => f.Key == "Cantidad");
            if (norm.Contains("fecha") || norm.Contains("date")) return tp.Fields.SingleOrDefault(f => f.Key == "Fecha");
            if (norm.Contains("tipo") || norm.Contains("type")) return tp.Fields.SingleOrDefault(f => f.Key == "Tipo");
            if (norm.Contains("visita") || norm.Contains("visit")) return tp.Fields.SingleOrDefault(f => f.Key == "Visita");
            if (norm.Contains("material")) return tp.Fields.SingleOrDefault(f => f.Key == "Material");
        }
        return null;
    }

    static FieldProfile? FindField(TableProfile tp, string raw)
    {
        var n = Norm(raw);
        return tp.Fields.SingleOrDefault(f => Norm(f.Key) == n || Norm(f.Header) == n) ?? MatchFieldInText(n, tp);
    }

    static string ResolveSheetName(string candidate, WorkbookSnapshot snapshot, WorkbookProfile profile)
    {
        var n = Norm(candidate);
        if (n is "inventory" or "inventories") n = "inventario";
        else if (n is "visit" or "visits") n = "visitas";
        else if (n is "movement" or "movements") n = "movimientos";
        else if (n is "summary") n = "resumen";
        foreach (var sh in snapshot.Sheets)
            if (Norm(sh) == n || Norm(sh).TrimEnd('s') == n.TrimEnd('s')) return sh;
        return candidate;
    }

    static string? DetectTable(string norm, WorkbookProfile profile)
    {
        if (Regex.Match(norm, @"(?:pesta[ñn]a|hoja|tabla|tab|sheet|table)\s+(?:de\s+|of\s+)?(inventario|inventory|visitas?|visits?|movimientos?|movements?)") is { Success: true } m)
            return MapTableWord(m.Groups[1].Value);
        if (Regex.Match(norm, @"\b(inventario|inventory|visitas?|visits?|movimientos?|movements?)\s+(?:pesta[ñn]a|hoja|tabla|tab|sheet|table)\b") is { Success: true } m2)
            return MapTableWord(m2.Groups[1].Value);
        if (Regex.IsMatch(norm, @"\binventarios?\b|\binventory\b|\bmaterial(?:es|s)?\b|\balmacen\b|\bwarehouse\b")) return "Inventario";
        if (Regex.IsMatch(norm, @"\bmovimientos?\b|\bmovements?\b|\bconsumos?\b|\bconsumptions?\b|\bentradas?\b")) return "Movimientos";
        if (Regex.IsMatch(norm, @"\bvisitas?\b|\bvisits?\b")) return "Visitas";
        return null;
    }

    static string MapTableWord(string w) => w.StartsWith("inv") ? "Inventario" : w.StartsWith("vis") ? "Visitas" : "Movimientos";

    static readonly (string Pattern, int Index)[] Ordinals =
    [
        (@"primera|primer|1ra|1er|1a|first|1st", 1),
        (@"segunda|segnda|2da|2a|segundo|second|2nd", 2),
        (@"tercera|tercer|3ra|3a|tercero|third|3rd", 3),
        (@"cuarta|4ta|4a|cuarto|fourth|4th", 4),
        (@"quinta|5ta|5a|quinto|fifth|5th", 5),
        (@"sexta|6ta|6a|sexto|sixth|6th", 6),
        (@"septima|7ma|7a|septimo|seventh|7th", 7),
        (@"octava|8va|8a|octavo|eighth|8th", 8),
        (@"novena|9na|9a|noveno|ninth|9th", 9),
        (@"decima|10ma|10a|decimo|tenth|10th", 10),
        (@"penultima|penultimo|second\s+to\s+last", -2),
        (@"ultima|ultimo|last", -1)
    ];

    static List<int> ExtractRowReferences(string norm, TableProfile tp, TableState ts)
    {
        var matches = new List<(int Pos, int Row)>();
        var ordExpr = string.Join("|", Ordinals.Select(o => o.Pattern));
        foreach (Match m in Regex.Matches(norm, $@"\b(?:({ordExpr})\s+(?:fila|registro|renglon|row|record)|(?:fila|registro|renglon|row|record)\s+(?:numero\s+|number\s+)?(\d+|{ordExpr})|(?:la|el|the)\s+({ordExpr})(?!\s+(?:columna|column|hoja|pesta[ñn]a|sheet|tab|vez|time)))\b"))
        {
            var token = m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Success ? m.Groups[2].Value : m.Groups[3].Value;
            int idx = ResolveRowIndex(token, ts.Rows.Count);
            if (idx > 0) matches.Add((m.Index, idx));
        }
        var idPrefix = tp.Name == "Inventario" ? @"m\d{3}" : tp.Name == "Visitas" ? @"v\d{3}" : @"mov\d{3}";
        var idHeader = tp.Fields.Single(f => f.Key == "Id").Header;
        foreach (Match m in Regex.Matches(norm, $@"\b({idPrefix})\b"))
        {
            var idx = ts.Rows.FindIndex(r => string.Equals(r[idHeader], m.Groups[1].Value, StringComparison.OrdinalIgnoreCase)) + 1;
            if (idx > 0) matches.Add((m.Index, idx));
        }
        return matches.OrderBy(x => x.Pos).Select(x => x.Row).ToList();
    }

    static List<int> ExtractRowRangeOrList(string norm, TableProfile tp, TableState ts)
    {
        // "las primeras N filas" / "the first N rows" / "las últimas N filas" / "the last N rows"
        if (Regex.Match(norm, @"\b(primeras|ultimas|first|last)\s+(\d+)\s+(?:filas|registros|rows|records)\b") is { Success: true } mult && int.TryParse(mult.Groups[2].Value, out var count) && count >= 1 && count <= ts.Rows.Count)
        {
            return mult.Groups[1].Value is "primeras" or "first"
                ? Enumerable.Range(1, count).ToList()
                : Enumerable.Range(ts.Rows.Count - count + 1, count).ToList();
        }
        // "filas X a Y" / "rows X to Y"
        if (Regex.Match(norm, @"\b(?:filas?|rows?)\s+(\d+)\s+(?:a|al|hasta|to|through)\s+(?:la\s+|the\s+)?(\d+)\b") is { Success: true } rng && int.TryParse(rng.Groups[1].Value, out var a) && int.TryParse(rng.Groups[2].Value, out var b))
        {
            int start = Math.Min(a, b), end = Math.Max(a, b);
            return Enumerable.Range(start, end - start + 1).ToList();
        }
        return ExtractRowReferences(norm, tp, ts);
    }

    static int ResolveRowIndex(string token, int total)
    {
        if (int.TryParse(token, out var n)) return n;
        foreach (var (pat, idx) in Ordinals)
            if (Regex.IsMatch(token, $@"^(?:{pat})$"))
                return idx < 0 ? total + 1 + idx : idx;
        return 0;
    }

    static List<(string Table, string Id)> ExtractExplicitIds(string text)
    {
        var list = new List<(string, string)>();
        foreach (Match m in Regex.Matches(text, @"\b(MOV\d{3}|M\d{3}|V\d{3})\b", RegexOptions.IgnoreCase))
        {
            var id = m.Groups[1].Value.ToUpperInvariant();
            var tbl = id.StartsWith("MOV") ? "Movimientos" : id.StartsWith("M") ? "Inventario" : "Visitas";
            list.Add((tbl, id));
        }
        return list;
    }

    static bool TryExtractDate(string text, out string iso)
    {
        iso = "";
        if (Regex.Match(text, @"\b(\d{4}-\d{2}-\d{2})\b") is { Success: true } isoM && DateTime.TryParseExact(isoM.Groups[1].Value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
        {
            iso = isoM.Groups[1].Value;
            return true;
        }
        var months = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["enero"] = 1, ["febrero"] = 2, ["marzo"] = 3, ["abril"] = 4, ["mayo"] = 5, ["junio"] = 6,
            ["julio"] = 7, ["agosto"] = 8, ["septiembre"] = 9, ["setiembre"] = 9, ["octubre"] = 10, ["noviembre"] = 11, ["diciembre"] = 12,
            ["january"] = 1, ["february"] = 2, ["march"] = 3, ["april"] = 4, ["may"] = 5, ["june"] = 6,
            ["july"] = 7, ["august"] = 8, ["september"] = 9, ["october"] = 10, ["november"] = 11, ["december"] = 12
        };
        if (Regex.Match(Norm(text), @"\b(\d{1,2})\s+(?:de\s+|of\s+)?([a-z]+)\s*,?\s*(?:de\s+|del\s+)?(\d{4})\b") is { Success: true } esDate && months.TryGetValue(esDate.Groups[2].Value, out var month) && int.TryParse(esDate.Groups[1].Value, out var day) && int.TryParse(esDate.Groups[3].Value, out var year))
        {
            try { iso = new DateTime(year, month, day).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture); return true; } catch { return false; }
        }
        if (Regex.Match(Norm(text), @"\b([a-z]+)\s+(\d{1,2})\s*,?\s*(\d{4})\b") is { Success: true } enDate && months.TryGetValue(enDate.Groups[1].Value, out var enMonth) && int.TryParse(enDate.Groups[2].Value, out var enDay) && int.TryParse(enDate.Groups[3].Value, out var enYear))
        {
            try { iso = new DateTime(enYear, enMonth, enDay).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture); return true; } catch { return false; }
        }
        return false;
    }

    static string FormatValueForWrite(FieldProfile f, string raw)
    {
        if (f.Type == "date" && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var oa) && oa >= 1 && oa <= 2958465)
            return DateTime.FromOADate(oa).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        return raw;
    }

    static string Norm(string s)
    {
        var formD = s.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder(formD.Length);
        foreach (var ch in formD)
            if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                sb.Append(char.ToLowerInvariant(ch));
        return sb.ToString();
    }
}
