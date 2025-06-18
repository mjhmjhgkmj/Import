// PositionProcessor.cs
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Reflection.Emit;
using System.Text.RegularExpressions;

namespace импорт_должностей;
public class PositionProcessor
{
    private readonly HierarchyTree _hierarchyTree;
    private readonly SqlGenerator _sqlGenerator;

    public PositionProcessor(SqlGenerator sqlGenerator)
    {
        _hierarchyTree = new HierarchyTree();
        _sqlGenerator = sqlGenerator;
    }

    public void ProcessPositionRow(List<TableCell> cells)
    {
        string name = GetCellText(cells[0]).Replace("'", "''");
        string regNumber = cells[1].InnerText.Trim();

        if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(regNumber))
            return;

        regNumber = ProcessRegNumber(regNumber);
        _sqlGenerator.AddPositionRecord(name, regNumber, _hierarchyTree.GetCurrentNode()?.Parent, _hierarchyTree.GetSectionNode()?.Name ?? string.Empty);
    }

    private static string GetCellText(TableCell cell) =>
        string.Join(" ", cell.Descendants<Paragraph>()
            .Select(p => string.Join(" ",
                p.Descendants<Text>()
                .Where(t => !t.Ancestors<Hyperlink>().Any())
                .Select(t => t.Text.Trim())))
            .Where(s => !string.IsNullOrEmpty(s)))
        .Replace("'", "''");

    private static string ProcessRegNumber(string regNumber)
    {
        var match = Regex.Match(regNumber, @"(\d{2}-\d-\d-\d{3})\s*<(\d+)>");
        if (match.Success)
            return $"{match.Groups[1].Value}.{match.Groups[2].Value}";
        return regNumber;
    }
}
