// DocxReader.cs
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace импорт_должностей
{
    public class DocxReader
    {
        public IEnumerable<Table> GetValidTables(WordprocessingDocument doc)
        {
            return doc.MainDocumentPart?.Document.Body?.Elements<Table>()
                .Where(IsValidTable)
                .ToList();
        }

        private static bool IsValidTable(Table table)
        {
            var firstRow = table.Elements<TableRow>().FirstOrDefault();
            var headerCells = firstRow?.Elements<TableCell>().ToList();
            return IsValidHeader(headerCells);
        }

        private static bool IsValidHeader(List<TableCell>? cells) =>
            cells != null && cells.Count == 2 &&
            cells[0].InnerText.Contains("Наименование должности") &&
            cells[1].InnerText.Contains("Регистрационный номер (код)");
    }
}