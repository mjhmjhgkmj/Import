// HierarchyProcessor.cs
using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace импорт_должностей
{
    public class HierarchyProcessor
    {
        private readonly HierarchyTree _hierarchyTree;
        private readonly TextParser _textParser;
        private readonly SqlGenerator _sqlGenerator;

        public HierarchyProcessor(SqlGenerator sqlGenerator)
        {
            _hierarchyTree = new HierarchyTree();
            _textParser = new TextParser();
            _sqlGenerator = sqlGenerator;
        }

        public void ProcessHierarchyRow(TableCell cell)
        {
            string text = GetCellText(cell);
            if (string.IsNullOrEmpty(text)) return;

            var parser = _textParser.GetParser(text);
            if (parser == null) return;// мусор игнорируем

            var (value, name, node) = parser.Parse(text);

            if (parser is ПарсерРаздела)// чтобы не усложнять парсинг лучше бы привести таблицу к хорошему виду заранее
            {
                _hierarchyTree.НачатьНовыйРаздел(value, name);
            }
            else if (parser is ПарсерИмениРаздела)
            {
                _hierarchyTree.ОбновитьИмяРаздела($"Раздел {_hierarchyTree.GetSectionNode()?.Value} {name}");
                _sqlGenerator.AddSQLRecord(_hierarchyTree.GetSectionNode()!, name);
            }
            else if (node != null)
            {
                _hierarchyTree.AddNode(node);
                _sqlGenerator.AddSQLRecord(node, _hierarchyTree.GetSectionNode()?.Name ?? string.Empty);
            }
        }

        private static string GetCellText(TableCell cell) =>
            string.Join(" ", cell.Descendants<Paragraph>()
                .Select(p => string.Join(" ",
                    p.Descendants<Text>()
                    .Where(t => !t.Ancestors<Hyperlink>().Any())
                    .Select(t => t.Text.Trim())))
                .Where(s => !string.IsNullOrEmpty(s)))
            .Replace("'", "''");
    }
}