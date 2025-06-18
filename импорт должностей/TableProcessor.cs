// TableProcessor.cs
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace импорт_должностей
{
    public class TableProcessor
    {
        private readonly HierarchyProcessor _hierarchyProcessor;
        private readonly PositionProcessor _positionProcessor;

        public TableProcessor(SqlGenerator sqlGenerator)
        {
            _hierarchyProcessor = new HierarchyProcessor(sqlGenerator);
            _positionProcessor = new PositionProcessor(sqlGenerator);
        }

        public void ProcessTable(Table table)
        {
            var rows = table.Elements<TableRow>();
            foreach (var row in rows.Skip(1))
            {
                var cells = row.Elements<TableCell>().ToList();
                if (cells.Count == 1 || cells[0].InnerText.Contains("группа должностей"))
                {
                    _hierarchyProcessor.ProcessHierarchyRow(cells[0]);
                }
                else if (cells.Count == 2)
                {
                    _positionProcessor.ProcessPositionRow(cells);
                }
            }
        }
    }
}