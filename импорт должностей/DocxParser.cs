using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace импорт_должностей_жилсубсидий;

public class DocxParser
{
    private readonly HierarchyTree  _hierarchyTree  = new(); // Хранит иерархию узлов
    private readonly TextParser     _textParser     = new(); // Парсер текста для определения типа узлов
    private readonly SqlGenerator   _sqlGenerator   = new(); // Генератор SQL-запросов

    /// <summary>
    /// Парсит документ Word и генерирует SQL-скрипт для создания таблиц и вставки данных.
    /// </summary>
    /// <param name="doc">Документ Word для обработки.</param>
    /// <returns>Сгенерированный SQL-скрипт.</returns>
    public string ParseAndGenerateSql(WordprocessingDocument doc)
    {
        // Обрабатываем валидные таблицы в документе
        doc.MainDocumentPart?.Document.Body?.Elements<Table>()
            .Where(IsValidTable)
            .ToList()
            .ForEach(ProcessTable);

        // Добавляем SQL-запросы для вставки данных
        return SqlGenerator.AddTableCreationSql + _sqlGenerator.GenerateInsertStatements(); ;
    }


    /// <summary>
    /// Обрабатывает таблицу, разделяя строки на иерархические и должностные.
    /// </summary>
    /// <param name="table">Таблица для обработки.</param>
    private void ProcessTable(Table table)
    {
        // Получаем все строки таблицы
        var rows = table.Elements<TableRow>();

        // Пропускаем первую строку (заголовок) и обрабатываем остальные
        foreach (var row in rows.Skip(1))
        {
            var cells = row.Elements<TableCell>().ToList();
            //Console.WriteLine(cells[0].InnerText); // Выводим текст первой ячейки для отладки

            // Если строка содержит 1 ячейку — это иерархическая строка
            if (cells.Count == 1 || cells[0].InnerText.Contains("группа должностей"))
                ProcessHierarchyRow(cells[0]);
            // Если строка содержит 2 ячейки — это строка с должностью - с именем и кодом
            else if (cells.Count == 2)
                ProcessPositionRow(cells);
        }
    }

    /// <summary>
    /// Обрабатывает строку иерархии (раздел, подраздел, категория и т.д.).
    /// </summary>
    /// <param name="cell">Ячейка с текстом иерархической строки.</param>
    private void ProcessHierarchyRow(TableCell cell)
    {
        // Извлекаем текст из ячейки
        string text = GetCellText(cell);
        if (string.IsNullOrEmpty(text)) return; // Пропускаем пустые строки

        // Получаем парсер (например, для раздела, категории и т.д.)
        var parser = _textParser.GetParser(text);
        if (parser == null) return; // Если парсер не найден, пропускаем: это комментарий 

        // Парсим текст, получаем значение, имя и узел
        var (value, name, node) = parser.Parse(text);

        // Если это раздел
        if (parser is ПарсерРаздела)
        {
            // Начинаем новый раздел в иерархии
            _hierarchyTree.НачатьНовыйРаздел(value, name);
        }
        // Если это обновление имени раздела на следующей строке
        else if (parser is ПарсерИмениРаздела)
        {
            // Обновляем имя раздела
            _hierarchyTree.ОбновитьИмяРаздела($"Раздел {_hierarchyTree.GetSectionNode()?.Value} {name}");
            _sqlGenerator.AddSQLRecord(_hierarchyTree.GetSectionNode()!, name);
        }
        // Если это другой узел иерархии то обработка однотипная
        else if (node != null)
        {
            // Добавляем узел в иерархию и записываем в SQL
            _hierarchyTree.AddNode(node);
            _sqlGenerator.AddSQLRecord(node, _hierarchyTree.GetSectionNode()?.Name ?? string.Empty);
        }
    }

    /// <summary>
    /// Обрабатывает строку с должностью (наименование и регистрационный номер).
    /// </summary>
    /// <param name="cells">Ячейки строки (наименование и регистрационный номер).</param>
    private void ProcessPositionRow(List<TableCell> cells)
    {
        // Извлекаем наименование должности и экранируем кавычки
        string name = GetCellText(cells[0]).Replace("'", "''");
        // Извлекаем регистрационный номер
        string regNumber = cells[1].InnerText.Trim();

        // Проверяем, что данные валидны и текущий узел — группа
        if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(regNumber) ||
            _hierarchyTree.GetCurrentNode() is not Группа groupNode)
            return;

        // Обрабатываем регистрационный номер (например, добавляем суффикс)
        regNumber = ProcessRegNumber(regNumber);
        // Добавляем запись о должности в SQL
        _sqlGenerator.AddPositionRecord(name, regNumber, groupNode, _hierarchyTree.GetSectionNode()?.Name ?? string.Empty);
    }

    /// <summary>
    /// Проверяет, является ли таблица валидной (имеет правильные заголовки).
    /// </summary>
    /// <param name="table">Таблица для проверки.</param>
    /// <returns>True, если таблица валидна, иначе false.</returns>
    private static bool IsValidTable(Table table)
    {
        // Проверяем первую строку таблицы (заголовок)
        var firstRow = table.Elements<TableRow>().FirstOrDefault();
        var headerCells = firstRow?.Elements<TableCell>().ToList();
        return IsValidHeader(headerCells);
    }

    /// <summary>
    /// Проверяет, что заголовок таблицы содержит нужные колонки.
    /// </summary>
    /// <param name="cells">Ячейки заголовка таблицы.</param>
    /// <returns>True, если заголовок валиден, иначе false.</returns>
    private static bool IsValidHeader(List<TableCell>? cells) =>
        cells != null && cells.Count == 2 &&
        cells[0].InnerText.Contains("Наименование должности") &&
        cells[1].InnerText.Contains("Регистрационный номер (код)");

    /// <summary>
    /// Извлекает текст из ячейки, игнорируя гиперссылки и экранируя кавычки.
    /// </summary>
    /// <param name="cell">Ячейка для извлечения текста.</param>
    /// <returns>Обработанный текст ячейки.</returns>
    private static string GetCellText(TableCell cell) =>
        string.Join(" ", cell.Descendants<Paragraph>()
            .Select(p => string.Join(" ",
                p.Descendants<Text>()
                .Where(t => !t.Ancestors<Hyperlink>().Any()) // Игнорируем текст в гиперссылках
                .Select(t => t.Text.Trim())))
            .Where(s => !string.IsNullOrEmpty(s))) // Пропускаем пустые строки
        .Replace("'", "''"); // Экранируем кавычки для SQL

    /// 03-2-2-028 <1>  => 03-2-2-028.1
    /// 11-1-1-050 <32> => 11-1-1-050.32 (хотя можно бы и 11-1-1-050)
    private static string ProcessRegNumber(string regNumber)
    {
        // Проверяем формат номера с суффиксом (например, "11-3-4-001 <123>")
        var match = Regex.Match(regNumber, @"(\d{2}-\d-\d-\d{3})\s*<(\d+)>");
        if (match.Success)
            return $"{match.Groups[1].Value}.{match.Groups[2].Value}"; // Добавляем суффикс
        return regNumber; // Возвращаем исходный номер, если формат не соответствует
    }
}