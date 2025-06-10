/*
структура кодов:


Правила генерации для каждого уровня:
ARB - кодирует подраздел и главу
но его формат не регламентируется, поэтому бывают разные форматы

Вариант 1 (300 штук) - пишем так!
Уровень         REGNUMBER           Пример
Раздел          XX-0-0-0R0          11-0-0-0R0
Подраздел       XX-0-0-AR0          11-0-0-3R0
Глава           XX-0-0-ARB          11-0-0-3R6
Категория       XX-Y-0-ARB          11-3-0-3R6
Группа          XX-Y-Z-ARB          11-3-4-3R6
Должность       XX-Y-Z-abc          11-3-4-001

Уровень         REGNUMBER           Пример
Раздел          XX-0-0-0R0          11-0-0-0R0
Категория       XX-Y-0-ARB          11-3-0-0R0
Группа          XX-Y-Z-ARB          11-3-4-0R0
Должность       XX-Y-Z-abc          11-3-4-001

Уровень         REGNUMBER           Пример
Раздел          XX-0-0-0R0          15.1-0-0-0R0
Глава           XX-0-0-ARB          15.1-0-0-0R2
Категория       XX-Y-0-ARB          15.1-3-0-0R2
Группа          XX-Y-Z-ARB          15.1-3-4-0R2
Должность       XX-Y-Z-abc          15.1-3-4-001

Вариант 2 (100 штук) - валидно, но новые так не пишем!
Уровень         REGNUMBER           Пример
Раздел          XX-0-0-0R0          11-0-0-0R0
Подраздел       XX-0-0-0RX          11-0-0-0R1
Глава           XX-0-0-0RXY         11-0-0-0R16
Категория       XX-Y-0-0RXY         11-1-0-0R16
Группа          XX-Y-Z-0RXY         11-1-2-0R16
Должность       XX-Y-Z-abc          11-1-2-001

есть и другие варианты!

Мутаторы: допустимы дополнительные разряды в разделе и должности
Раздел          XX.X-0-0-0R0        15.1-0-0-0R2
Должность       XX-YY-ZZ-AAA.A      11-3-4-001.1

если в тексте сразу после 15 раздела идет раздел 15.1, то 15 игнорируется, он пустой


Должность: XX-Y-Z-ABC или XX-Y-Z-ABC.D
ABC - цифры для уникального номера должности, нумерация должностей сквозная в разделе



14-3-4-047 <- 14-3-4-2R0 <- 14-3-0-2R0 <- 14-0-0-2R0 <- 14-0-0-0R0 <- NULL
должность 047  группа 4     категория 3   подраздел 2    раздел 14

11-1-2-050.1 <- 11-1-2-0R16 <- 11-1-0-0R16 <- 11-0-0-0R16 <- 11-0-0-0R1 <- 11-0-0-0R0 <- NULL
должность 050.1  группа 2   категория 1       глава 6        подраздел 1    раздел 11


Особенности:

Нумерация должностей (ABC часть) продолжается сквозным образом через подразделы и главы. С нового раздела начинается новая нумерация.
В главах используется формат YRZ, где Y - номер подраздела, Z - номер главы.
Некоторые должности имеют дополнительный номер после точки (например, 11-1-3-029.1).
Категории должностей: 1 - руководители, 2 - помощники (советники), 3 - специалисты, 4 - обеспечивающие специалисты.
Группы должностей: 1 - высшая, 2 - главная, 3 - ведущая, 4 - старшая, 5 - младшая.

 */


using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Text.RegularExpressions;

namespace импорт_должностей_жилсубсидий_олд;


// Перечисление для определения уровней иерархии в документе
public enum HierarchyLevel
{
    Section,    // Раздел
    Subsection, // Подраздел
    Chapter,    // Глава
    Category,   // Категория
    Group       // Группа
}

// Класс для хранения базовой информации об элементе иерархии
public class HierarchyItem
{
    public HierarchyLevel Level { get; set; }           // Уровень иерархии
    public string? Value { get; set; }                  // Значение (например, номер раздела)
    public string? Name { get; set; }                   // Название элемента
}

// Класс для представления узла в иерархической структуре
public class HierarchyNode
{
    public required HierarchyItem Item { get; set; }     // Данные узла
    public string? DictId { get; set; } = Guid.NewGuid().ToString();                  // Уникальный идентификатор
    public string? ParentDictId { get; set; }            // Идентификатор родителя
    public HierarchyNode? Parent { get; set; }           // Ссылка на родительский узел
}

// Класс для представления записи в базе данных
public class PositionEntry
{
    public string? DictId { get; set; } = Guid.NewGuid().ToString();                 // Уникальный идентификатор
    public string? ParentDictId { get; set; }            // Идентификатор родителя
    public string? Name { get; set; }                    // Название должности
    public string? RegNumber { get; set; }               // Регистрационный номер
    public string? SectionName { get; set; }             // Название раздела
    public int Nnumber { get; set; }                     // Порядковый номер

    public string ToValueString()
    {
        string parentDictId = ParentDictId != null ? $"'{ParentDictId}'" : "NULL";
        string safeName = Name?.Replace("'", "''")??"нет имени";
        string safeSectionName = SectionName?.Replace("'", "''") ?? "";
        return $"('{DictId}', {parentDictId}, '{RegNumber}', '{safeName}', {Nnumber}, NULL, '{safeSectionName}')";
    }
}

// Состояние иерархии
public class HierarchyState
{
    private readonly Dictionary<HierarchyLevel, HierarchyNode> _nodes = [];

    public void Clear() => _nodes.Clear();

    public void AddNode(HierarchyItem item, HierarchyNode? parent = null)
    {
        var node = new HierarchyNode
        {
            Item = item,
            Parent = parent,
            ParentDictId = parent?.DictId
        };
        _nodes[item.Level] = node;
        Console.WriteLine($"node id =  { node.DictId}, parent = {parent?.DictId}, Name = { item.Name}");
    }

    public HierarchyNode? GetNode(HierarchyLevel level) =>
        _nodes.GetValueOrDefault(level);

    public HierarchyNode? GetParentNode()
    {
        if (_nodes.TryGetValue(HierarchyLevel.Chapter, out var chapter))
            return chapter;
        if (_nodes.TryGetValue(HierarchyLevel.Subsection, out var subsection))
            return subsection;
        return _nodes.GetValueOrDefault(HierarchyLevel.Section);
    }

    public void RemoveDescendants(HierarchyLevel level)
    {
        var levelsToRemove = _nodes.Keys.Where(k => k > level).ToList();
        foreach (var key in levelsToRemove)
            _nodes.Remove(key);
    }

    public string GenerateRegNumber()
    {
        var section         = _nodes.GetValueOrDefault(HierarchyLevel.Section)      ?.Item.Value ?? "0";
        var category        = _nodes.GetValueOrDefault(HierarchyLevel.Category)     ?.Item.Value ?? "0";
        var group           = _nodes.GetValueOrDefault(HierarchyLevel.Group)        ?.Item.Value ?? "0";
        var subsection      = _nodes.GetValueOrDefault(HierarchyLevel.Subsection)   ?.Item.Value ?? "0";
        var chapter         = _nodes.GetValueOrDefault(HierarchyLevel.Chapter)      ?.Item.Value ?? "0";
        return $"{section}-{category}-{group}-{subsection}R{chapter}";
    }
}

public class DocxParser
{
    private readonly List<PositionEntry> _entries = [];
    private readonly HierarchyState _hierarchyState = new();
    private int _nnumber = 1;

    public string ParseAndGenerateSql(WordprocessingDocument doc)
    {
        var sql = AddTableCreationSql();
        var t = doc.MainDocumentPart;
        var body = doc.MainDocumentPart?.Document.Body
            ?? throw new Exception("Некорректная структура документа");

        body.Elements<Table>()
            .Where(IsValidTable)
            .ToList()
            .ForEach(ProcessTable);

        sql += GenerateInsertStatements();
        return sql;
    }

    private static string AddTableCreationSql() =>
        """
        IF OBJECT_ID('output', 'U') IS NOT NULL 
            DROP TABLE [output];
        GO

        CREATE TABLE [output] (
            DICT_ID UNIQUEIDENTIFIER NOT NULL,
            DICT_PARENT UNIQUEIDENTIFIER NULL,
            NAME NVARCHAR(2000) NOT NULL,
            REGNUMBER NVARCHAR(20) NOT NULL,
            NNUMBER INT NOT NULL,
            ORCL_ID NVARCHAR(50) NULL,
            S_SECTION_NAME NVARCHAR(500) NULL,
            CONSTRAINT PK_output PRIMARY KEY (DICT_ID),
            CONSTRAINT FK_output_DICT_PARENT FOREIGN KEY (DICT_PARENT) 
                REFERENCES [output] (DICT_ID)
        );
        GO

    """;

    private string GenerateInsertStatements()
    {
        StringBuilder sql = new();
        if (_entries.Count == 0) throw new Exception("пусто");

        const int batchSize = 1000;
        for (int i = 0; i < _entries.Count; i += batchSize)
        {
            var batch = _entries.Skip(i).Take(batchSize).ToList();
            sql.AppendLine("INSERT INTO [output] (DICT_ID, DICT_PARENT, REGNUMBER, NAME, NNUMBER, ORCL_ID, S_SECTION_NAME)");
            sql.AppendLine("VALUES");

            for (int j = 0; j < batch.Count; j++)
            {
                sql.Append(batch[j].ToValueString());
                if (j < batch.Count - 1) sql.AppendLine(",");
            }

            sql.AppendLine(";");
            sql.AppendLine();
        }
        sql.Append("GO");
        return sql.ToString();
    }

    private bool IsValidTable(Table table)
    {
        var firstrow = table.Elements<TableRow>().FirstOrDefault();
        var headerCells = firstrow?.Elements<TableCell>().ToList();
        return IsValidHeader(headerCells);
    }

    private void ProcessTable(Table table)
    {
        var rows = table.Elements<TableRow>().ToList();

        foreach (var row in rows.Skip(1))
        {
            var cells = row.Elements<TableCell>().ToList();
            if (cells.Count == 1)
                ProcessHierarchyRow(cells[0]);
            else if (cells.Count == 2)
                ProcessPositionRow(cells);
        }
    }

    private void ProcessHierarchyRow(TableCell cell)
    {
        string text = GetCellText(cell);
        if (string.IsNullOrEmpty(text)) return;

        if (text.StartsWith("Раздел"))
            ProcessSection(text);
        else if (text.StartsWith("Перечень"))
            ProcessSectionName(text);
        else if (text.StartsWith("Подраздел"))
            ProcessSubsection(text);
        else if (text.StartsWith("Глава"))
            ProcessChapter(text);
        else if (text.Contains("категории"))
            ProcessCategory(text);
        else if (text.Contains("группа должностей"))
            ProcessGroup(text);
    }

    private void ProcessPositionRow(List<TableCell> cells)
    {
        string name = GetCellText(cells[0]).Replace("'", "''");
        string regNumber = cells[1].InnerText.Trim();

        if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(regNumber) ||
            _hierarchyState.GetNode(HierarchyLevel.Group) == null)
            return;

        regNumber = ProcessRegNumber(regNumber);
        AddPositionEntry(name, regNumber, HierarchyLevel.Group);
        Console.WriteLine($"regNumber = {regNumber}, Name = {name}");
    }

    private void ProcessSection(string text)
    {
        _hierarchyState.Clear();
        var item = new HierarchyItem
        {
            Level = HierarchyLevel.Section,
            Value = NormalizeSectionValue(text),
            Name = string.Empty
        };
        _hierarchyState.AddNode(item);
    }

    private void ProcessSectionName(string text)
    {
        var sectionNode = _hierarchyState.GetNode(HierarchyLevel.Section);
        if (sectionNode != null)
        {
            sectionNode.Item.Name = "Раздел " + sectionNode.Item.Value + " " + text;
            AddHierarchyEntry(sectionNode);
        }
    }

    private void ProcessSubsection(string text)
    {
        var match = Regex.Match(text, @"Подраздел (\d+)");
        if (match.Success && _hierarchyState.GetNode(HierarchyLevel.Section) != null)
        {
            var item = new HierarchyItem
            {
                Level = HierarchyLevel.Subsection,
                Value = match.Groups[1].Value,
                Name = text
            };
            var parent = _hierarchyState.GetParentNode();
            _hierarchyState.AddNode(item, parent);
            AddHierarchyEntry(_hierarchyState.GetNode(HierarchyLevel.Subsection)!);
            _hierarchyState.RemoveDescendants(HierarchyLevel.Subsection);
        }
    }

    private void ProcessChapter(string text)
    {
        var match = Regex.Match(text, @"Глава (\d+)");
        if (match.Success && _hierarchyState.GetNode(HierarchyLevel.Section) != null)
        {
            var item = new HierarchyItem
            {
                Level = HierarchyLevel.Chapter,
                Value = match.Groups[1].Value,
                Name = text
            };
            var parent = _hierarchyState.GetParentNode();
            _hierarchyState.AddNode(item, parent);
            AddHierarchyEntry(_hierarchyState.GetNode(HierarchyLevel.Chapter)!);
            _hierarchyState.RemoveDescendants(HierarchyLevel.Chapter);
        }
    }

    private void ProcessCategory(string text)
    {
        var match = Regex.Match(text, @"(\d+)\.\s*(Должности категории.*)");
        if (match.Success && _hierarchyState.GetNode(HierarchyLevel.Section) != null)
        {
            var item = new HierarchyItem
            {
                Level = HierarchyLevel.Category,
                Value = match.Groups[1].Value,
                Name = text
            };
            var parent = _hierarchyState.GetParentNode();
            _hierarchyState.AddNode(item, parent);
            AddHierarchyEntry(_hierarchyState.GetNode(HierarchyLevel.Category)!);
            _hierarchyState.RemoveDescendants(HierarchyLevel.Category);
        }
    }

    private void ProcessGroup(string text)
    {
        if (_hierarchyState.GetNode(HierarchyLevel.Category) != null)
        {
            var item = new HierarchyItem
            {
                Level = HierarchyLevel.Group,
                Value = GetGroupValue(text),
                Name = text
            };
            var parent = _hierarchyState.GetNode(HierarchyLevel.Category);
            _hierarchyState.AddNode(item, parent);
            AddHierarchyEntry(_hierarchyState.GetNode(HierarchyLevel.Group)!);
        }
    }

    private static string GetCellText(TableCell cell) =>
        string.Join(" ", cell.Descendants<Paragraph>()
            .Select(p => string.Join(" ",
                p.Descendants<Text>()
                .Where(t => !t.Ancestors<Hyperlink>().Any()) // Исключаем текст, находящийся внутри гиперссылок
                .Select(t => t.Text.Trim())))
            .Where(s => !string.IsNullOrEmpty(s)))
        .Replace("'", "''");

    private static string NormalizeSectionValue(string value)
    {
        value = value.Replace("Раздел ", "");
        if (value.Contains('.'))
        {
            var parts = value.Split('.');
            if (int.TryParse(parts[0], out int integerPart))
                parts[0] = integerPart.ToString("D2");
            return string.Join(".", parts);
        }
        else if (int.TryParse(value, out int number))
            return number.ToString("D2");
        return "error";
    }

    private static string ProcessRegNumber(string regNumber)
    {
        var match = Regex.Match(regNumber, @"(\d{2}-\d-\d-\d{3})\s*<(\d+)>");
        if (match.Success)
            return $"{match.Groups[1].Value}.{match.Groups[2].Value}";
        return regNumber;
    }

    private void AddHierarchyEntry(HierarchyNode node)
    {
        if (string.IsNullOrEmpty(node.Item.Name)) return;

        _entries.Add(new PositionEntry
        {
            DictId = node.DictId,
            ParentDictId = node.ParentDictId,
            Name = node.Item.Name,
            RegNumber = _hierarchyState.GenerateRegNumber(),
            SectionName = _hierarchyState.GetNode(HierarchyLevel.Section)?.Item.Name,
            Nnumber = _nnumber++
        });
    }

    private void AddPositionEntry(string name, string regNumber, HierarchyLevel parentLevel)
    {
        var parent = _hierarchyState.GetNode(parentLevel);
        _entries.Add(new PositionEntry
        {
            ParentDictId = parent?.DictId,
            Name = name,
            RegNumber = regNumber,
            SectionName = _hierarchyState.GetNode(HierarchyLevel.Section)?.Item.Name,
            Nnumber = _nnumber++
        });
    }

    private static string GetGroupValue(string text)
    {
        return text switch
        {
            _ when text.Contains("Высшая")  => "1",
            _ when text.Contains("Главная") => "2",
            _ when text.Contains("Ведущая") => "3",
            _ when text.Contains("Старшая") => "4",
            _ when text.Contains("Младшая") => "5",
            _ => "0"
        };
    }

    private static bool IsValidHeader(List<TableCell>? cells) =>
        cells != null && cells.Count == 2 &&
        cells[0].InnerText.Contains("Наименование должности") &&
        cells[1].InnerText.Contains("Регистрационный номер (код)");
}

public class Program_old
{
    public static void Main2(string[] args)
    {
        string docxFilePath = args.Length > 0 ? args[0] : "test.docx";
        string sqlFilePath = Path.ChangeExtension(docxFilePath, ".sql");

        if (!File.Exists(docxFilePath))
        {
            Console.WriteLine($"Файл '{docxFilePath}' не найден.");
            return;
        }

        string sqlContent = ConvertDocxToSqlInsert(docxFilePath);
        File.WriteAllText(sqlFilePath, sqlContent, Encoding.UTF8);
        Console.WriteLine($"Файл успешно преобразован и сохранен в {sqlFilePath}");
    }

    public static string ConvertDocxToSqlInsert(string filePath)
    {
        using WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false);
        var parser = new DocxParser();
        return parser.ParseAndGenerateSql(doc);
    }
}