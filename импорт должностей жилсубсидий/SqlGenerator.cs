using System.Text;

public class SqlGenerator
{
    private readonly List<PositionEntry> _entries = [];
    private int _nnumber = 1;

    /// <summary>
    /// Добавляет запись иерархии в список для последующей генерации SQL.
    /// </summary>
    /// <param name="node">Узел иерархии для добавления.</param>
    /// <param name="sectionName">Имя раздела, связанного с узлом.</param>
    public void AddHierarchyEntry(Узел node, string sectionName)
    {
        if (string.IsNullOrEmpty(node.Name)) return;

        var entry = new PositionEntry
        {
            DictId = node.DictId,
            ParentDictId = node.ParentDictId,
            Name = node.Name,
            RegNumber = GenerateRegNumber(node, sectionName),
            SectionName = sectionName,
            Nnumber = _nnumber++
        };
        _entries.Add(entry);
    }

    /// <summary>
    /// Добавляет запись должности в список для последующей генерации SQL.
    /// </summary>
    /// <param name="name">Наименование должности.</param>
    /// <param name="regNumber">Регистрационный номер должности.</param>
    /// <param name="parentNode">Родительский узел (группа).</param>
    /// <param name="sectionName">Имя раздела, связанного с должностью.</param>
    public void AddPositionEntry(string name, string regNumber, Узел parentNode, string sectionName)
    {
        var entry = new PositionEntry
        {
            ParentDictId = parentNode.DictId,
            Name = name,
            RegNumber = regNumber,
            SectionName = sectionName,
            Nnumber = _nnumber++
        };
        _entries.Add(entry);
    }

    /// <summary>
    /// Генерирует регистрационный номер для узла иерархии на основе его типа и родителей.
    /// </summary>
    /// <param name="node">Узел иерархии для генерации номера.</param>
    /// <param name="sectionName">Имя раздела (не используется в текущей реализации).</param>
    /// <returns>Сгенерированный регистрационный номер.</returns>
    /// <exception cref="InvalidOperationException">Выбрасывается, если тип узла не поддерживается.</exception>
    private static string GenerateRegNumber(Узел node, string sectionName)
    {
        // Уровень REGNUMBER           Пример
        // Раздел          XX-0-0-0R0          11-0-0-0R0
        // Подраздел       XX-0-0-AR0          11-0-0-3R0
        // Глава           XX-0-0-ARB          11-0-0-3R6
        // Категория       XX-Y-0-ARB          11-3-0-3R6
        // Группа          XX-Y-Z-ARB          11-3-4-3R6
        // Должность       XX-Y-Z-abc          11-3-4-001

        // Начальные значения для всех частей номера
        string раздел = "00";   // XX - значение раздела
        string категория = "0";   // Y - значение категории
        string группа = "0";      // Z - значение группы
        string подраздел = "0"; // A - значение подраздела
        string глава = "0";    // B - значение главы

        // Текущий узел
        var currentNode = node;

        // Проходим по иерархии вверх, чтобы найти значения для всех частей номера
        while (currentNode != null)
        {
            if      (currentNode is Раздел)
                раздел = currentNode.Value;
            else if (currentNode is Подраздел)
                подраздел = currentNode.Value;
            else if (currentNode is Глава)
                глава = currentNode.Value;
            else if (currentNode is Категория)
                категория = currentNode.Value;
            else if (currentNode is Группа)
                группа = currentNode.Value;
            
            // Переходим к родителю
            currentNode = currentNode.Parent;
        }

        // Формируем номер в зависимости от типа узла
        
        
        if      (node is Раздел)
            return $"{раздел}-0-0-0R0";
        else if (node is Подраздел)
            return $"{раздел}-0-0-{подраздел}R0";
        else if (node is Глава)
            return $"{раздел}-0-0-{подраздел}R{глава}";
        else if (node is Категория)
            return $"{раздел}-{категория}-0-{подраздел}R{глава}";
        else if (node is Группа)
            return $"{раздел}-{категория}-{группа}-{подраздел}R{глава}";
        else
            throw new InvalidOperationException($"Unsupported node type: {node.GetType().Name}");
    }

    /// <summary>
    /// Генерирует SQL-запросы для вставки данных в таблицу.
    /// </summary>
    /// <returns>Сгенерированный SQL-скрипт для вставки данных.</returns>
    /// <exception cref="Exception">Выбрасывается, если список записей пуст.</exception>
    public string GenerateInsertStatements()
    {
        if (_entries.Count == 0) throw new Exception("пусто");

        // Разбиваем записи на пакеты по 1000 для оптимизации вставки
        const int batchSize = 1000;
        var sql = new StringBuilder();
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

    /// <summary>
    /// Генерирует SQL-скрипт для создания таблицы output.
    /// </summary>
    /// <returns>SQL-скрипт для создания таблицы.</returns>
    public static string AddTableCreationSql() =>
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
}