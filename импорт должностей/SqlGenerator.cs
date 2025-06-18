using System.Text;

public class SqlGenerator
{
    private readonly List<SQLPositionRecord> _SQLRecords = [];
    private int _nnumber = 1;

    /// <summary>
    /// Добавляет запись в список для последующей генерации SQL.
    /// </summary>
    /// <param name="node">Узел иерархии для добавления.</param>
    /// <param name="sectionName">Имя раздела, связанного с узлом.</param>
    public void AddSQLRecord(Node node, string sectionName)
    {
        if (string.IsNullOrEmpty(node.Name)) return;

        var record = new SQLPositionRecord(node.ParentDictId, node.Name, GenerateRegNumber(node), sectionName, _nnumber++)
        {
            DictId = node.DictId
        };
        _SQLRecords.Add(record);
    }

    /// <summary>
    /// Добавляет запись должности в список для последующей генерации SQL.
    /// </summary>
    /// <param name="name">Наименование должности.</param>
    /// <param name="regNumber">Регистрационный номер должности.</param>
    /// <param name="parentNode">Родительский узел (группа).</param>
    /// <param name="sectionName">Имя раздела, связанного с должностью.</param>
    public void AddPositionRecord(string name, string regNumber, Node parentNode, string sectionName)
    {
        var record = new SQLPositionRecord(parentNode.DictId, name, regNumber, sectionName, _nnumber++);
        _SQLRecords.Add(record);
    }

    /// <summary>
    /// Генерирует регистрационный номер для узла иерархии на основе его типа и родителей.
    /// </summary>
    /// <param name="node">Узел иерархии для генерации номера.</param>
    /// <param name="sectionName">Имя раздела (не используется в текущей реализации).</param>
    /// <returns>Сгенерированный регистрационный номер.</returns>
    /// <exception cref="InvalidOperationException">Выбрасывается, если тип узла не поддерживается.</exception>
        // code Уровень          REGNUMBER           Пример
        // XX   Раздел           XX-0-0-0R0          11-0-0-0R0
        // A    Подраздел        XX-0-0-AR0          11-0-0-3R0
        // B    Глава            XX-0-0-ARB          11-0-0-3R6
        // Y    Категория        XX-Y-0-ARB          11-3-0-3R6
        // Z    Группа           XX-Y-Z-ARB          11-3-4-3R6
        // abc  Должность        XX-Y-Z-abc          11-3-4-001
private static string GenerateRegNumber(Node node)
    {
        // дефолтные значения
        string section  = "00",
               category = "0",
               group    = "0",
               sub      = "0",
               chapter  = "0";

        // поднимаемся к корню и берем нужные значения
        for (var n = node; n is not null; n = n.Parent)
            switch (n)
            {
                case Раздел     s: section  = s.Value; break;
                case Категория  c: category = c.Value; break;
                case Группа     g: group    = g.Value; break;
                case Подраздел  p: sub      = p.Value; break;
                case Глава      h: chapter  = h.Value; break;
            }

        return $"{section}-{category}-{group}-{sub}R{chapter}";
    }

    /// <summary>
    /// Генерирует SQL-запросы для вставки данных в таблицу.
    /// </summary>
    /// <returns>Сгенерированный SQL-скрипт для вставки данных.</returns>
    /// <exception cref="Exception">Выбрасывается, если список записей пуст.</exception>
    public string GenerateInsertStatements()
    {
        if (_SQLRecords.Count == 0) throw new Exception("пусто");

        // Разбиваем записи на пакеты по 1000 для оптимизации вставки
        const int batchSize = 1000;
        var sql = new StringBuilder();
        for (int i = 0; i < _SQLRecords.Count; i += batchSize)
        {
            var batch = _SQLRecords.Skip(i).Take(batchSize).ToList();
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
    public static string AddTableCreationSql =>
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