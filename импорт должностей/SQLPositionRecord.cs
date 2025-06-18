namespace импорт_должностей;

public record SQLPositionRecord(string? ParentDictId, string? Name, string? RegNumber, string? SectionName, int Nnumber)
{
    public string? DictId { get; set; } = Guid.NewGuid().ToString();

    public string ToValueString()
    {
        string parentDictId = ParentDictId != null ? $"'{ParentDictId}'" : "NULL";
        string safeName = Name?.Replace("'", "''") ?? "нет имени";
        string safeSectionName = SectionName?.Replace("'", "''") ?? "";
        return $"('{DictId}', {parentDictId}, '{RegNumber}', '{safeName}', {Nnumber}, NULL, '{safeSectionName}')";
    }
}