public class PositionEntry
{
    public string? DictId { get; set; } = Guid.NewGuid().ToString();
    public string? ParentDictId { get; set; }
    public string? Name { get; set; }
    public string? RegNumber { get; set; }
    public string? SectionName { get; set; }
    public int Nnumber { get; set; }

    public string ToValueString()
    {
        string parentDictId = ParentDictId != null ? $"'{ParentDictId}'" : "NULL";
        string safeName = Name?.Replace("'", "''") ?? "нет имени";
        string safeSectionName = SectionName?.Replace("'", "''") ?? "";
        return $"('{DictId}', {parentDictId}, '{RegNumber}', '{safeName}', {Nnumber}, NULL, '{safeSectionName}')";
    }
}
