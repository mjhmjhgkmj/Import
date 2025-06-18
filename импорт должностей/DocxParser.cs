// DocxParser.cs
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
namespace импорт_должностей;

public class DocxParser
{
    private readonly DocxReader _docxReader;
    private readonly TableProcessor _tableProcessor;
    private readonly SqlGenerator _sqlGenerator;

    public DocxParser()
    {
        _docxReader = new DocxReader();
        _sqlGenerator = new SqlGenerator();
        _tableProcessor = new TableProcessor(_sqlGenerator);
    }

    public string ParseAndGenerateSql(WordprocessingDocument doc)
    {
        var validTables = _docxReader.GetValidTables(doc);
        if (validTables == null) throw new Exception("не найдена таблица");
        foreach (var table in validTables)
        {
            _tableProcessor.ProcessTable(table);
        }

        return SqlGenerator.AddTableCreationSql + _sqlGenerator.GenerateInsertStatements();
    }
}
