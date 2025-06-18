using DocumentFormat.OpenXml.Packaging;
using System.Text;

namespace импорт_должностей_жилсубсидий;

public class Program
{
    public static void Main(string[] args)
    {
        string docxFilePath = args.Length > 0 ? args[0] : "test.docx";
        string sqlFilePath = Path.ChangeExtension(docxFilePath, ".sql");

        if (!File.Exists(docxFilePath))
        {
            Console.WriteLine($"Файл '{docxFilePath}' не найден.");
            return;
        }

        try
        {
            string sqlContent = ConvertDocxToSqlInsert(docxFilePath);
            File.WriteAllText(sqlFilePath, sqlContent, Encoding.UTF8);
            Console.WriteLine($"Файл успешно преобразован и сохранен в {sqlFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при обработке файла: {ex.Message}");
        }
    }

    public static string ConvertDocxToSqlInsert(string filePath)
    {
        using WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false);
        var parser = new DocxParser();
        return parser.ParseAndGenerateSql(doc);
    }
}