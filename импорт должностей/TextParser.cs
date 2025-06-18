using System.Text.RegularExpressions;

namespace импорт_должностей;

public interface IParser
{
    bool CanParse(string text);

    (string Value, string Name, Node? Node) Parse(string text);
}

public class ПарсерРаздела : IParser
{
    public bool CanParse(string text) => text.StartsWith("Раздел");

    public (string Value, string Name, Node? Node) Parse(string text)
    {
        var value = NormalizeSectionValue(text);
        return (value, "Раздел", new Раздел(value, string.Empty));
    }

    /// <summary>
    /// Нормализует значение раздела, форматируя номер в двухзначный формат.
    /// </summary>
    /// <param name="value">Текст для нормализации.</param>
    /// <returns>Нормализованное значение раздела или "error", если парсинг невозможен.</returns>
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
        return int.TryParse(value, out int number) ? number.ToString("D2") : "error";
    }
}

public class ПарсерИмениРаздела : IParser
{
    public bool CanParse(string text) => text.StartsWith("Перечень");

    public (string Value, string Name, Node? Node) Parse(string text) =>
        (string.Empty, text, null);
}

public class ПарсерПодраздела : IParser
{
    public bool CanParse(string text) => text.StartsWith("Подраздел");

    public (string Value, string Name, Node? Node) Parse(string text)
    {
        // Регулярное выражение: "Подраздел" и номер (например, "Подраздел 1")
        var match = Regex.Match(text, @"Подраздел (\d+)");
        if (!match.Success) throw new InvalidOperationException("Invalid subsection format");
        return (match.Groups[1].Value, text, new Подраздел(match.Groups[1].Value, text));
    }
}

public class ПарсерГлавы : IParser
{
    public bool CanParse(string text) => text.StartsWith("Глава");

    public (string Value, string Name, Node? Node) Parse(string text)
    {
        // Регулярное выражение: "Глава" и номер (например, "Глава 1")
        var match = Regex.Match(text, @"Глава (\d+)");
        if (!match.Success) throw new InvalidOperationException("Invalid chapter format");
        return (match.Groups[1].Value, text, new Глава(match.Groups[1].Value, text));
    }
}

public class ПасерКатегории : IParser
{
    public bool CanParse(string text) => text.Contains("Должность категории") || text.Contains("Должности категории");

    public (string Value, string Name, Node? Node) Parse(string text)
    {
        // Регулярное выражение: номер, любые символы, слово "категории" и все остальное (включая возможные кавычки)
        var match = Regex.Match(text, @"(\d+)\.\s*.*категории\s*""?.*""?");
        if (!match.Success)
        {
            Console.WriteLine($"Failed to parse: {text}");
            throw new InvalidOperationException("Invalid category format");
        }
        return (match.Groups[1].Value, text, new Категория(match.Groups[1].Value, text));
    }
}

public class ПарсерГруппы : IParser
{
    public bool CanParse(string text) => text.Contains("группа должностей");

    public (string Value, string Name, Node? Node) Parse(string text) =>
        (GetGroupValue(text), text, new Группа(GetGroupValue(text), text));

    /// <summary>
    /// Определяет значение группы на основе ключевых слов в тексте.
    /// </summary>
    /// <param name="text">Текст для анализа.</param>
    /// <returns>Значение группы (например, "1" для "Высшая").</returns>
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
}

public class TextParser
{
    private readonly IParser[] _parsers =
    {
        new ПарсерРаздела(),
        new ПарсерИмениРаздела(),
        new ПарсерПодраздела(),
        new ПарсерГлавы(),
        new ПасерКатегории(),
        new ПарсерГруппы()
    };

    /// <summary>
    /// Возвращает подходящий парсер для указанного текста.
    /// </summary>
    /// <param name="text">Текст для проверки.</param>
    /// <returns>Подходящий парсер или null, если парсер не найден.</returns>
    public IParser? GetParser(string text) =>
        _parsers.FirstOrDefault(p => p.CanParse(text));
}