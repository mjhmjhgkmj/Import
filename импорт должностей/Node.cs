namespace импорт_должностей;

/// <summary>
/// Абстрактный базовый класс для узлов иерархии. От него наследуются классы разделов, подразделов, глав, категорий, групп.
/// Определяет общие свойства и методы для всех узлов.
/// </summary>
public abstract class Node
{
    public string DictId { get; } = Guid.NewGuid().ToString();
    public string? ParentDictId { get; set; }
    public Node? Parent { get; set; }
    public List<Node> Children { get; } = [];
    public string Value { get; }
    public string Name { get; set; }
    public abstract int Level { get; }

    protected Node(string value, string name)
    {
        Value = value;
        Name = name;
    }

    /// <summary>
    /// Добавляет дочерний узел, если это разрешено.
    /// </summary>
    /// <param name="child">Дочерний узел для добавления.</param>
    /// <exception cref="InvalidOperationException">Выбрасывается, если добавление невозможно.</exception>
    public void AddChild(Node child)
    {
        if (CanAddChild(child))
        {
            child.Parent = this;
            child.ParentDictId = DictId;
            Children.Add(child);
        }
        else
        {
            throw new InvalidOperationException($"Cannot add {child.GetType().Name} to {GetType().Name}");
        }
    }

    /// <summary>
    /// Проверяет, может ли текущий узел добавить указанный дочерний узел.
    /// </summary>
    /// <param name="child">Дочерний узел для проверки.</param>
    /// <returns>True, если добавление возможно, иначе false.</returns>
    public abstract bool CanAddChild(Node child);

    /// <summary>
    /// Возвращает часть регистрационного номера, соответствующую данному узлу.
    /// </summary>
    /// <returns>Часть регистрационного номера.</returns>

    /// <summary>
    /// Проверяет, является ли узел пустым (не имеет дочерних узлов).
    /// </summary>
    /// <returns>True, если узел пуст, иначе false.</returns>
    public bool IsEmpty() => Children.Count == 0;
}

public class Раздел : Node
{
    public Раздел(string value, string name) : base(value, name) { }
    public override int Level => 1;

    /// <summary>
    /// Проверяет, может ли раздел добавить указанный дочерний узел.
    /// Разрешены подразделы, главы и категории.
    /// </summary>
    /// <param name="child">Дочерний узел для проверки.</param>
    /// <returns>True, если добавление возможно, иначе false.</returns>
    public override bool CanAddChild(Node child) =>
        child is Подраздел || child is Глава || child is Категория;

    /// <summary>
    /// Возвращает часть регистрационного номера для раздела.
    /// </summary>
    /// <returns>Значение раздела.</returns>
}

public class Подраздел : Node
{
    public Подраздел(string value, string name) : base(value, name) { }
    public override int Level => 2;

    /// <summary>
    /// Проверяет, может ли подраздел добавить указанный дочерний узел.
    /// Разрешены главы и категории.
    /// </summary>
    /// <param name="child">Дочерний узел для проверки.</param>
    /// <returns>True, если добавление возможно, иначе false.</returns>
    public override bool CanAddChild(Node child) =>
        child is Глава || child is Категория;

    /// <summary>
    /// Возвращает часть регистрационного номера для подраздела.
    /// </summary>
    /// <returns>Значение подраздела.</returns>
}

public class Глава : Node
{
    public Глава(string value, string name) : base(value, name) { }
    public override int Level => 3;

    /// <summary>
    /// Проверяет, может ли глава добавить указанный дочерний узел.
    /// Разрешены только категории.
    /// </summary>
    /// <param name="child">Дочерний узел для проверки.</param>
    /// <returns>True, если добавление возможно, иначе false.</returns>
    public override bool CanAddChild(Node child) =>
        child is Категория;

    /// <summary>
    /// Возвращает часть регистрационного номера для главы.
    /// </summary>
    /// <returns>Значение главы с префиксом "R".</returns>
}

public class Категория : Node
{
    public Категория(string value, string name) : base(value, name) { }
    public override int Level => 4;

    /// <summary>
    /// Проверяет, может ли категория добавить указанный дочерний узел.
    /// Разрешены только группы.
    /// </summary>
    /// <param name="child">Дочерний узел для проверки.</param>
    /// <returns>True, если добавление возможно, иначе false.</returns>
    public override bool CanAddChild(Node child) =>
        child is Группа;

}

public class Группа : Node
{
    public Группа(string value, string name) : base(value, name) { }
    public override int Level => 5;

    /// <summary>
    /// Проверяет, может ли группа добавить указанный дочерний узел.
    /// Группа является листовым узлом, добавление дочерних узлов запрещено.
    /// </summary>
    /// <param name="child">Дочерний узел для проверки.</param>
    /// <returns>Всегда false, так как группа — листовой узел.</returns>
    public override bool CanAddChild(Node child) => false;
}