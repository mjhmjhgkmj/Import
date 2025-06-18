public class HierarchyTree
{
    /// <summary>
    /// Корневой узел иерархии. Может быть null, если иерархия еще не инициализирована.
    /// </summary>
    public Node? Root { get; private set; }

    /// <summary>
    /// Стек для отслеживания текущего пути в иерархии. 
    /// Вершина стека — последний добавленный узел.
    /// </summary>
    public readonly Stack<Node> _currentPath = new();

    /// <summary>
    /// Начинает новый раздел, создавая корневой узел и очищая текущий путь.
    /// </summary>
    /// <param name="value">Значение раздела (например, номер или код).</param>
    /// <param name="name">Имя раздела.</param>
    public void НачатьНовыйРаздел(string value, string name)
    {
        // Создаем новый узел раздела и устанавливаем его как корень
        Root = new Раздел(value, name);
        // Очищаем текущий путь, так как начинаем новую иерархию
        _currentPath.Clear();
        // Добавляем корень в стек
        _currentPath.Push(Root);
    }

    /// <summary>
    /// Обновляет имя текущего раздела (корневого узла).
    /// </summary>
    /// <param name="name">Новое имя раздела.</param>
    public void ОбновитьИмяРаздела(string name)
    {
        // Проверяем, что корень является узлом раздела, и обновляем его имя
        if (Root is Раздел section)
            section.Name = name;
    }

    /// <summary>
    /// Добавляет новый узел в иерархию, определяя его родителя и обновляя стек.
    /// </summary>
    /// <param name="node">Новый узел для добавления.</param>
    /// <exception cref="InvalidOperationException">Выбрасывается, если подходящий родитель не найден.</exception>
    public void AddNode(Node node)
    {
        // Находим подходящего родителя для нового узла
        var parent = FindSuitableParent(node);
        if (parent == null)
            throw new InvalidOperationException($"Cannot find suitable parent for {node.GetType().Name}");

        // Добавляем новый узел как дочерний к найденному родителю
        parent.AddChild(node);
        // Выводим имя узла для отладки
        //Console.WriteLine(node.Name);

        // Обновляем стек: удаляем все узлы с уровнем >= уровня нового узла
        // Это гарантирует, что стек содержит только узлы, которые находятся выше нового узла
        while (_currentPath.Any() && _currentPath.Peek().Level >= node.Level)
            _currentPath.Pop();

        // Добавляем новый узел в стек
        _currentPath.Push(node);
    }

    /// <summary>
    /// Находит подходящего родителя для нового узла на основе его уровня.
    /// </summary>
    /// <param name="node">Новый узел, для которого нужно найти родителя.</param>
    /// <returns>Подходящий родитель или null, если родитель не найден.</returns>
    private Node? FindSuitableParent(Node node)
    {
        // Если стек пуст, проверяем, может ли корень быть родителем
        if (!_currentPath.Any())
        {
            if (Root != null && Root.CanAddChild(node))
                return Root;
            return null;
        }

        // Получаем последний добавленный узел (вершина стека)
        var lastNode = _currentPath.Peek();

        // Если уровень нового узла ниже последнего (глубже в иерархии)
        if (node.Level > lastNode.Level)
        {
            // Проверяем, может ли последний узел быть родителем
            if (lastNode.CanAddChild(node))
                return lastNode;
            return null;
        }

        // Если уровень нового узла выше или равен последнему,
        // ищем родителя среди узлов в стеке
        // Копируем стек для итерации, чтобы не изменять оригинальный
        var pathCopy = new Stack<Node>(_currentPath.Reverse());
        Node? parent = null;

        // Проходим по копии стека, ищем узел с уровнем ниже уровня нового узла
        while (pathCopy.Any())
        {
            var current = pathCopy.Pop();
            // Если текущий узел имеет меньший уровень и может быть родителем
            if (current.Level < node.Level && current.CanAddChild(node))
            {
                parent = current;
                break;
            }
        }

        return parent;
    }

    /// <summary>
    /// Возвращает текущий узел (последний добавленный) из стека.
    /// </summary>
    /// <returns>Текущий узел или null, если стек пуст.</returns>
    public Node? GetCurrentNode() => _currentPath.TryPeek(out var node) ? node : null;

    /// <summary>
    /// Возвращает корневой узел иерархии (раздел).
    /// </summary>
    /// <returns>Корневой узел или null, если иерархия не инициализирована.</returns>
    public Node? GetSectionNode() => Root;
}