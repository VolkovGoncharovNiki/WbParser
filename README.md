# WbUnifiedParser

Консольная программа для обработки JSON/TXT ответов поиска Wildberries и выгрузки данных в Excel.

## Что делает программа

- Читает один JSON/TXT файл или папку с JSON/TXT файлами.
- Находит товары в массиве `products`.
- Удаляет дубли по артикулу WB.
- Дополняет данные продавца: полное имя, ИНН, ОГРН/ОГРНИП.
- Дополняет данные карточки товара из `card.json`.
- Выгружает отдельные поля из блока карточки товара:
  - бренд;
  - рейтинг;
  - количество отзывов;
  - количество вопросов, если это поле есть во входном JSON;
  - состав;
  - количество упаковок;
  - количество капсул/таблеток;
  - вкус;
  - назначение добавки;
  - страна производства.
- Сохраняет результат в `.xlsx`.

## Где находится проект

Решение Visual Studio:

```text
C:\Users\V7918\source\repos\WbUnifiedParser\WbUnifiedParser.sln
```

Основной код:

```text
C:\Users\V7918\source\repos\WbUnifiedParser\WbUnifiedParser\Program.cs
```

## Требования

- .NET 8 SDK или новее.
- Интернет для обогащения карточек товара и продавцов.

NuGet-зависимости:

- `ClosedXML`

## Запуск из PowerShell

Обработка файла или папки:

```powershell
dotnet run --project "C:\Users\V7918\source\repos\WbUnifiedParser\WbUnifiedParser\WbUnifiedParser.csproj" -- "C:\путь\к\файлу-или-папке"
```

По умолчанию Excel создается рядом с входным файлом или внутри входной папки под именем:

```text
wb_unified.xlsx
```

Можно указать путь итогового Excel вручную:

```powershell
dotnet run --project "C:\Users\V7918\source\repos\WbUnifiedParser\WbUnifiedParser\WbUnifiedParser.csproj" -- "C:\input\search.txt" "C:\output\result.xlsx"
```

## Быстрый запуск без сетевых запросов

Если нужно проверить только парсинг JSON и создание Excel:

```powershell
dotnet run --project "C:\Users\V7918\source\repos\WbUnifiedParser\WbUnifiedParser\WbUnifiedParser.csproj" -- "C:\input\search.txt" "C:\output\result.xlsx" --offline
```

Поддерживается и альтернативный флаг:

```text
--no-enrich
```

## Формат входных данных

Программа ожидает JSON с массивом `products`, например:

```json
{
  "metadata": {},
  "products": [
    {
      "brand": "NOW",
      "id": 92028596,
      "supplierId": 304497,
      "supplier": "Название продавца",
      "name": "Кверцетин с бромелайном 120 капсул",
      "entity": "капсулы/таблетки",
      "subjectId": 1524,
      "subjectParentId": 4268,
      "nmReviewRating": 4.9,
      "nmFeedbacks": 245
    }
  ]
}
```

## Колонки в Excel

Итоговый файл содержит:

- файл-источник;
- артикул WB;
- название товара;
- бренд;
- рейтинг;
- количество отзывов;
- количество вопросов;
- категорию и ID категории;
- родительскую категорию и ID родительской категории;
- тип товара `entity`;
- состав;
- количество упаковок;
- количество капсул/таблеток;
- вкус;
- назначение добавки;
- страну производства;
- ID продавца;
- название продавца на WB;
- ссылку на продавца;
- ссылку на товар;
- полное имя продавца;
- ИНН;
- ОГРН/ОГРНИП;
- основную информацию карточки;
- дополнительную информацию карточки;
- описание;
- документы.

## Сборка

```powershell
dotnet build "C:\Users\V7918\source\repos\WbUnifiedParser\WbUnifiedParser.sln"
```

## Пример

Для файла:

```text
C:\Users\V7918\Downloads\search (43).txt
```

запуск:

```powershell
dotnet run --project "C:\Users\V7918\source\repos\WbUnifiedParser\WbUnifiedParser\WbUnifiedParser.csproj" -- "C:\Users\V7918\Downloads\search (43).txt" "C:\Users\V7918\Downloads\search-43-selected-block.xlsx"
```

## Примечания

- Если категория отсутствует в ручном словаре, текстовое название категории может быть пустым, но `subjectId` и `subjectParentId` сохраняются.
- Часть полей карточки зависит от конкретного товара. Если у товара нет параметра в `card.json`, соответствующая колонка останется пустой.
- Количество вопросов заполняется только если оно есть во входном JSON.
- Повторяющиеся продавцы, карточки и рейтинги кэшируются в рамках одного запуска.
