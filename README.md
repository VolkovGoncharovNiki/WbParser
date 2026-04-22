# WbUnifiedParser

Консольная программа для обработки JSON/TXT ответов поиска Wildberries и выгрузки данных в один Excel-файл.

## Что делает программа

- Читает один JSON/TXT файл или папку с JSON/TXT файлами.
- Находит товары в массиве `products`.
- Формирует ссылки на товар и продавца Wildberries.
- Удаляет дубли по ID товара.
- Дополняет данные продавца: полное имя, ИНН, ОГРН/ОГРНИП.
- Дополняет данные карточки товара: основная информация, дополнительная информация, описание, документы.
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
- Интернет-соединение для обогащения данных продавцов и карточек товаров.

Зависимости устанавливаются через NuGet:

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

Если нужно только проверить парсинг JSON и создание Excel, можно отключить запросы к Wildberries:

```powershell
dotnet run --project "C:\Users\V7918\source\repos\WbUnifiedParser\WbUnifiedParser\WbUnifiedParser.csproj" -- "C:\input\search.txt" "C:\output\result.xlsx" --offline
```

Также поддерживается флаг:

```text
--no-enrich
```

## Формат входных данных

Программа ожидает JSON с массивом товаров `products`, например ответ поиска Wildberries:

```json
{
  "metadata": {},
  "products": [
    {
      "id": 215427638,
      "supplierId": 93600,
      "supplier": "Название продавца",
      "name": "Название товара",
      "entity": "тип товара",
      "subjectId": 7549,
      "subjectParentId": 3109
    }
  ]
}
```

## Колонки в Excel

Итоговый файл содержит:

- файл-источник;
- ID товара;
- название товара;
- категорию и ID категории;
- родительскую категорию и ID родительской категории;
- тип товара `entity`;
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

запуск может быть таким:

```powershell
dotnet run --project "C:\Users\V7918\source\repos\WbUnifiedParser\WbUnifiedParser\WbUnifiedParser.csproj" -- "C:\Users\V7918\Downloads\search (43).txt" "C:\Users\V7918\Downloads\search-43-result.xlsx"
```

## Примечания

- Если категория неизвестна ручному словарю, текстовое название категории может быть пустым, но `subjectId` и `subjectParentId` сохраняются в Excel.
- Сетевые данные Wildberries могут быть недоступны для отдельных товаров или продавцов. В этом случае программа продолжает обработку остальных строк.
- Повторяющиеся продавцы и товары кэшируются во время запуска, чтобы не делать одинаковые запросы несколько раз.
