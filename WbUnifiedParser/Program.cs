using System.Net;
using System.Text;
using System.Text.Json;
using ClosedXML.Excel;

namespace WbUnifiedParser;

internal static class Program
{
    public static async Task<int> Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        return await RunAsync(args);
    }

    public static async Task<int> RunAsync(string[] args)
    {
        Console.WriteLine("=== WB JSON -> Excel: товары, продавцы, карточки ===");

        var skipEnrichment = args.Any(IsOfflineFlag);
        var positionalArgs = args.Where(arg => !IsOfflineFlag(arg)).ToArray();

        var inputPath = positionalArgs.Length > 0
            ? positionalArgs[0]
            : Ask("Введите путь к JSON/TXT файлу или папке: ");

        inputPath = NormalizePath(inputPath);

        if (string.IsNullOrWhiteSpace(inputPath))
        {
            Console.WriteLine("Путь не указан.");
            return 1;
        }

        if (!File.Exists(inputPath) && !Directory.Exists(inputPath))
        {
            Console.WriteLine("Файл или папка не найдены.");
            return 1;
        }

        var outputPath = positionalArgs.Length > 1
            ? NormalizePath(positionalArgs[1])
            : BuildDefaultOutputPath(inputPath);

        var sourceFiles = GetSourceFiles(inputPath);
        if (sourceFiles.Count == 0)
        {
            Console.WriteLine("Не найдено JSON/TXT файлов для обработки.");
            return 1;
        }

        var rows = new List<WbRow>();
        foreach (var file in sourceFiles)
        {
            Console.WriteLine($"Читаю: {file}");
            var json = await File.ReadAllTextAsync(file, Encoding.UTF8);
            rows.AddRange(WbJsonParser.ExtractRows(json, Path.GetFileName(file)));
        }

        rows = rows
            .Where(row => row.ProductId > 0 && row.SellerId > 0)
            .GroupBy(row => row.ProductId)
            .Select(group => group.First())
            .OrderBy(row => row.ProductId)
            .ToList();

        if (rows.Count == 0)
        {
            Console.WriteLine("В файлах не найдено товаров.");
            return 1;
        }

        Console.WriteLine($"Найдено уникальных товаров: {rows.Count}");

        if (skipEnrichment)
            Console.WriteLine("Сетевое обогащение пропущено: указан флаг --offline.");
        else
            await WbEnricher.EnrichAsync(rows);

        ExcelWriter.Save(rows, outputPath);

        Console.WriteLine($"Готово! Создан файл: {outputPath}");
        Console.WriteLine($"Всего уникальных товаров: {rows.Count}");
        return 0;
    }

    private static string Ask(string text)
    {
        Console.Write(text);
        return Console.ReadLine() ?? "";
    }

    private static string NormalizePath(string path)
    {
        return (path ?? "").Trim().Trim('"');
    }

    private static string BuildDefaultOutputPath(string inputPath)
    {
        var directory = Directory.Exists(inputPath)
            ? inputPath
            : Path.GetDirectoryName(inputPath);

        if (string.IsNullOrWhiteSpace(directory))
            directory = Environment.CurrentDirectory;

        return Path.Combine(directory, "wb_unified.xlsx");
    }

    private static bool IsOfflineFlag(string arg)
    {
        return arg.Equals("--offline", StringComparison.OrdinalIgnoreCase)
            || arg.Equals("--no-enrich", StringComparison.OrdinalIgnoreCase);
    }

    private static List<string> GetSourceFiles(string inputPath)
    {
        if (File.Exists(inputPath))
            return IsJsonOrText(inputPath) ? [inputPath] : [];

        return Directory.EnumerateFiles(inputPath, "*.*", SearchOption.TopDirectoryOnly)
            .Where(IsJsonOrText)
            .OrderBy(path => path)
            .ToList();
    }

    private static bool IsJsonOrText(string path)
    {
        var extension = Path.GetExtension(path);
        return extension.Equals(".json", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".txt", StringComparison.OrdinalIgnoreCase);
    }
}

internal sealed class WbProduct
{
    public long Id { get; set; }
    public long SupplierId { get; set; }
    public string? Supplier { get; set; }
    public string? Name { get; set; }
    public string? Entity { get; set; }
    public int SubjectId { get; set; }
    public int SubjectParentId { get; set; }
}

internal sealed class WbRoot
{
    public List<WbProduct>? Products { get; set; }
}

internal sealed class WbRow
{
    public string SourceFile { get; init; } = "";
    public long ProductId { get; init; }
    public string ProductName { get; init; } = "";
    public string Category { get; init; } = "";
    public int CategoryId { get; init; }
    public string ParentCategory { get; init; } = "";
    public int ParentCategoryId { get; init; }
    public string Entity { get; init; } = "";
    public long SellerId { get; init; }
    public string SellerName { get; init; } = "";
    public string SellerUrl => $"https://www.wildberries.ru/seller/{SellerId}";
    public string ProductUrl => $"https://www.wildberries.ru/catalog/{ProductId}/detail.aspx";

    public string SellerFullName { get; set; } = "";
    public string Inn { get; set; } = "";
    public string RegistrationNumber { get; set; } = "";
    public string MainInfo { get; set; } = "";
    public string ExtraInfo { get; set; } = "";
    public string Description { get; set; } = "";
    public string Documents { get; set; } = "";
}

internal static class WbJsonParser
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public static List<WbRow> ExtractRows(string json, string sourceFile)
    {
        var result = new List<WbRow>();

        try
        {
            var cleanJson = CleanJson(json);
            var root = JsonSerializer.Deserialize<WbRoot>(cleanJson, SerializerOptions);

            if (root?.Products == null || root.Products.Count == 0)
                return result;

            foreach (var product in root.Products)
            {
                if (product.Id == 0 || product.SupplierId == 0)
                    continue;

                var category = WbCategories.Map.TryGetValue(product.SubjectId, out var info)
                    ? info
                    : ("", "");

                result.Add(new WbRow
                {
                    SourceFile = sourceFile,
                    ProductId = product.Id,
                    ProductName = product.Name ?? "",
                    Category = category.Item1,
                    CategoryId = product.SubjectId,
                    ParentCategory = category.Item2,
                    ParentCategoryId = product.SubjectParentId,
                    Entity = product.Entity ?? "",
                    SellerId = product.SupplierId,
                    SellerName = product.Supplier ?? ""
                });
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка в файле {sourceFile}: {ex.Message}");
        }

        return result;
    }

    private static string CleanJson(string json)
    {
        json = json.Trim('\uFEFF', '\u200B', '\u0000', '\u001F');

        if (!json.Contains("<html", StringComparison.OrdinalIgnoreCase))
            return json;

        var start = json.IndexOf('{');
        var end = json.LastIndexOf('}');
        return start >= 0 && end > start
            ? json[start..(end + 1)]
            : json;
    }
}

internal static class WbEnricher
{
    public static async Task EnrichAsync(List<WbRow> rows)
    {
        using var http = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(5)
        };

        http.DefaultRequestHeaders.UserAgent.ParseAdd(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");

        var sellerCache = new Dictionary<long, SellerDetails>();
        var productCache = new Dictionary<long, ProductDetails>();

        for (var i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            Console.WriteLine($"[{i + 1}/{rows.Count}] Продавец {row.SellerId}, товар {row.ProductId}");

            if (!sellerCache.TryGetValue(row.SellerId, out var seller))
            {
                seller = await GetSellerDetailsAsync(http, row.SellerId);
                sellerCache[row.SellerId] = seller;
            }

            row.SellerFullName = seller.FullName;
            row.Inn = seller.Inn;
            row.RegistrationNumber = seller.RegistrationNumber;

            if (!productCache.TryGetValue(row.ProductId, out var product))
            {
                product = await GetProductDetailsAsync(http, row.ProductId);
                productCache[row.ProductId] = product;
            }

            row.MainInfo = product.MainInfo;
            row.ExtraInfo = product.ExtraInfo;
            row.Description = product.Description;
            row.Documents = product.Documents;
        }
    }

    private static async Task<SellerDetails> GetSellerDetailsAsync(HttpClient http, long sellerId)
    {
        var url = $"https://static-basket-01.wbbasket.ru/vol0/data/supplier-by-id/{sellerId}.json";

        try
        {
            using var response = await http.GetAsync(url);
            if (response.StatusCode == HttpStatusCode.NotFound)
                return new SellerDetails("Нет данных (404)", "", "");

            response.EnsureSuccessStatusCode();

            await using var stream = await response.Content.ReadAsStreamAsync();
            using var document = await JsonDocument.ParseAsync(stream);
            var root = document.RootElement;

            var fullName = GetString(root, "supplierFullName");
            var inn = GetString(root, "inn");
            var ogrn = GetString(root, "ogrn");
            var ogrnip = GetString(root, "ogrnip");
            var registrationNumber = !string.IsNullOrWhiteSpace(ogrn)
                ? ogrn
                : !string.IsNullOrWhiteSpace(ogrnip)
                    ? ogrnip
                    : "нет данных";

            return new SellerDetails(fullName, inn, registrationNumber);
        }
        catch (Exception ex)
        {
            return new SellerDetails($"Ошибка: {ex.Message}", "", "");
        }
    }

    private static async Task<ProductDetails> GetProductDetailsAsync(HttpClient http, long productId)
    {
        var vol = productId / 100000;
        var part = productId / 1000;
        var candidates = GetBasketCandidates(vol);

        foreach (var basket in candidates)
        {
            var url =
                $"https://basket-{basket:D2}.wbbasket.ru/vol{vol}/part{part}/{productId}/info/ru/card.json";

            try
            {
                using var response = await http.GetAsync(url);
                if (!response.IsSuccessStatusCode)
                    continue;

                await using var stream = await response.Content.ReadAsStreamAsync();
                using var document = await JsonDocument.ParseAsync(stream);
                return ParseProductDetails(document.RootElement);
            }
            catch
            {
                // Если конкретная корзина не отвечает, пробуем следующую.
            }
        }

        return new ProductDetails("", "", "Описание не найдено", "Отсутствует");
    }

    private static ProductDetails ParseProductDetails(JsonElement root)
    {
        var description = GetString(root, "description").Trim();
        var mainInfo = new StringBuilder();
        var extraInfo = new StringBuilder();
        var documents = new StringBuilder();

        if (root.TryGetProperty("grouped_options", out var groups)
            && groups.ValueKind == JsonValueKind.Array)
        {
            foreach (var group in groups.EnumerateArray())
            {
                var groupName = GetString(group, "group_name");
                if (!group.TryGetProperty("options", out var options)
                    || options.ValueKind != JsonValueKind.Array)
                {
                    continue;
                }

                var target = groupName switch
                {
                    "Основная информация" => mainInfo,
                    "Дополнительная информация" => extraInfo,
                    "Документы" => documents,
                    _ => null
                };

                if (target == null)
                    continue;

                foreach (var option in options.EnumerateArray())
                {
                    var name = GetString(option, "name");
                    var value = GetValueAsText(option, "value");

                    if (!string.IsNullOrWhiteSpace(name) || !string.IsNullOrWhiteSpace(value))
                        target.AppendLine($"{name}: {value}".Trim());
                }
            }
        }

        var documentText = documents.ToString().Trim();
        return new ProductDetails(
            mainInfo.ToString().Trim(),
            extraInfo.ToString().Trim(),
            description,
            string.IsNullOrWhiteSpace(documentText) ? "Отсутствует" : documentText);
    }

    private static IEnumerable<int> GetBasketCandidates(long vol)
    {
        var preferred = ResolveBasketNumber(vol);
        return new[] { preferred }
            .Concat(Enumerable.Range(1, 100))
            .Distinct();
    }

    private static int ResolveBasketNumber(long vol)
    {
        return vol switch
        {
            <= 143 => 1,
            <= 287 => 2,
            <= 431 => 3,
            <= 719 => 4,
            <= 1007 => 5,
            <= 1061 => 6,
            <= 1115 => 7,
            <= 1169 => 8,
            <= 1313 => 9,
            <= 1601 => 10,
            <= 1655 => 11,
            <= 1919 => 12,
            <= 2045 => 13,
            <= 2189 => 14,
            <= 2405 => 15,
            <= 2621 => 16,
            <= 2837 => 17,
            <= 3053 => 18,
            <= 3269 => 19,
            <= 3485 => 20,
            <= 3701 => 21,
            <= 3917 => 22,
            <= 4133 => 23,
            <= 4349 => 24,
            <= 4565 => 25,
            <= 4781 => 26,
            <= 4997 => 27,
            <= 5213 => 28,
            <= 5429 => 29,
            _ => 30
        };
    }

    private static string GetString(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value)
            || value.ValueKind == JsonValueKind.Null
            || value.ValueKind == JsonValueKind.Undefined)
        {
            return "";
        }

        return value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? ""
            : value.ToString();
    }

    private static string GetValueAsText(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value)
            || value.ValueKind == JsonValueKind.Null
            || value.ValueKind == JsonValueKind.Undefined)
        {
            return "";
        }

        if (value.ValueKind == JsonValueKind.Array)
        {
            return string.Join(", ", value.EnumerateArray().Select(GetElementText));
        }

        return GetElementText(value);
    }

    private static string GetElementText(JsonElement value)
    {
        return value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? ""
            : value.ToString();
    }
}

internal sealed record SellerDetails(
    string FullName,
    string Inn,
    string RegistrationNumber);

internal sealed record ProductDetails(
    string MainInfo,
    string ExtraInfo,
    string Description,
    string Documents);

internal static class ExcelWriter
{
    private static readonly string[] Headers =
    [
        "Файл",
        "ID товара",
        "Название товара",
        "Категория",
        "ID категории",
        "Родительская категория",
        "ID родительской категории",
        "Тип товара (entity)",
        "ID продавца",
        "Название продавца на WB",
        "Ссылка на продавца",
        "Ссылка на товар",
        "Полное имя продавца",
        "ИНН",
        "ОГРН/ОГРНИП",
        "Основная информация",
        "Дополнительная информация",
        "Описание",
        "Документы"
    ];

    public static void Save(IReadOnlyList<WbRow> rows, string outputPath)
    {
        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(directory))
            Directory.CreateDirectory(directory);

        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("WB");

        for (var col = 0; col < Headers.Length; col++)
            worksheet.Cell(1, col + 1).Value = Headers[col];

        var header = worksheet.Range(1, 1, 1, Headers.Length);
        header.Style.Font.Bold = true;
        header.Style.Fill.BackgroundColor = XLColor.FromHtml("#D9EAD3");

        for (var i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            var excelRow = i + 2;

            worksheet.Cell(excelRow, 1).Value = row.SourceFile;
            worksheet.Cell(excelRow, 2).Value = row.ProductId;
            worksheet.Cell(excelRow, 3).Value = row.ProductName;
            worksheet.Cell(excelRow, 4).Value = row.Category;
            worksheet.Cell(excelRow, 5).Value = row.CategoryId;
            worksheet.Cell(excelRow, 6).Value = row.ParentCategory;
            worksheet.Cell(excelRow, 7).Value = row.ParentCategoryId;
            worksheet.Cell(excelRow, 8).Value = row.Entity;
            worksheet.Cell(excelRow, 9).Value = row.SellerId;
            worksheet.Cell(excelRow, 10).Value = row.SellerName;
            worksheet.Cell(excelRow, 11).Value = row.SellerUrl;
            worksheet.Cell(excelRow, 12).Value = row.ProductUrl;
            worksheet.Cell(excelRow, 13).Value = row.SellerFullName;
            worksheet.Cell(excelRow, 14).Value = row.Inn;
            worksheet.Cell(excelRow, 15).Value = row.RegistrationNumber;
            worksheet.Cell(excelRow, 16).Value = row.MainInfo;
            worksheet.Cell(excelRow, 17).Value = row.ExtraInfo;
            worksheet.Cell(excelRow, 18).Value = row.Description;
            worksheet.Cell(excelRow, 19).Value = row.Documents;

            worksheet.Cell(excelRow, 11).SetHyperlink(new XLHyperlink(row.SellerUrl));
            worksheet.Cell(excelRow, 12).SetHyperlink(new XLHyperlink(row.ProductUrl));
        }

        worksheet.SheetView.FreezeRows(1);
        worksheet.RangeUsed()?.SetAutoFilter();
        worksheet.Columns().AdjustToContents();
        worksheet.Columns(16, 19).Width = 45;
        worksheet.Columns(16, 19).Style.Alignment.WrapText = true;
        worksheet.Rows().Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;

        workbook.SaveAs(outputPath);
    }
}

internal static class WbCategories
{
    public static readonly Dictionary<int, (string Name, string Parent)> Map = new()
    {
        { 6496, ("Виниры накладные", "Здоровье") },
        { 1862, ("Отбеливающее средство для зубов", "Красота") },
        { 3809, ("Воск для брекетов", "Здоровье") },
        { 387,  ("Прорезыватель", "Товары для малышей") },
        { 1422, ("Наклейка для техники", "Аксессуары") },
        { 2786, ("Зубной порошок", "Красота") },
        { 2826, ("Уход за зубными протезами", "Красота") },
        { 4011, ("Блестки декоративные", "Рукоделие") },
        { 3050, ("Стоматологический набор", "Здоровье") },
        { 2930, ("Грилзы", "Красота") },
        { 438,  ("Зубная паста", "Красота") },
        { 3017, ("Пломба", "Канцелярские товары") },
        { 2824, ("Грим", "Для праздника") },
        { 403,  ("Косметический набор для ухода", "Красота") },
        { 3780, ("Средство ухода за стомой", "Здоровье") },
        { 6350, ("Гель для зубов/десен", "Здоровье") },
        { 592,  ("Стол туристический", "Спортивный товар") },
        { 3771, ("Капа стоматологическая", "Здоровье") },
        { 4585, ("Наконечник стоматологический", "Здоровье") },
        { 4942, ("Проволка пломбировочная", "Канцелярские товары") },
        { 4936, ("Пломба-наклейка", "Канцелярские товары") },
        { 6461, ("Оборудование зуботехническое", "Здоровье") },
        { 4052, ("Аппарат для отбеливания зубов", "Бытовая техника") },
        { 815,  ("Лупа", "Канцелярские товары") },
        { 544,  ("Кейс для камер", "Фото и Видеотехника") },
        { 6375, ("Витаминно-минеральный препарат", "Фарма") },
        { 2926, ("Косметический актив", "Красота") },
        { 1095, ("Аминокислота", "Спортивное питание и косметика") },
        { 1524, ("БАД", "Здоровье") },
        { 4378, ("Комплексная пищевая добавка", "Здоровье") },
        { 1938, ("Оздоровительная косметика", "Здоровье") },
        { 5520, ("Специализированное питание", "Здоровье") },
        { 3470, ("Антиоксидант", "Спортивное питание и косметика") },
        { 3043, ("Средство профилактики", "Здоровье") },
        { 4336, ("Экстракт пищевой растительный", "Продукты") },
        { 1109, ("Жиросжигатель", "Спортивное питание и косметика") },
        { 1103, ("Добавка для суставов и связок", "Спортивное питание и косметика") },
        { 1508, ("Травяной сбор", "Продукты") }
    };
}
