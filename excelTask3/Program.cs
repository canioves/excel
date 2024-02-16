using excelTask3.Service;
using Models;

App app = new App();

bool appRunning = true;
Console.WriteLine("Введите путь до Excel файла: ");

while (app.repeatInit)
{
    // string path = @"Data\dataTable.xlsx";
    string path = Console.ReadLine();
    app.Init(path);
}

while (appRunning)
{
    app.PrintMenu();
    string key = Console.ReadLine();
    switch (key)
    {
        case "1":
            Console.Clear();
            string productName = Console.ReadLine();
            // app.PrintProductId(productName);
            break;
    }
}

// app.PrintAllProducts();
// start.PrintProductId("Йогурт");

public class App
{
    private ProductService productService;
    private ClientService clientService;
    private RequestService requestService;
    private ExcelProcess excelProcess;
    public bool repeatInit = true;
    public void Init(string path)
    {
        excelProcess = new ExcelProcess(path);
        productService = new ProductService(excelProcess);
        clientService = new ClientService(excelProcess);
        requestService = new RequestService(excelProcess);
        repeatInit = excelProcess.repeatInit;
    }
    public void PrintMenu()
    {
        Console.WriteLine("1. Получить информацию о клиентах по названию товара.");
        Console.WriteLine("2. Определить \"Золотого\" клиента.");
        Console.WriteLine("3. Изменить ФИО контактного лица организации.");
    }
    public void PrintAllProducts()
    {
        Console.WriteLine("Все продукты:");
        var lst = productService.GetAllProducts();
        foreach (var product in lst)
        {
            Console.WriteLine($"Айди: {product.Id}");
            Console.WriteLine($"Название: {product.Name}");
            Console.WriteLine($"Единица измерения: {product.UnitName}");
            Console.WriteLine($"Цена: {product.Price}\n");
        }
    }
    public void PrintProduct(string name)
    {
        Console.WriteLine("Введите название товара: ");
        Product product = productService.GetProductByName(name);
        Console.WriteLine($"Айди: {product.Id}");
        Console.WriteLine($"Название: {product.Name}");
        Console.WriteLine($"Единица измерения: {product.UnitName}");
        Console.WriteLine($"Цена: {product.Price}\n");
    }
}