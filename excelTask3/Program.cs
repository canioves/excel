using System.Text;
using excelTask3.Service;
using Models;

Console.InputEncoding = Encoding.UTF8;

App app = new App();

bool appRunning = true;
Console.WriteLine("Введите путь до Excel файла: ");

while (app.repeatInit)
{
    string path = @"Data\dataTable.xlsx";
    // string path = Console.ReadLine();
    app.Init(path);
}

while (appRunning)
{
    // Console.Clear();
    app.PrintMenu();
    string key = Console.ReadLine();
    
    switch (key)
    {
        case "1":
            Console.Write("Введите название товара: ");
            string productName = Console.ReadLine();
            app.PrintProduct(productName);
            break;
        case "2":
            Console.WriteLine("Здесь будет золотой клиент");
            break;
        case "3":
            Console.WriteLine("Здесь будет измененная информация о клиенте");
            break;
        case "q":
            Console.WriteLine("Выход...");
            appRunning = false;
            break;
    }
    if (!appRunning) break;
    Console.WriteLine("Чтобы продолжить, нажмите любую клавишу...");
    Console.ReadLine();
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
        Console.WriteLine("3. Изменить ФИО контактного лица организации.\n");
        Console.WriteLine("Нажмите q, чтобы выйти...");
    }
    public void PrintProduct(string name)
    {
        
        Product product = productService.GetProductByName(name);
        Console.WriteLine($"Айди: {product.Id}");
        Console.WriteLine($"Название: {product.Name}");
        Console.WriteLine($"Единица измерения: {product.UnitName}");
        Console.WriteLine($"Цена: {product.Price}\n");
    }
}