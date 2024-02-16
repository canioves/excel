using System.Text;
using excelTask3.Exceptions;
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
    bool keep = true;

    switch (key)
    {
        case "1":
            while (keep)
            {
                Console.Write("Введите название товара: ");
                string productName = Console.ReadLine();
                keep = app.GetAllProductInfo(productName);
            }
            break;
        case "2":
            Console.WriteLine("Здесь будет золотой клиент");
            break;
        case "3":
            app.ChangeOrgContactName("ООО Тест", "Измененный Тест Тестович");
            break;
        case "q":
            Console.WriteLine("Выход...");
            appRunning = false;
            break;
    }
    if (!appRunning) break;
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

    public bool GetAllProductInfo(string productName)
    {
        bool startOver = false;
        int clientId = 0;
        Console.WriteLine($"Информация по товару {productName}:\n");
        try
        {
            var product = productService.GetProductByName(productName) ?? throw new ProductException("Товар не найден");
            var clients = clientService.ClientsList;
            var requests = requestService.RequestsList;
            var targetRequests = requests.Where(x => x.ProductId == product.Id);

            if (targetRequests.Count() == 0) throw new RequestException("У данного товара отсутствуют заказы.");

            foreach (var req in targetRequests)
            {
                clientId = req.ClientId;
                var reqClient = clients.SingleOrDefault(x => x.Id == req.ClientId) ?? throw new ClientException("Клиент не найден!", clientId);
                string org = reqClient.Organization;
                int amount = req.Amount;
                double totalPrice = product.Price * amount;
                string shortDate = req.Date.ToShortDateString();

                PrintAllProductInfo(org, amount, totalPrice, shortDate);
                startOver = false;
            }
        }
        catch (ProductException pEx)
        {
            Console.WriteLine(pEx.Message);
        }
        catch (ClientException cEx)
        {
            Console.Write(cEx.Message);
            Console.WriteLine($" ID клиента: {cEx.Value}");
        }
        catch (RequestException rEx)
        {
            Console.WriteLine(rEx.Message);
        }
        finally
        {
            Console.WriteLine("Попробовать еще раз? (y/n)");
            var answer = Console.ReadLine();
            startOver = answer == "y";
        }
        return startOver;
    }

    public bool ChangeOrgContactName(string organization, string newFullName)
    {
        var client = clientService.GetClientByOrganization(organization);
        client.FullName = newFullName;
        clientService.UpdateClientInfo(client);
        return true;
    }
    private void PrintAllProductInfo(string organization, int amount, double price, string date)
    {
        Console.WriteLine($"- Организация: {organization}");
        Console.WriteLine($"- Количество заказанного товара: {amount}");
        Console.WriteLine($"- Итоговая цена: {price}");
        Console.WriteLine($"- Дата заказа: {date}");
    }

    public void PrintMenu()
    {
        Console.WriteLine("1. Получить информацию о клиентах по названию товара.");
        Console.WriteLine("2. Определить \"Золотого\" клиента.");
        Console.WriteLine("3. Изменить ФИО контактного лица организации.\n");
        Console.WriteLine("Нажмите q, чтобы выйти...");
    }
}