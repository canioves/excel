using System.Text;
using excelTask3.Exceptions;
using excelTask3.Service;

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
    bool keepAlive = true;

    switch (key)
    {
        case "1":
            while (keepAlive)
            {
                Console.Write("Введите название товара: ");
                string productName = Console.ReadLine();
                keepAlive = app.GetAllProductInfo(productName);
            }
            break;
        case "2":
            Console.WriteLine("Здесь будет золотой клиент");
            break;
        case "3":
            while (keepAlive)
            {
                Console.Write("Введите название орагнизации: ");
                string organization = Console.ReadLine();
                Console.Write("Введите новое ФИО контактного лица: ");
                string newName = Console.ReadLine();
                keepAlive = app.ChangeOrgContactName(organization, newName);
            }
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
        bool startOver = true;
        bool waitKey = true;
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
            }
            startOver = false;
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
            if (startOver)
            {
                waitKey = false;
                Console.WriteLine("Попробовать еще раз? (y/n)\n");
                var answer = Console.ReadLine();
                startOver = answer == "y";
            }
        }

        if (waitKey)
        {
            waitKey = false;
            Console.WriteLine("Для продолжения нажмите любую кнопку...");
            Console.ReadLine();
        }

        return startOver;
    }

    public bool ChangeOrgContactName(string organization, string newFullName)
    {
        bool startOver = true;
        bool waitKey = true;
        try
        {
            var client = clientService.GetClientByOrganization(organization)
                ?? throw new ClientException("Не найден клиент с такой организацией!");

            if (newFullName.Length == 0) throw new ArgumentException("Имя не может быть пустым!");

            string oldName = client.FullName;

            client.FullName = newFullName;
            clientService.UpdateClientInfo(client);

            Console.WriteLine($"Изменения имени контактного лица: {oldName} => {newFullName}\n");

            startOver = false;
        }
        catch (ClientException cEx)
        {
            Console.WriteLine(cEx.Message);
        }
        catch (ArgumentException aEx)
        {
            Console.WriteLine(aEx.Message);
        }
        finally
        {
            if (startOver)
            {
                Console.WriteLine("Попробовать еще раз? (y/n)\n");
                var answer = Console.ReadLine();
                startOver = answer == "y";
            }
        }

        if (waitKey)
        {
            waitKey = false;
            Console.WriteLine("Для продолжения нажмите любую кнопку...");
            Console.ReadLine();
        }
        return startOver;
    }

    public bool GetGoldenClient (int year, int month)
    {
        bool startOver = true;
        bool waitKey = true;

        try
        {
            Dictionary<int, int> clientCount = new();
            var clients = clientService.ClientsList;
            var requests = requestService.RequestsList;

            var filterRequests = requests.Where(x => x.Date.Year == year && x.Date.Month == month);
            var filterClients = filterRequests.Select(x => x.ClientId).Distinct();

            int targetId = 0;
            int max = 0;

            foreach (int clientId in filterClients)
            {
                int reqCount = filterRequests.Count(x => x.ClientId == clientId);
                if (reqCount > max)
                {
                    max = reqCount;
                    targetId = clientId;
                }                 
            }

            var targetClient = clientService.GetClientById(targetId);

            Console.WriteLine($"Золотой клиент: {targetClient.Organization}");

            startOver = false;
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
            if (startOver)
            {
                waitKey = false;
                Console.WriteLine("Попробовать еще раз? (y/n)\n");
                var answer = Console.ReadLine();
                startOver = answer == "y";
            }
        }

        if (waitKey)
        {
            waitKey = false;
            Console.WriteLine("Для продолжения нажмите любую кнопку...");
            Console.ReadLine();
        }

        return startOver;
    }

    private void PrintAllProductInfo(string organization, int amount, double price, string date)
    {
        Console.WriteLine($"- Организация: {organization}");
        Console.WriteLine($"- Количество заказанного товара: {amount}");
        Console.WriteLine($"- Итоговая цена: {price}");
        Console.WriteLine($"- Дата заказа: {date}\n");
    }

    public void PrintMenu()
    {
        Console.WriteLine("1. Получить информацию о клиентах по названию товара.");
        Console.WriteLine("2. Определить \"Золотого\" клиента.");
        Console.WriteLine("3. Изменить ФИО контактного лица организации.\n");
        Console.WriteLine("Нажмите q, чтобы выйти...\n");
    }
}