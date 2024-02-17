using System.Text;
using excelTask3.Exceptions;
using excelTask3.Service;

Console.InputEncoding = Encoding.UTF8;

App app = new App();

bool appRunning = true;
Console.WriteLine("Введите путь до Excel файла: \n");

while (app.repeatInit)
{
    // string path = @"Data\dataTable.xlsx";
    Console.Write("Путь => ");
    string path = Console.ReadLine();
    Console.Clear();
    app.InitExcel(path);
}

app.InitSheets();

Thread.Sleep(2000);

while (appRunning)
{
    Console.Clear();
    app.PrintMenu();
    var key = Console.ReadKey().Key;
    // var key = Console.ReadLine();
    bool keepAlive = true;
    Console.Clear();

    switch (key)
    {
        case ConsoleKey.D1:
        // case "1":
            while (keepAlive)
            {
                Console.Write("Введите название товара: ");
                string productName = Console.ReadLine();
                keepAlive = app.GetAllProductInfo(productName);
            }
            break;
        case ConsoleKey.D2:
        // case "2":
            while (keepAlive)
            {
                Console.Write("Введите год: ");
                int.TryParse(Console.ReadLine(), out int year);
                Console.Write("Введите месяц: ");
                int.TryParse(Console.ReadLine(), out int month);
                keepAlive = app.GetGoldenClient(year, month);
            }
            break;
        case ConsoleKey.D3:
        // case "3":
            while (keepAlive)
            {
                Console.Write("Введите название орагнизации: ");
                string organization = Console.ReadLine();
                Console.Write("Введите новое ФИО контактного лица: ");
                string newName = Console.ReadLine();
                keepAlive = app.ChangeOrgContactName(organization, newName);
            }
            break;
        case ConsoleKey.Q:
        // case "q":
            Console.WriteLine("Выход...");
            appRunning = false;
            break;
        default:
            break;
    }
    if (!appRunning) break;
}

public class App
{
    private ProductService productService;
    private ClientService clientService;
    private RequestService requestService;
    private ExcelProcess excelProcess;
    public bool repeatInit = true;

    public void InitExcel(string path)
    {
        excelProcess = new ExcelProcess(path);
        repeatInit = excelProcess.repeatInit;
    }

    public void InitSheets()
    {
        productService = new ProductService(excelProcess);
        clientService = new ClientService(excelProcess);
        requestService = new RequestService(excelProcess);
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

        return Navigation(startOver, waitKey);
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

        return Navigation(startOver, waitKey);
    }

    public bool GetGoldenClient(int year, int month)
    {
        bool startOver = true;
        bool waitKey = true;

        try
        {
            Dictionary<int, int> dc = new Dictionary<int, int>();
            if (year == 0) throw new ArgumentException("Год введен неверно!");
            else if (month == 0) throw new ArgumentException("Месяц введен неверно!");
            var clients = clientService.ClientsList;
            var requests = requestService.RequestsList;

            var filterRequests = requests.Where(x => x.Date.Year == year && x.Date.Month == month);
            var filterClients = filterRequests.Select(x => x.ClientId).Distinct();

            int max = 0;

            foreach (int clientId in filterClients)
            {
                int reqCount = filterRequests.Count(x => x.ClientId == clientId);
                if (reqCount > max) max = reqCount;
                dc[clientId] = reqCount;
            }

            var goldenClients = dc.Where(x => x.Value == max).Select(x => x.Key).ToList();

            if (goldenClients.Count == 1)
            {
                string goldenClient = clientService.GetClientById(goldenClients.First()).Organization;
                Console.WriteLine($"\nЗолотой клиент: {goldenClient}");
            }
            else
            {
                Console.WriteLine($"\nКлиентов с наибольшим количеством заказов найдено: {goldenClients.Count}");
                foreach (var targetId in goldenClients)
                {
                    var targetClients = clientService.GetClientById(targetId);
                    Console.WriteLine(targetClients.Organization);
                }
            }

            startOver = false;
        }
        catch (ArgumentException aEx)
        {
            Console.WriteLine(aEx.Message);
        }

        return Navigation(startOver, waitKey);
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
        Console.WriteLine("1 => Получить информацию о клиентах по названию товара.");
        Console.WriteLine("2 => Определить \"Золотого\" клиента.");
        Console.WriteLine("3 => Изменить ФИО контактного лица организации.\n");
        Console.WriteLine("Нажмите q, чтобы выйти...\n");
    }

    private bool Navigation(bool startOver, bool wait)
    {
        if (startOver)
        {
            wait = false;
            Console.WriteLine("Попробовать еще раз? (y/n)\n");
            var answer = Console.ReadKey().Key;
            startOver = answer == ConsoleKey.Y;
            Console.Clear();
        }

        if (wait)
        {
            Console.WriteLine("Для продолжения нажмите любую кнопку...");
            Console.ReadKey();
        }

        return startOver;
    }
}