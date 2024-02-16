using ClosedXML.Excel;
using excelTask3.Interfaces;
using Models;

namespace excelTask3.Service
{
    public class ModelService
    {
        protected readonly IExcelProcess _excelProcess;
        protected List<Product> AllProducts;
        protected List<Client> AllClients;
        protected List<Request> AllRequests;
        public ModelService(IExcelProcess excelProcess)
        {
            _excelProcess = excelProcess;
            AllProducts = GetAllProducts();
            AllClients = GetAllClients();
            AllRequests = GetAllRequsets();
        }

        public List<Product> GetAllProducts()
        {
            var table = _excelProcess.GetProductsTable();
            var rowsCount = table.RowCount();
            var products = new List<Product>();
            for (int i = 2; i <= rowsCount; i++)
            {
                var product = new Product
                {
                    Id = Convert.ToInt32(table.Field("Код товара").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString()),
                    Name = table.Field("Наименование").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString(),
                    UnitName = table.Field("Ед. измерения").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString(),
                    Price = Convert.ToDouble(table.Field("Цена товара за единицу").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString())
                };
                products.Add(product);
            }
            return products;
        }

        private List<Client> GetAllClients()
        {
            var table = _excelProcess.GetClientsTable();
            var rowsCount = table.RowCount();
            var clients = new List<Client>();
            for (int i = 2; i <= rowsCount; i++)
            {
                var client = new Client
                {
                    Id = Convert.ToInt32(table.Field("Код клиента").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString()),
                    Organization = table.Field("Наименование организации").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString(),
                    Address = table.Field("Адрес").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString(),
                    FullName = table.Field("Контактное лицо (ФИО)").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString()
                };
                clients.Add(client);
            }
            return clients;
        }

        private List<Request> GetAllRequsets()
        {
            var table = _excelProcess.GetRequestsTable();
            var rowsCount = table.RowCount();
            var requests = new List<Request>();
            for (int i = 2; i <= rowsCount; i++)
            {
                var request = new Request
                {
                    Id = Convert.ToInt32(table.Field("Код заявки").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString()),
                    ProductId = Convert.ToInt32(table.Field("Код товара").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString()),
                    ClientId = Convert.ToInt32(table.Field("Код клиента").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString()),
                    RequestNumber = Convert.ToInt32(table.Field("Номер заявки").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString()),
                    Amount = Convert.ToInt32(table.Field("Требуемое количество").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString()),
                    Date = Convert.ToDateTime(table.Field("Дата размещения").DataCells.Single(x => x.Address.RowNumber == i).Value.ToString())
                };
                requests.Add(request);
            }
            return requests;
        }

        public void UpdateAllClients()
        {
            var table = _excelProcess.GetClientsTable();
            var rowsCount = table.RowCount();
            for (int i = 2; i <= rowsCount; i++)
            {
                foreach (var fieldName in table.Fields)
                {
                    switch (fieldName.Name)
                    {
                        case "Код клиента":
                            table.Field(fieldName.Name).DataCells.Single(x => x.Address.RowNumber == i).Value = AllClients[i - 2].Id;
                            break;
                        case "Наименование организации":
                            table.Field(fieldName.Name).DataCells.Single(x => x.Address.RowNumber == i).Value = AllClients[i - 2].Organization;
                            break;
                        case "Адрес":
                            table.Field(fieldName.Name).DataCells.Single(x => x.Address.RowNumber == i).Value = AllClients[i - 2].Address;
                            break;
                        case "Контактное лицо (ФИО)":
                            table.Field(fieldName.Name).DataCells.Single(x => x.Address.RowNumber == i).Value = AllClients[i - 2].FullName;
                            break;
                    }
                }
            }
            _excelProcess.UpdateClientsTable(table);
        }
    }
}