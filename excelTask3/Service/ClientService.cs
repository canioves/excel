using excelTask3.Interfaces;
using Models;

namespace excelTask3.Service
{
    public class ClientService : ModelService
    {
        public ClientService(IExcelProcess excelProcess) : base(excelProcess)
        {
        }
        public Client GetClientByName(string name)
        {
            var clients = AllClients;
            var result = clients.Single(x => x.FullName == name);
            return result;
        }

        public Client GetClientById(int id)
        {
            var clients = AllClients;
            var result = clients.Single(x => x.Id == id);
            return result;
        }
    }
}