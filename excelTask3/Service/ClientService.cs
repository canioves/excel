using excelTask3.Interfaces;
using Models;

namespace excelTask3.Service
{
    public class ClientService : ModelService
    {
        public List<Client> ClientsList { get { return AllClients; } }
        public ClientService(IExcelProcess excelProcess) : base(excelProcess)
        {
        }
        public Client GetClientByName(string name)
        {
            var clients = ClientsList;
            var result = clients.SingleOrDefault(x => x.FullName == name);
            return result;
        }

        public Client GetClientById(int id)
        {
            var clients = ClientsList;
            var result = clients.SingleOrDefault(x => x.Id == id);
            return result;
        }
    }
}