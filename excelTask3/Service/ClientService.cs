using excelTask3.Interfaces;
using Models;

namespace excelTask3.Service
{
    public class ClientService : ModelService
    {
        public List<Client> ClientsList
        {
            get => AllClients;
            set
            {
                ClientsList = value;
                AllClients = ClientsList;
            }
        }
        public ClientService(IExcelProcess excelProcess) : base(excelProcess)
        {
        }
        public Client GetClientByOrganization(string name)
        {
            var clients = ClientsList;
            var result = clients.SingleOrDefault(x => x.Organization == name);
            return result;
        }

        public Client GetClientById(int id)
        {
            var clients = ClientsList;
            var result = clients.SingleOrDefault(x => x.Id == id);
            return result;
        }

        public void UpdateClientInfo(Client updatedClient)
        {
            var clientsIds = ClientsList.Select(x => x.Id).ToList();
            int idx = clientsIds.IndexOf(updatedClient.Id);
            ClientsList[idx] = updatedClient;
            UpdateAllClients();
        }
    }
}