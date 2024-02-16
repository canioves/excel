using excelTask3.Interfaces;
using Models;

namespace excelTask3.Service
{
    public class RequestService : ModelService
    {
        public List<Request> RequestsList { get { return AllRequests; } }
        public RequestService(IExcelProcess excelProcess) : base(excelProcess)
        {
        }

        public Request GetRequestByNumber(int number)
        {
            var requests = RequestsList;
            var result = requests.SingleOrDefault(x => x.RequestNumber == number);
            return result;
        }

        public Request GetRequestById(int id)
        {
            var requests = RequestsList;
            var result = requests.SingleOrDefault(x => x.Id == id);
            return result;
        }
    }
}