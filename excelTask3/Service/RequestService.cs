using excelTask3.Interfaces;
using Models;

namespace excelTask3.Service
{
    public class RequestService : ModelService
    {
        public RequestService(IExcelProcess excelProcess) : base(excelProcess)
        {
        }

        public Request GetRequestByNumber(int number)
        {
            var requests = AllRequests;
            var result = requests.Single(x => x.RequestNumber == number);
            return result;
        }

        public Request GetRequestById(int id)
        {
            var requests = AllRequests;
            var result = requests.Single(x => x.Id == id);
            return result;
        }
    }
}