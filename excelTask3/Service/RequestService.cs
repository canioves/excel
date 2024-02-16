using excelTask3.Interfaces;

namespace excelTask3.Service
{
    public class RequestService : ModelService
    {
        public RequestService(IExcelProcess excelProcess) : base(excelProcess)
        {
        }
    }
}